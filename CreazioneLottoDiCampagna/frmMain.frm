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
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20910
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
   ScaleHeight     =   10965
   ScaleWidth      =   20910
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   66
      Top             =   10620
      Width           =   20910
      _ExtentX        =   36883
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   10620
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   20910
      _LayoutVersion  =   2
      _ExtentX        =   36883
      _ExtentY        =   18733
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
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
         Height          =   10545
         Left            =   0
         ScaleHeight     =   10515
         ScaleWidth      =   20805
         TabIndex        =   70
         Top             =   0
         Width           =   20835
         Begin VB.Frame fraAttesa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   6720
            TabIndex        =   217
            Top             =   3960
            Visible         =   0   'False
            Width           =   6735
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1815
               Left            =   240
               Picture         =   "frmMain.frx":479EA
               ScaleHeight     =   1815
               ScaleWidth      =   6375
               TabIndex        =   218
               Top             =   120
               Width           =   6375
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
                  Left            =   0
                  TabIndex        =   219
                  Top             =   1320
                  Width           =   8055
               End
            End
         End
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   10215
            Left            =   120
            ScaleHeight     =   10185
            ScaleWidth      =   20505
            TabIndex        =   71
            Top             =   120
            Width           =   20535
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   "PROVVISORIO"
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
               Left            =   18600
               TabIndex        =   235
               Top             =   840
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "ACQUISTO MERCE"
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
               Left            =   16320
               TabIndex        =   234
               Top             =   840
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox chkChiuso 
               Alignment       =   1  'Right Justify
               Caption         =   "LOTTO CHIUSO"
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
               Left            =   14160
               TabIndex        =   221
               Top             =   840
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.Frame Frame3 
               Caption         =   "Altri dati"
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
               Height          =   3495
               Left            =   6480
               TabIndex        =   102
               Top             =   1200
               Width           =   13935
               Begin VB.Frame FraOrdine 
                  Caption         =   "Riepilogo ordini"
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
                  Left            =   7680
                  TabIndex        =   210
                  Top             =   2520
                  Width           =   6135
                  Begin DMTEDITNUMLib.dmtNumber txtQtaOrdinata 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   211
                     Top             =   435
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
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQtaEvasa 
                     Height          =   315
                     Left            =   2040
                     TabIndex        =   213
                     Top             =   420
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
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
                  Begin DMTEDITNUMLib.dmtNumber txtQtaDaEvadere 
                     Height          =   315
                     Left            =   4080
                     TabIndex        =   215
                     Top             =   435
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
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label19 
                     Caption         =   "Da evadere"
                     Height          =   255
                     Index           =   2
                     Left            =   4080
                     TabIndex        =   216
                     Top             =   240
                     Width           =   1815
                  End
                  Begin VB.Label Label19 
                     Caption         =   "Evasa"
                     Height          =   255
                     Index           =   1
                     Left            =   2040
                     TabIndex        =   214
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.Label Label19 
                     Caption         =   "Ordinata"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   212
                     Top             =   240
                     Width           =   1575
                  End
               End
               Begin VB.CommandButton cmdUtente 
                  Height          =   325
                  Left            =   5640
                  Picture         =   "frmMain.frx":511A4
                  Style           =   1  'Graphical
                  TabIndex        =   171
                  ToolTipText     =   "Visualizza utenti"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.TextBox txtOraModifica 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   12480
                  TabIndex        =   169
                  Top             =   2160
                  Width           =   1335
               End
               Begin DMTDATETIMELib.dmtDate txtDataModifica 
                  Height          =   315
                  Left            =   10920
                  TabIndex        =   168
                  Top             =   2160
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin VB.TextBox txtNomePCMod 
                  Height          =   285
                  Left            =   10920
                  Locked          =   -1  'True
                  TabIndex        =   163
                  Top             =   1095
                  Width           =   2895
               End
               Begin VB.TextBox txtUtentePCMod 
                  Height          =   285
                  Left            =   10920
                  Locked          =   -1  'True
                  TabIndex        =   160
                  Top             =   660
                  Width           =   2895
               End
               Begin VB.TextBox txtOraInserimento 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   9240
                  TabIndex        =   158
                  Top             =   2160
                  Width           =   1335
               End
               Begin DMTDATETIMELib.dmtDate txtDataInserimento 
                  Height          =   315
                  Left            =   7680
                  TabIndex        =   157
                  Top             =   2160
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin VB.TextBox txtNomePCIns 
                  Height          =   285
                  Left            =   7680
                  Locked          =   -1  'True
                  TabIndex        =   152
                  Top             =   1095
                  Width           =   2895
               End
               Begin VB.TextBox txtUtentePCIns 
                  Height          =   285
                  Left            =   7680
                  Locked          =   -1  'True
                  TabIndex        =   149
                  Top             =   660
                  Width           =   2895
               End
               Begin VB.CommandButton cmdProtocollo 
                  Height          =   325
                  Left            =   5040
                  Picture         =   "frmMain.frx":5172E
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  ToolTipText     =   "Trova protocollo di certificazione"
                  Top             =   1080
                  Width           =   375
               End
               Begin VB.TextBox txtProtocolloCertificazione 
                  Height          =   325
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   4920
               End
               Begin VB.TextBox txtSuperficieHA 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1920
                  Locked          =   -1  'True
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.TextBox txtSuperficieHAEffettiva 
                  Enabled         =   0   'False
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
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   103
                  TabStop         =   0   'False
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtSuperficieMQ 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   30
                  Top             =   1680
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
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
               Begin DmtCodDescCtl.DmtCodDesc CDCertificazione 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   7335
                  _ExtentX        =   12938
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":51CB8
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":51D16
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":51D7C
                  ForeColor       =   8388608
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
               Begin DMTEDITNUMLib.dmtNumber txtSuperficieMQEffettiva 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   104
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   375
                  _Version        =   65536
                  _ExtentX        =   661
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
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
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDATETIMELib.dmtDate txtDataImportazione 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboMagazzinoCaricoLotto 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   33
                  Top             =   2280
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
               Begin DMTDataCmb.DMTCombo cboFunzioneCaricoLotto 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   34
                  Top             =   2280
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroPianteMQ 
                  Height          =   315
                  Left            =   3840
                  TabIndex        =   32
                  Top             =   1680
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
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
               End
               Begin DMTDataCmb.DMTCombo cboSchema 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   127
                  Top             =   2880
                  Width           =   5655
                  _ExtentX        =   9975
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
               Begin DMTDataCmb.DMTCombo cboUtenteDMTIns 
                  Height          =   315
                  Left            =   7680
                  TabIndex        =   154
                  Top             =   1620
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin DMTDataCmb.DMTCombo cboUtenteDMTMod 
                  Height          =   315
                  Left            =   10920
                  TabIndex        =   165
                  Top             =   1620
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   7560
                  X2              =   7560
                  Y1              =   240
                  Y2              =   3360
               End
               Begin VB.Label Label2 
                  Caption         =   "Ora "
                  Height          =   255
                  Index           =   25
                  Left            =   12480
                  TabIndex        =   170
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Data "
                  Height          =   255
                  Index           =   24
                  Left            =   10920
                  TabIndex        =   167
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Utente DMT Professional"
                  Height          =   255
                  Index           =   23
                  Left            =   10920
                  TabIndex        =   166
                  Top             =   1395
                  Width           =   2295
               End
               Begin VB.Label Label16 
                  Caption         =   "Nome P.C."
                  Height          =   255
                  Left            =   10920
                  TabIndex        =   164
                  Top             =   915
                  Width           =   2895
               End
               Begin VB.Label Label15 
                  Caption         =   "Utente P.C. "
                  Height          =   255
                  Left            =   10920
                  TabIndex        =   162
                  Top             =   480
                  Width           =   2895
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ULTIMA MODIFICA"
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
                  Left            =   10920
                  TabIndex        =   161
                  Top             =   195
                  Width           =   2895
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   10755
                  X2              =   10755
                  Y1              =   240
                  Y2              =   2400
               End
               Begin VB.Label Label2 
                  Caption         =   "Ora "
                  Height          =   255
                  Index           =   22
                  Left            =   9240
                  TabIndex        =   159
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Data "
                  Height          =   255
                  Index           =   21
                  Left            =   7680
                  TabIndex        =   156
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Utente DMT Professional"
                  Height          =   255
                  Index           =   20
                  Left            =   7680
                  TabIndex        =   155
                  Top             =   1395
                  Width           =   2295
               End
               Begin VB.Label Label13 
                  Caption         =   "Nome P.C."
                  Height          =   255
                  Left            =   7680
                  TabIndex        =   153
                  Top             =   915
                  Width           =   2895
               End
               Begin VB.Label Label12 
                  Caption         =   "Utente P.C. "
                  Height          =   255
                  Left            =   7680
                  TabIndex        =   151
                  Top             =   480
                  Width           =   2895
               End
               Begin VB.Label lblUtenteInserimento 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "INSERIMENTO"
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
                  Left            =   7680
                  TabIndex        =   150
                  Top             =   195
                  Width           =   2895
               End
               Begin VB.Label Label2 
                  Caption         =   "Schema del socio utilizzato"
                  Height          =   255
                  Index           =   18
                  Left            =   120
                  TabIndex        =   128
                  Top             =   2640
                  Width           =   4740
               End
               Begin VB.Label Label7 
                  Caption         =   "N° piante M.q."
                  Height          =   255
                  Index           =   1
                  Left            =   3840
                  TabIndex        =   111
                  Top             =   1440
                  Width           =   1890
               End
               Begin VB.Label Label2 
                  Caption         =   "Causale di carico lotto"
                  Height          =   255
                  Index           =   16
                  Left            =   3240
                  TabIndex        =   110
                  Top             =   2040
                  Width           =   2295
               End
               Begin VB.Label Label2 
                  Caption         =   "Magazzino di carico lotto"
                  Height          =   255
                  Index           =   15
                  Left            =   120
                  TabIndex        =   109
                  Top             =   2040
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Protocollo di certificazione"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   108
                  Top             =   840
                  Width           =   3255
               End
               Begin VB.Label Label1 
                  Caption         =   "Sup. in Ha"
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
                  Index           =   4
                  Left            =   1920
                  TabIndex        =   107
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Sup. in Mq"
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
                  Index           =   5
                  Left            =   120
                  TabIndex        =   106
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.Label Label7 
                  Caption         =   "Data Importazione"
                  Height          =   255
                  Index           =   0
                  Left            =   5520
                  TabIndex        =   105
                  Top             =   840
                  Width           =   1890
               End
            End
            Begin VB.CommandButton cmdLottoSeme 
               Height          =   325
               Left            =   12480
               Picture         =   "frmMain.frx":51DD6
               Style           =   1  'Graphical
               TabIndex        =   144
               ToolTipText     =   "Lotto del seme utilizzato"
               Top             =   1200
               Width           =   375
            End
            Begin VB.Frame Frame5 
               Caption         =   "Lotto del seme utilizzato"
               ForeColor       =   &H00FF0000&
               Height          =   1000
               Left            =   6480
               TabIndex        =   143
               Top             =   1200
               Visible         =   0   'False
               Width           =   6375
               Begin VB.TextBox txtDescrizioneLottoSeme 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   146
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   3975
               End
               Begin VB.TextBox txtCodiceLottoSeme 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   145
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   2175
               End
               Begin VB.Label Label11 
                  Caption         =   "Descrizione"
                  Height          =   255
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   148
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.Label Label11 
                  Caption         =   "Codice"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   147
                  Top             =   360
                  Width           =   2175
               End
            End
            Begin VB.TextBox txtCodicePeriodoCampagna 
               Height          =   285
               Left            =   3840
               TabIndex        =   125
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtAnnoRifPeriodoCampagna 
               Height          =   285
               Left            =   5040
               TabIndex        =   124
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtCodiceTipoProduzione 
               Height          =   285
               Left            =   840
               TabIndex        =   123
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtCodiceVarieta 
               Height          =   285
               Left            =   11760
               TabIndex        =   122
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtCodiceFamigliaProdotto 
               Height          =   285
               Left            =   5280
               TabIndex        =   121
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Frame FraLottoDiCampagna 
               Caption         =   "LOTTO DI PRODUZIONE"
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
               TabIndex        =   117
               Top             =   4560
               Width           =   20295
               Begin VB.TextBox Text1 
                  Height          =   315
                  Left            =   11760
                  Locked          =   -1  'True
                  TabIndex        =   237
                  Top             =   480
                  Width           =   3330
               End
               Begin VB.CheckBox chkAnnullato 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ANNULLATO"
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
                  Left            =   18600
                  TabIndex        =   236
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1575
               End
               Begin VB.TextBox TxtCodiceImEx 
                  Height          =   315
                  Left            =   15240
                  TabIndex        =   37
                  Top             =   480
                  Width           =   3330
               End
               Begin VB.TextBox txtCodiceLotto 
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
                  Left            =   120
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   35
                  Top             =   480
                  Width           =   5295
               End
               Begin VB.TextBox txtDescrizioneLotto 
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
                  Left            =   5520
                  MaxLength       =   40
                  TabIndex        =   36
                  Top             =   480
                  Width           =   6135
               End
               Begin VB.Label Label2 
                  Caption         =   "Codice lotto collegato"
                  Height          =   255
                  Index           =   36
                  Left            =   11760
                  TabIndex        =   238
                  Top             =   240
                  Width           =   2100
               End
               Begin VB.Label Label2 
                  Caption         =   "Codice Import / Export"
                  Height          =   255
                  Index           =   17
                  Left            =   15240
                  TabIndex        =   120
                  Top             =   240
                  Width           =   2100
               End
               Begin VB.Label Label2 
                  Caption         =   "Codice"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   119
                  Top             =   240
                  Width           =   5295
               End
               Begin VB.Label Label2 
                  Caption         =   "Descrizione "
                  Height          =   255
                  Index           =   2
                  Left            =   5520
                  TabIndex        =   118
                  Top             =   240
                  Width           =   6135
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Gestione date"
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
               Height          =   3375
               Left            =   120
               TabIndex        =   91
               Top             =   1200
               Width           =   6255
               Begin DMTDATETIMELib.dmtDate txtDataInizioProduzione 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   17
                  Top             =   2400
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataFineProduzione 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   18
                  Top             =   2760
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataSemina 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataFineProduzioneEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   25
                  Top             =   2760
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataInizioProduzioneEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   24
                  Top             =   2400
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataSeminaEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   19
                  Top             =   600
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataGerm 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   13
                  Top             =   960
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataNascita 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   14
                  Top             =   1320
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataTrapianto 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   15
                  Top             =   1680
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataRipic 
                  Height          =   315
                  Left            =   2160
                  TabIndex        =   16
                  Top             =   2040
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataGermEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   20
                  Top             =   960
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataNascitaEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   21
                  Top             =   1320
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataTrapiantoEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   22
                  Top             =   1680
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataRipicEffettiva 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   23
                  Top             =   2040
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di ripicchett."
                  Height          =   255
                  Index           =   14
                  Left            =   120
                  TabIndex        =   100
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di trapianto"
                  Height          =   255
                  Index           =   13
                  Left            =   120
                  TabIndex        =   99
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di nascita"
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   98
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di germinaz."
                  Height          =   255
                  Index           =   11
                  Left            =   120
                  TabIndex        =   97
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   "Presunta"
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
                  Left            =   2280
                  TabIndex        =   93
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  Caption         =   "Effettiva"
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
                  Left            =   4320
                  TabIndex        =   92
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di fine racc."
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   96
                  Top             =   2760
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di inizio racc."
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   95
                  Top             =   2400
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  Caption         =   "Data di semina"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   94
                  Top             =   645
                  Width           =   1695
               End
               Begin VB.Shape Shape1 
                  Height          =   2775
                  Left            =   2040
                  Top             =   360
                  Width           =   1935
               End
               Begin VB.Shape Shape2 
                  Height          =   2775
                  Left            =   4080
                  Top             =   360
                  Width           =   1935
               End
            End
            Begin VB.TextBox txtVarieta 
               Height          =   315
               Left            =   9765
               Locked          =   -1  'True
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   210
               Width           =   3735
            End
            Begin VB.CommandButton cmdSelezionaVarieta 
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
               Left            =   9400
               Picture         =   "frmMain.frx":52360
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Trova varietà prodotto"
               Top             =   210
               Width           =   375
            End
            Begin VB.CommandButton cmdProgressivo 
               Height          =   315
               Left            =   13560
               Picture         =   "frmMain.frx":528EA
               Style           =   1  'Graphical
               TabIndex        =   90
               ToolTipText     =   "Progressivo sblocco del lotto"
               Top             =   840
               Width           =   375
            End
            Begin DMTEDITNUMLib.dmtNumber txtProgressivo 
               Height          =   315
               Left            =   12000
               TabIndex        =   11
               Top             =   840
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   253
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
            End
            Begin VB.TextBox txtNomeSocio 
               BackColor       =   &H8000000F&
               Height          =   340
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   240
               Width           =   2295
            End
            Begin MSComDlg.CommonDialog CD 
               Left            =   120
               Top             =   2040
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin DMTDATETIMELib.dmtDate txtDataSbloccoLotto 
               Height          =   315
               Left            =   10320
               TabIndex        =   10
               Top             =   840
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTDataCmb.DMTCombo cboFamigliaProdotti 
               Height          =   315
               Left            =   6480
               TabIndex        =   1
               Top             =   240
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
            Begin DMTDataCmb.DMTCombo cboPeriodoCampagna 
               Height          =   315
               Left            =   14160
               TabIndex        =   3
               Top             =   210
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
            Begin TabDlg.SSTab SSTab1 
               Height          =   4455
               Left            =   120
               TabIndex        =   72
               Top             =   5520
               Width           =   20295
               _ExtentX        =   35798
               _ExtentY        =   7858
               _Version        =   393216
               Tabs            =   8
               TabsPerRow      =   8
               TabHeight       =   706
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Prodotti"
               TabPicture(0)   =   "frmMain.frx":52E74
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label1(1)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label1(0)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "lblArticolo"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "CDArticolo"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "txtQtaPresunta"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "txtCalibro"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "cmdElimina_Quadratura"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "cmdSalva_Quadratura"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "cmdNuovo_Quadratura"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "TxtArticolo"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "cmdElabora"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "FraModProdotti"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "Griglia"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).Control(13)=   "cmdVisUtenteMod"
               Tab(0).Control(13).Enabled=   0   'False
               Tab(0).ControlCount=   14
               TabCaption(1)   =   "Serre e Appezzamenti"
               TabPicture(1)   =   "frmMain.frx":52E90
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label2(9)"
               Tab(1).Control(1)=   "Label2(8)"
               Tab(1).Control(2)=   "lblSupMq"
               Tab(1).Control(3)=   "lblSupHA"
               Tab(1).Control(4)=   "lblFoglioCatastale"
               Tab(1).Control(5)=   "GrigliaTerreni"
               Tab(1).Control(6)=   "CDSerre"
               Tab(1).Control(7)=   "txtSuperficieMQ_Serra"
               Tab(1).Control(8)=   "txtSuperficieHA_Serra"
               Tab(1).Control(9)=   "cmdSerreDisponibili"
               Tab(1).Control(9).Enabled=   0   'False
               Tab(1).Control(10)=   "cmdNuovoTerreno"
               Tab(1).Control(11)=   "cmdSalvaTerreno"
               Tab(1).Control(12)=   "cmdEliminaTerreno"
               Tab(1).Control(12).Enabled=   0   'False
               Tab(1).ControlCount=   13
               TabCaption(2)   =   "Documentazione"
               TabPicture(2)   =   "frmMain.frx":52EAC
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "txtPercorsoSelezionato"
               Tab(2).Control(0).Enabled=   0   'False
               Tab(2).Control(1)=   "Frame4"
               Tab(2).Control(2)=   "Frame1"
               Tab(2).Control(3)=   "Label9"
               Tab(2).ControlCount=   4
               TabCaption(3)   =   "Altre informazioni"
               TabPicture(3)   =   "frmMain.frx":52EC8
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "FraDatiAcquistoPiantine"
               Tab(3).Control(1)=   "txtDescrizioneConcia"
               Tab(3).Control(2)=   "Label8"
               Tab(3).ControlCount=   3
               TabCaption(4)   =   "Semi utilizzati"
               TabPicture(4)   =   "frmMain.frx":52EE4
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "lblLottoSeme"
               Tab(4).Control(1)=   "Label2(19)"
               Tab(4).Control(2)=   "GrigliaSemi"
               Tab(4).Control(3)=   "txtQuantitaSemi"
               Tab(4).Control(4)=   "CDArticoloSemi"
               Tab(4).Control(5)=   "txtDescrizioneLottoSemeRiga"
               Tab(4).Control(5).Enabled=   0   'False
               Tab(4).Control(6)=   "txtCodiceLottoSemeRiga"
               Tab(4).Control(6).Enabled=   0   'False
               Tab(4).Control(7)=   "cmdNuovoSemi"
               Tab(4).Control(8)=   "cmdSalvaSemi"
               Tab(4).Control(9)=   "cmdEliminaSemi"
               Tab(4).Control(9).Enabled=   0   'False
               Tab(4).Control(10)=   "txtIDLottoArticolo"
               Tab(4).Control(11)=   "cmdLottoArticoloSeme"
               Tab(4).ControlCount=   12
               TabCaption(5)   =   "Ordini da cliente"
               TabPicture(5)   =   "frmMain.frx":52F00
               Tab(5).ControlEnabled=   0   'False
               Tab(5).Control(0)=   "cmdNuovoOrdiniAssociato"
               Tab(5).Control(1)=   "cmdEliminaOrdineAssociato"
               Tab(5).Control(1).Enabled=   0   'False
               Tab(5).Control(2)=   "GrigliaOrdini"
               Tab(5).ControlCount=   3
               TabCaption(6)   =   "Verifiche"
               TabPicture(6)   =   "frmMain.frx":52F1C
               Tab(6).ControlEnabled=   0   'False
               Tab(6).Control(0)=   "txtOraVerifica"
               Tab(6).Control(1)=   "txtAnnotazioniVerifica"
               Tab(6).Control(2)=   "txtTemperaturaVerifica"
               Tab(6).Control(3)=   "txtTipoRilevazione"
               Tab(6).Control(4)=   "txtOperatoreVerifica"
               Tab(6).Control(5)=   "txtDataVerifica"
               Tab(6).Control(6)=   "cmdEliminaVerifica"
               Tab(6).Control(6).Enabled=   0   'False
               Tab(6).Control(7)=   "cmdSalvaVerifica"
               Tab(6).Control(8)=   "cmdNuovoVerifica"
               Tab(6).Control(9)=   "GrigliaVerifica"
               Tab(6).Control(10)=   "Label1(11)"
               Tab(6).Control(11)=   "Label1(10)"
               Tab(6).Control(12)=   "Label1(9)"
               Tab(6).Control(13)=   "Label1(8)"
               Tab(6).Control(14)=   "Label1(7)"
               Tab(6).Control(15)=   "Label1(6)"
               Tab(6).ControlCount=   16
               TabCaption(7)   =   "Gestione sfalci"
               TabPicture(7)   =   "frmMain.frx":52F38
               Tab(7).ControlEnabled=   0   'False
               Tab(7).Control(0)=   "cmdEliminaSfalcio"
               Tab(7).Control(0).Enabled=   0   'False
               Tab(7).Control(1)=   "cmdSalvaSfalcio"
               Tab(7).Control(2)=   "cmdNuovoSfalcio"
               Tab(7).Control(3)=   "GrigliaSfalcio"
               Tab(7).Control(4)=   "cboSfalcio"
               Tab(7).Control(5)=   "txtDataPresSfalcio"
               Tab(7).Control(6)=   "txtDataEffSfalcio"
               Tab(7).Control(7)=   "Label2(33)"
               Tab(7).Control(8)=   "Label2(32)"
               Tab(7).Control(9)=   "Label1(12)"
               Tab(7).ControlCount=   10
               Begin VB.CommandButton cmdEliminaSfalcio 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   225
                  TabStop         =   0   'False
                  Top             =   3240
                  Width           =   2055
               End
               Begin VB.CommandButton cmdSalvaSfalcio 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   224
                  Top             =   2400
                  Width           =   2055
               End
               Begin VB.CommandButton cmdNuovoSfalcio 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   223
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.CommandButton cmdVisUtenteMod 
                  Caption         =   "Altri dettagli"
                  Height          =   375
                  Left            =   18000
                  TabIndex        =   209
                  Top             =   960
                  Width           =   2055
               End
               Begin DmtGridCtl.DmtGrid Griglia 
                  Height          =   3135
                  Left            =   120
                  TabIndex        =   46
                  Top             =   1200
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   5530
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
               Begin VB.Frame FraModProdotti 
                  Caption         =   "Utente inserimento e ultima modifica"
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
                  Left            =   120
                  TabIndex        =   188
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   11415
                  Begin VB.TextBox txtUtentePCInsArt 
                     Height          =   315
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   196
                     Top             =   470
                     Width           =   2895
                  End
                  Begin VB.TextBox txtPCInsArt 
                     Height          =   315
                     Left            =   3120
                     Locked          =   -1  'True
                     TabIndex        =   195
                     Top             =   470
                     Width           =   2895
                  End
                  Begin VB.TextBox txtOraInsArt 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   10080
                     TabIndex        =   193
                     Top             =   480
                     Width           =   1215
                  End
                  Begin VB.TextBox txtUtentePCModArt 
                     Height          =   315
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   192
                     Top             =   960
                     Width           =   2895
                  End
                  Begin VB.TextBox txtPCModArt 
                     Height          =   315
                     Left            =   3120
                     Locked          =   -1  'True
                     TabIndex        =   191
                     Top             =   960
                     Width           =   2895
                  End
                  Begin VB.TextBox txtOraModArt 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   10080
                     TabIndex        =   189
                     Top             =   960
                     Width           =   1215
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataModArt 
                     Height          =   315
                     Left            =   8520
                     TabIndex        =   190
                     Top             =   960
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataInsArt 
                     Height          =   315
                     Left            =   8520
                     TabIndex        =   194
                     Top             =   465
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboUtenteDMTInsArt 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   197
                     Top             =   465
                     Width           =   2295
                     _ExtentX        =   4048
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
                  Begin DMTDataCmb.DMTCombo cboUtenteDMTModArt 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   198
                     Top             =   960
                     Width           =   2295
                     _ExtentX        =   4048
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
                  Begin VB.Label Label2 
                     Caption         =   "Ora "
                     Height          =   255
                     Index           =   31
                     Left            =   10080
                     TabIndex        =   208
                     Top             =   720
                     Width           =   975
                  End
                  Begin VB.Label Label21 
                     Caption         =   "Utente P.C. inserimento"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   207
                     Top             =   285
                     Width           =   2895
                  End
                  Begin VB.Label Label20 
                     Caption         =   "Nome P.C. inserimento"
                     Height          =   255
                     Left            =   3120
                     TabIndex        =   206
                     Top             =   255
                     Width           =   2895
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Utente DMT Professional"
                     Height          =   255
                     Index           =   30
                     Left            =   6120
                     TabIndex        =   205
                     Top             =   255
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Data "
                     Height          =   255
                     Index           =   29
                     Left            =   8520
                     TabIndex        =   204
                     Top             =   255
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Ora "
                     Height          =   255
                     Index           =   28
                     Left            =   10080
                     TabIndex        =   203
                     Top             =   255
                     Width           =   975
                  End
                  Begin VB.Label Label18 
                     Caption         =   "Utente P.C. ultima modifica"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   202
                     Top             =   765
                     Width           =   2895
                  End
                  Begin VB.Label Label17 
                     Caption         =   "Nome P.C. ultima modifica"
                     Height          =   255
                     Left            =   3120
                     TabIndex        =   201
                     Top             =   765
                     Width           =   2895
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Utente DMT Professional"
                     Height          =   255
                     Index           =   27
                     Left            =   6120
                     TabIndex        =   200
                     Top             =   765
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Data "
                     Height          =   255
                     Index           =   26
                     Left            =   8520
                     TabIndex        =   199
                     Top             =   765
                     Width           =   1335
                  End
               End
               Begin DMTDATETIMELib.dmtTime txtOraVerifica 
                  Height          =   315
                  Left            =   -73560
                  TabIndex        =   183
                  Top             =   840
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.TextBox txtAnnotazioniVerifica 
                  Height          =   555
                  Left            =   -74880
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   180
                  Top             =   1440
                  Width           =   12015
               End
               Begin VB.TextBox txtTemperaturaVerifica 
                  Height          =   315
                  Left            =   -64680
                  TabIndex        =   179
                  Top             =   840
                  Width           =   1815
               End
               Begin VB.TextBox txtTipoRilevazione 
                  Height          =   315
                  Left            =   -68880
                  TabIndex        =   178
                  Top             =   840
                  Width           =   4095
               End
               Begin VB.TextBox txtOperatoreVerifica 
                  Height          =   315
                  Left            =   -72480
                  TabIndex        =   177
                  Top             =   840
                  Width           =   3495
               End
               Begin DMTDATETIMELib.dmtDate txtDataVerifica 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   176
                  Top             =   840
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.CommandButton cmdEliminaVerifica 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -56280
                  TabIndex        =   174
                  TabStop         =   0   'False
                  Top             =   3600
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaVerifica 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -56280
                  TabIndex        =   173
                  Top             =   2880
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoVerifica 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -56280
                  TabIndex        =   172
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoOrdiniAssociato 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -56280
                  TabIndex        =   141
                  Top             =   1680
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEliminaOrdineAssociato 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -56280
                  TabIndex        =   140
                  TabStop         =   0   'False
                  Top             =   2640
                  Width           =   1455
               End
               Begin VB.CommandButton cmdLottoArticoloSeme 
                  Height          =   325
                  Left            =   -64560
                  Picture         =   "frmMain.frx":52F54
                  Style           =   1  'Graphical
                  TabIndex        =   134
                  ToolTipText     =   "Trova lotto articolo"
                  Top             =   1320
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDLottoArticolo 
                  Height          =   255
                  Left            =   -65640
                  TabIndex        =   138
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   450
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.CommandButton cmdEliminaSemi 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -56400
                  TabIndex        =   137
                  TabStop         =   0   'False
                  Top             =   3720
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaSemi 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -56400
                  TabIndex        =   135
                  Top             =   2880
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoSemi 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -56400
                  TabIndex        =   136
                  Top             =   2040
                  Width           =   1455
               End
               Begin VB.TextBox txtCodiceLottoSemeRiga 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   132
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   3255
               End
               Begin VB.TextBox txtDescrizioneLottoSemeRiga 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -71640
                  TabIndex        =   133
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   7095
               End
               Begin VB.Frame FraDatiAcquistoPiantine 
                  Caption         =   "Dati di acquisto"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1095
                  Left            =   -74880
                  TabIndex        =   115
                  Top             =   2640
                  Width           =   12135
                  Begin VB.TextBox txtNPassaportoProduttore 
                     Height          =   325
                     Left            =   6480
                     TabIndex        =   65
                     Top             =   480
                     Width           =   4215
                  End
                  Begin VB.TextBox txtNomeFornitore 
                     BackColor       =   &H8000000F&
                     Height          =   340
                     Left            =   4080
                     Locked          =   -1  'True
                     TabIndex        =   64
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   2295
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDProduttore 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   63
                     Top             =   240
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":534DE
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":5352C
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":53581
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
                  Begin VB.Label Label10 
                     Caption         =   "Numero passaporto"
                     Height          =   255
                     Left            =   6480
                     TabIndex        =   116
                     Top             =   240
                     Width           =   4215
                  End
               End
               Begin VB.TextBox txtDescrizioneConcia 
                  Height          =   1815
                  Left            =   -74880
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   62
                  Top             =   720
                  Width           =   12135
               End
               Begin VB.TextBox txtPercorsoSelezionato 
                  ForeColor       =   &H00C00000&
                  Height          =   285
                  Left            =   -74880
                  Locked          =   -1  'True
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   12255
               End
               Begin VB.Frame Frame4 
                  Caption         =   "File nella cartella selezionata"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3015
                  Left            =   -68040
                  TabIndex        =   112
                  Top             =   1140
                  Width           =   5415
                  Begin VB.FileListBox FileDocumentazione 
                     Appearance      =   0  'Flat
                     Height          =   2565
                     Left            =   120
                     TabIndex        =   61
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   5175
                  End
               End
               Begin VB.CommandButton cmdElabora 
                  Caption         =   "Elaborazione articoli"
                  Height          =   375
                  Left            =   18000
                  TabIndex        =   45
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Percorso Documentazione "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3015
                  Left            =   -74880
                  TabIndex        =   101
                  Top             =   1140
                  Width           =   6735
                  Begin VB.CommandButton cmdNuovaCartella 
                     Height          =   285
                     Left            =   6240
                     Picture         =   "frmMain.frx":535DB
                     Style           =   1  'Graphical
                     TabIndex        =   59
                     TabStop         =   0   'False
                     ToolTipText     =   "Nuova cartella"
                     Top             =   960
                     Width           =   375
                  End
                  Begin VB.DirListBox DirSelezionato 
                     Appearance      =   0  'Flat
                     Height          =   2340
                     Left            =   120
                     TabIndex        =   60
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   5895
                  End
                  Begin VB.TextBox txtPercorsoDocumentazione 
                     Appearance      =   0  'Flat
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   120
                     TabIndex        =   56
                     Top             =   240
                     Width           =   5895
                  End
                  Begin VB.CommandButton cmdTrovaPercorso 
                     Height          =   285
                     Left            =   6240
                     Picture         =   "frmMain.frx":53B65
                     Style           =   1  'Graphical
                     TabIndex        =   57
                     TabStop         =   0   'False
                     ToolTipText     =   "Percorso"
                     Top             =   240
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdRipristina 
                     Height          =   285
                     Left            =   6240
                     Picture         =   "frmMain.frx":540EF
                     Style           =   1  'Graphical
                     TabIndex        =   58
                     TabStop         =   0   'False
                     ToolTipText     =   "Ripristina"
                     Top             =   600
                     Width           =   375
                  End
                  Begin VB.Line Line1 
                     X1              =   6120
                     X2              =   6120
                     Y1              =   240
                     Y2              =   2880
                  End
               End
               Begin VB.TextBox TxtArticolo 
                  Height          =   330
                  Left            =   1920
                  TabIndex        =   39
                  Top             =   855
                  Width           =   3615
               End
               Begin VB.CommandButton cmdNuovo_Quadratura 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   18000
                  TabIndex        =   43
                  Top             =   1920
                  Width           =   2055
               End
               Begin VB.CommandButton cmdSalva_Quadratura 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   18000
                  TabIndex        =   42
                  Top             =   2760
                  Width           =   2055
               End
               Begin VB.CommandButton cmdElimina_Quadratura 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   18000
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   3600
                  Width           =   2055
               End
               Begin VB.TextBox txtCalibro 
                  Height          =   330
                  Left            =   5640
                  TabIndex        =   40
                  Top             =   855
                  Width           =   1095
               End
               Begin VB.CommandButton cmdEliminaTerreno 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -56520
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   3720
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaTerreno 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -56520
                  TabIndex        =   50
                  Top             =   2880
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoTerreno 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -56520
                  TabIndex        =   51
                  Top             =   2040
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSerreDisponibili 
                  Caption         =   "SERRE E APPEZZAMENTI DISPONIBILI"
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
                  Left            =   -56760
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.TextBox txtSuperficieHA_Serra 
                  Height          =   330
                  Left            =   -68400
                  TabIndex        =   49
                  Top             =   840
                  Width           =   1815
               End
               Begin DMTEDITNUMLib.dmtNumber txtSuperficieMQ_Serra 
                  Height          =   330
                  Left            =   -70080
                  TabIndex        =   48
                  Top             =   840
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DmtCodDescCtl.DmtCodDesc CDSerre 
                  Height          =   615
                  Left            =   -74880
                  TabIndex        =   47
                  Top             =   600
                  Width           =   4695
                  _ExtentX        =   8281
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":54679
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":546C9
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":54720
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
               Begin DMTEDITNUMLib.dmtNumber txtQtaPresunta 
                  Height          =   330
                  Left            =   6840
                  TabIndex        =   41
                  Top             =   855
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   582
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   38
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":5477A
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":547C9
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":54820
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
               Begin DmtGridCtl.DmtGrid GrigliaTerreni 
                  Height          =   3015
                  Left            =   -74880
                  TabIndex        =   54
                  Top             =   1320
                  Width           =   18015
                  _ExtentX        =   31776
                  _ExtentY        =   5318
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
               Begin DmtCodDescCtl.DmtCodDesc CDArticoloSemi 
                  Height          =   615
                  Left            =   -74880
                  TabIndex        =   129
                  Top             =   480
                  Width           =   9015
                  _ExtentX        =   15901
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":5487A
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":548CA
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":54921
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
               Begin DMTEDITNUMLib.dmtNumber txtQuantitaSemi 
                  Height          =   330
                  Left            =   -65760
                  TabIndex        =   130
                  Top             =   720
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DmtGridCtl.DmtGrid GrigliaSemi 
                  Height          =   2535
                  Left            =   -74880
                  TabIndex        =   139
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   18375
                  _ExtentX        =   32411
                  _ExtentY        =   4471
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
               Begin DmtGridCtl.DmtGrid GrigliaOrdini 
                  Height          =   3735
                  Left            =   -74880
                  TabIndex        =   142
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   18495
                  _ExtentX        =   32623
                  _ExtentY        =   6588
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
               Begin DmtGridCtl.DmtGrid GrigliaVerifica 
                  Height          =   2295
                  Left            =   -74880
                  TabIndex        =   175
                  Top             =   2040
                  Width           =   18495
                  _ExtentX        =   32623
                  _ExtentY        =   4048
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
               Begin DmtGridCtl.DmtGrid GrigliaSfalcio 
                  Height          =   3015
                  Left            =   -74880
                  TabIndex        =   222
                  Top             =   1320
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   5318
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
               Begin DMTDataCmb.DMTCombo cboSfalcio 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   226
                  Top             =   840
                  Width           =   3975
                  _ExtentX        =   7011
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
               Begin DMTDATETIMELib.dmtDate txtDataPresSfalcio 
                  Height          =   315
                  Left            =   -70800
                  TabIndex        =   228
                  Top             =   840
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataEffSfalcio 
                  Height          =   315
                  Left            =   -68760
                  TabIndex        =   229
                  Top             =   840
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.Label Label2 
                  Caption         =   "Data effettiva"
                  Height          =   255
                  Index           =   33
                  Left            =   -68760
                  TabIndex        =   231
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "Data presunta"
                  Height          =   255
                  Index           =   32
                  Left            =   -70800
                  TabIndex        =   230
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Sfalcio"
                  Height          =   255
                  Index           =   12
                  Left            =   -74880
                  TabIndex        =   227
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Caption         =   "Annotazioni"
                  Height          =   255
                  Index           =   11
                  Left            =   -74880
                  TabIndex        =   187
                  Top             =   1200
                  Width           =   4335
               End
               Begin VB.Label Label1 
                  Caption         =   "Temperatura"
                  Height          =   255
                  Index           =   10
                  Left            =   -64680
                  TabIndex        =   186
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo rilevazione"
                  Height          =   255
                  Index           =   9
                  Left            =   -68880
                  TabIndex        =   185
                  Top             =   600
                  Width           =   3495
               End
               Begin VB.Label Label1 
                  Caption         =   "Operatore"
                  Height          =   255
                  Index           =   8
                  Left            =   -72480
                  TabIndex        =   184
                  Top             =   600
                  Width           =   3495
               End
               Begin VB.Label Label1 
                  Caption         =   "Ora"
                  Height          =   255
                  Index           =   7
                  Left            =   -73560
                  TabIndex        =   182
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Data"
                  Height          =   255
                  Index           =   6
                  Left            =   -74880
                  TabIndex        =   181
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Quantità"
                  Height          =   255
                  Index           =   19
                  Left            =   -65880
                  TabIndex        =   131
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label lblLottoSeme 
                  Caption         =   "Codice lotto del seme"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   126
                  Top             =   1080
                  Width           =   4695
               End
               Begin VB.Label Label9 
                  Caption         =   "Indirizzo "
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   114
                  Top             =   480
                  Width           =   11295
               End
               Begin VB.Label Label8 
                  Caption         =   "Annotazioni"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   113
                  Top             =   480
                  Width           =   12135
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
                  Height          =   270
                  Left            =   1920
                  TabIndex        =   80
                  Top             =   630
                  Width           =   3615
               End
               Begin VB.Label Label1 
                  Caption         =   "Calibro"
                  Height          =   255
                  Index           =   0
                  Left            =   5640
                  TabIndex        =   79
                  Top             =   645
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Caption         =   "Quantità presunta"
                  Height          =   255
                  Index           =   1
                  Left            =   6840
                  TabIndex        =   78
                  Top             =   645
                  Width           =   1935
               End
               Begin VB.Label lblFoglioCatastale 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   -70200
                  TabIndex        =   77
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.Label lblSupHA 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   -70200
                  TabIndex        =   76
                  Top             =   2160
                  Width           =   1215
               End
               Begin VB.Label lblSupMq 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   -70200
                  TabIndex        =   75
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Superficie in Mq"
                  Height          =   255
                  Index           =   8
                  Left            =   -70080
                  TabIndex        =   74
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.Label Label2 
                  Caption         =   "Superficie in HA"
                  Height          =   255
                  Index           =   9
                  Left            =   -68400
                  TabIndex        =   73
                  Top             =   600
                  Width           =   1455
               End
            End
            Begin DMTDataCmb.DMTCombo cboStatoLotto 
               Height          =   315
               Left            =   7680
               TabIndex        =   9
               Top             =   840
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
            Begin DMTDataCmb.DMTCombo cboTipoProduzione 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   840
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
            Begin DmtCodDescCtl.DmtCodDesc CDSocio 
               Height          =   615
               Left            =   120
               TabIndex        =   0
               Top             =   0
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1085
               PropCodice      =   $"frmMain.frx":5497B
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmMain.frx":549C9
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmMain.frx":54A19
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
            Begin DMTEDITNUMLib.dmtNumber txtIDVarieta 
               Height          =   330
               Left            =   13560
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   210
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
               _ExtentY        =   582
               _StockProps     =   253
               Text            =   "0"
               BackColor       =   16777215
               Enabled         =   0   'False
               Appearance      =   1
               AllowEmpty      =   0   'False
            End
            Begin DMTDataCmb.DMTCombo cboClass1 
               Height          =   315
               Left            =   2880
               TabIndex        =   7
               Top             =   840
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
            Begin DMTDataCmb.DMTCombo cboClass2 
               Height          =   315
               Left            =   5280
               TabIndex        =   8
               Top             =   840
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
            Begin VB.Label Label2 
               Caption         =   "Classificazione 2"
               Height          =   255
               Index           =   35
               Left            =   5280
               TabIndex        =   233
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label2 
               Caption         =   "Classificazione 1"
               Height          =   255
               Index           =   34
               Left            =   2880
               TabIndex        =   232
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label6 
               Caption         =   "Prog. sblocco"
               Height          =   255
               Left            =   12000
               TabIndex        =   89
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo produzione"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   87
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "Data sblocco"
               Height          =   255
               Index           =   3
               Left            =   10320
               TabIndex        =   86
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label2 
               Caption         =   "Varietà"
               Height          =   255
               Index           =   0
               Left            =   9765
               TabIndex        =   84
               Top             =   0
               Width           =   2970
            End
            Begin VB.Label Label2 
               Caption         =   "Periodo di campagna"
               Height          =   255
               Index           =   3
               Left            =   14160
               TabIndex        =   83
               Top             =   0
               Width           =   5295
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo coltura (famiglia prodotti)"
               Height          =   255
               Index           =   7
               Left            =   6480
               TabIndex        =   82
               Top             =   0
               Width           =   2775
            End
            Begin VB.Label Label1 
               Caption         =   "Stato del lotto"
               Height          =   255
               Index           =   2
               Left            =   7680
               TabIndex        =   81
               Top             =   600
               Width           =   3135
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   735
            Left            =   0
            TabIndex        =   220
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
      Begin DmtPrnDlgCtl.DMTDialog DmtPrnDlg 
         Left            =   480
         Top             =   1290
         _ExtentX        =   661
         _ExtentY        =   661
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
         Left            =   960
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   0
         TabIndex        =   68
         Top             =   2160
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   7875
         Left            =   0
         TabIndex        =   85
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
         Left            =   945
         Top             =   660
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   2580
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   60
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   6720
         Y1              =   3360
         Y2              =   3360
      End
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
'Il filtro attivo
Private m_ActiveFilter As DmtDocManLib.Filter
'La vista tabellare attiva
Private m_ActiveTableView As DmtDocManLib.TableView
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
'ADEGUAMENTI
Private WithEvents m_DocumentsLink3 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink3.VB_VarHelpID = -1
'SFALCI
Private WithEvents m_DocumentsLink4 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink4.VB_VarHelpID = -1
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
Private Const CAMPO_PER_CAPTION = "CodiceLotto"


'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut

'****************************VARIABILI CONTRATTO**********************************
Private NuovoDocumento As Long

Private AggiornamentoDocumento As Integer


Private ArrayDelete(150) As Long
Private ContDelete As Long
Private AggiornamentoGriglia As Integer
'******************************************************************************

Private Mov As DmtMovim.cMovimentazione
Private rsGrigliaOrd As ADODB.Recordset





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
        
        BarMenu.Bands("Standard").Tools("Rinnova").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
        
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
        Case "Rinnova"
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
    BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Enabled = False
    
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
            BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Enabled = False
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
            BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Enabled = False
            
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
            
            If BrwMain.GuiMode = dgNormal Then
                BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Enabled = True
            End If
            
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
Dim Link_Socio_Inserito As Boolean
    PermissionToSave = False

    If Me.CDSocio.KeyFieldID = 0 Then
        MsgBox "Manca il socio", vbCritical, "Controllo inserimento dati"
        Me.CDSocio.SetFocus
    Exit Function
    End If
    


    If Me.txtIDVarieta.Value <= 0 Then
        MsgBox "Manca la varietà del lotto", vbCritical, "Controllo inserimento dati"
        Me.txtVarieta.SetFocus
    Exit Function
    End If

    'If Len(Me.txtCodiceLotto.Text) <= 0 Then
    '    MsgBox "Manca il codice del lotto", vbCritical, "Inserimento dati"
    '    Me.txtCodiceLotto.SetFocus
    'Exit Function
    'End If

    'If Len(Me.txtDescrizioneLotto.Text) <= 0 Then
    '    MsgBox "Manca la descrizione del lotto", vbCritical, "Inserimento dati"
    '    Me.txtDescrizioneLotto.SetFocus
    'Exit Function
    'End If

    If Me.cboPeriodoCampagna.CurrentID <= 0 Then
        MsgBox "Manca il riferimento del periodo di campagna", vbCritical, "Controllo inserimento dati"
        Me.cboPeriodoCampagna.SetFocus
    Exit Function
    End If

    If Len(Me.txtDataSemina.Text) <= 0 Then
        MsgBox "Manca la data di semina del lotto", vbCritical, "Controllo inserimento dati"
        Me.txtDataSemina.SetFocus
    Exit Function
    End If

    If Len(Me.txtDataInizioProduzione.Text) <= 0 Then
        MsgBox "Manca la data di inizio raccolta", vbCritical, "Controllo inserimento dati"
        Me.txtDataInizioProduzione.SetFocus
    Exit Function
    End If

    If Len(Me.txtDataFineProduzione.Text) <= 0 Then
        MsgBox "Manca la data di fine raccolta", vbCritical, "Controllo inserimento dati"
        Me.txtDataFineProduzione.SetFocus
    Exit Function
    End If
    
    If m_Document(m_Document.PrimaryKey).Value > 0 Then
        If ControlloEsistenzaCodiceLotto(Me.txtCodiceLotto.Text, m_Document(m_Document.PrimaryKey).Value) = True Then
            MsgBox "Il codice lotto è esistente", vbCritical, "Controllo inserimento dati"
        Exit Function
        End If
'    Else
'        If StringaLottoStd <> Me.txtCodiceLotto.Text Then
'            If GET_TIPO_CODICE < 2 Then
'                If ControlloEsistenzaCodiceLotto(Me.txtCodiceLotto.Text, m_Document(m_Document.PrimaryKey).Value) = True Then
'                    MsgBox "Il codice lotto è esistente", vbCritical, "Inserimento dati"
'                    Exit Function
'                End If
'            End If
'        End If
    End If
        
    If Me.cboStatoLotto.CurrentID = 0 Then
        MsgBox "Manca lo stato del lotto", vbCritical, "Controllo inserimento dati"
        Me.cboStatoLotto.SetFocus
        Exit Function
    End If
    
    If Me.txtProgressivo.Value > 0 Then
        cmdProgressivo_Click
    End If

    If Me.CDCertificazione.KeyFieldID > 0 Then
        If Len(Trim(Me.txtProtocolloCertificazione.Text)) = 0 Then
            MsgBox "Manca il protocollo della certificazione", vbCritical, "Controllo inserimento dati"
            Me.cmdProtocollo.SetFocus
            Exit Function
        End If
    End If

    If LINK_TIPO_GESTIONE_MOVIMENTI = 1 Then
        If Me.cboMagazzinoCaricoLotto.CurrentID = 0 Then
            MsgBox "Inserire il magazzino di carico del lotto", vbCritical, "Controllo inserimento dati"
            'Me.cboMagazzinoCaricoLotto.SetFocus
            Exit Function
        End If
        If Me.cboFunzioneCaricoLotto.CurrentID = 0 Then
            MsgBox "Inserire la causale di magazzino per il carico del lotto", vbCritical, "Controllo inserimento dati"
            Me.cboFunzioneCaricoLotto.SetFocus
            Exit Function
        End If
    End If

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
    
    Me.chkChiuso.Value = vbChecked
    Me.cboMagazzinoCaricoLotto.WriteOn GET_LINK_MAGAZZINO_PARAMETRI
    Me.cboFunzioneCaricoLotto.WriteOn GET_LINK_CAUSALE_PARAMETRI
        
    Me.cboUtenteDMTIns.WriteOn TheApp.IDUser
    Me.txtNomePCIns.Text = GET_NOMECOMPUTER
    Me.txtUtentePCIns.Text = GET_NOMEUTENTE
    Me.txtDataInserimento.Value = Date
    Me.txtOraInserimento.Text = GET_ORARIO(Now)
    
    GET_PARAMETRI_PROGRESSIVO_LOTTO TheApp.Branch
    
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
    
    'TipoAttivazione = TipoAttivazioneLicenza
    'Select Case TipoAttivazione
    '    Case 1
    '
    '    Case 2
    '        NumeroRecordTabella = GetConteggioRecord("RV_PO01_LottoCampagna", "IDAzienda")
    '    Case 0
    '        FrmLicenza.Show vbModal
    '        Unload frmMain
    'End Select
   
        
    
    
End Sub
Private Function GetConteggioRecord(NomeTabella As String, CampoTabella As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(" & CampoTabella & ") AS Conteggio "
sSQL = sSQL & "FROM  " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm

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
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0
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
        Case "Mnu_Rinnova_lotti"
            
            GET_DATI_PER_RINNOVO
            'frmRinnovo.Show vbModal
            
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
        Case "cmdAvviaListaQualita"
             cmdGestioneQualitaLista
        Case "cmdAvviaQualita"
             cmdGestioneQualita
        Case "cmdSalvaComeNuovo"
             cmdSalvaComeNuovo_Click
        Case "cmdGestioneSfalci"
            
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
    
    
    'IDSocio
    Set Field = New FormField
    Set Field.Control = Me.CDSocio
    Field.Name = "IDSocio"
    Field.Visible = True
    Me.CDSocio.Tag = Field.Name
    m_FormFields.Add Field


    'IDRV_PO01_Varieta
    Set Field = New FormField
    Set Field.Control = Me.txtIDVarieta
    Field.Name = "IDRV_PO01_Varieta"
    Field.Visible = True
    Me.txtIDVarieta.Tag = Field.Name
    m_FormFields.Add Field

    'IDRV_PO01_PeriodoCampagna
    Set Field = New FormField
    Set Field.Control = Me.cboPeriodoCampagna
    Field.Name = "IDRV_PO01_PeriodoCampagna"
    Field.Visible = True
    Me.cboPeriodoCampagna.Tag = Field.Name
    m_FormFields.Add Field

    'CodiceLotto
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceLotto
    Field.Name = "CodiceLotto"
    Field.Visible = True
    Me.txtCodiceLotto.Tag = Field.Name
    m_FormFields.Add Field

    'DescrizioneLotto
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneLotto
    Field.Name = "DescrizioneLotto"
    Field.Visible = True
    Me.txtDescrizioneLotto.Tag = Field.Name
    m_FormFields.Add Field

    'Data semina
    Set Field = New FormField
    Set Field.Control = Me.txtDataSemina
    Field.Name = "DataSemina"
    Field.Visible = True
    Me.txtDataSemina.Tag = Field.Name
    m_FormFields.Add Field
    
    'DataInizioProduzione
    Set Field = New FormField
    Set Field.Control = Me.txtDataInizioProduzione
    Field.Name = "DataInizioProduzione"
    Field.Visible = True
    Me.txtDataInizioProduzione.Tag = Field.Name
    m_FormFields.Add Field

    'DataFineProduzione
    Set Field = New FormField
    Set Field.Control = Me.txtDataFineProduzione
    Field.Name = "DataFineProduzione"
    Field.Visible = True
    Me.txtDataFineProduzione.Tag = Field.Name
    m_FormFields.Add Field

    'Lotto chiuso
    Set Field = New FormField
    Set Field.Control = Me.chkChiuso
    Field.Name = "Chiuso"
    Field.Visible = True
    Me.chkChiuso.Tag = Field.Name
    m_FormFields.Add Field

    'Codice lotto del seme
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceLottoSeme
    Field.Name = "CodiceLottoSemi"
    Field.Visible = True
    Me.txtCodiceLottoSeme.Tag = Field.Name
    m_FormFields.Add Field

    'Descrizione lotto del seme
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneLottoSeme
    Field.Name = "DescrizioneLottoSemi"
    Field.Visible = True
    Me.txtDescrizioneLottoSeme.Tag = Field.Name
    m_FormFields.Add Field

    'Identificativo della famiglia prodotti
    Set Field = New FormField
    Set Field.Control = Me.cboFamigliaProdotti
    Field.Name = "IDRV_PO01_FamigliaProdotti"
    Field.Visible = True
    Me.cboFamigliaProdotti.Tag = Field.Name
    m_FormFields.Add Field
    
    'Superficie in MQ
    Set Field = New FormField
    Set Field.Control = Me.txtSuperficieMQ
    Field.Name = "DimensioneMQ"
    Field.Visible = True
    Me.txtSuperficieMQ.Tag = Field.Name
    m_FormFields.Add Field
    
    'Superficie in HA
    Set Field = New FormField
    Set Field.Control = Me.txtSuperficieHA
    Field.Name = "DimensioneHA"
    Field.Visible = True
    Me.txtSuperficieHA.Tag = Field.Name
    m_FormFields.Add Field
    
    
    'Certificazione Socio
    Set Field = New FormField
    Set Field.Control = Me.CDCertificazione
    Field.Name = "IDRV_PO01_CertificazioneSocio"
    Field.Visible = True
    Me.CDCertificazione.Tag = Field.Name
    m_FormFields.Add Field
    
    
   'Stato del lotto
    Set Field = New FormField
    Set Field.Control = Me.cboStatoLotto
    Field.Name = "IDRV_PO01_StatoLotto"
    Field.Visible = True
    Me.cboStatoLotto.Tag = Field.Name
    m_FormFields.Add Field


   'Data di sblocco del lotto
    Set Field = New FormField
    Set Field.Control = Me.txtDataSbloccoLotto
    Field.Name = "DataSbloccoLotto"
    Field.Visible = True
    Me.txtDataSbloccoLotto.Tag = Field.Name
    m_FormFields.Add Field


   'Data di semina effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataSeminaEffettiva
    Field.Name = "DataSeminaEffettiva"
    Field.Visible = True
    Me.txtDataSeminaEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'Data di inizio produzione effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataInizioProduzioneEffettiva
    Field.Name = "DataInizioProduzioneEffettiva"
    Field.Visible = True
    Me.txtDataInizioProduzioneEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'Data di fine produzione effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataFineProduzioneEffettiva
    Field.Name = "DataFineProduzioneEffettiva"
    Field.Visible = True
    Me.txtDataFineProduzioneEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'Dimensione in metri quadrati effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtSuperficieMQEffettiva
    Field.Name = "DimensioneMQEffettiva"
    Field.Visible = True
    Me.txtSuperficieMQEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'Dimensione in ettari effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtSuperficieHAEffettiva
    Field.Name = "DimensioneHAEffettiva"
    Field.Visible = True
    Me.txtSuperficieHAEffettiva.Tag = Field.Name
    m_FormFields.Add Field
    

   'Percorso della documentazione
    Set Field = New FormField
    Set Field.Control = Me.txtPercorsoDocumentazione
    Field.Name = "PercorsoDocumentazione"
    Field.Visible = True
    Me.txtPercorsoDocumentazione.Tag = Field.Name
    m_FormFields.Add Field

   'Tipo produzione
    Set Field = New FormField
    Set Field.Control = Me.cboTipoProduzione
    Field.Name = "IDRV_PO01_TipoProduzione"
    Field.Visible = True
    Me.cboTipoProduzione.Tag = Field.Name
    m_FormFields.Add Field

   'Numero progressivo sblocco
    Set Field = New FormField
    Set Field.Control = Me.txtProgressivo
    Field.Name = "NumeroSbloccoLotto"
    Field.Visible = True
    Me.txtProgressivo.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data importazione
    Set Field = New FormField
    Set Field.Control = Me.txtDataImportazione
    Field.Name = "DataImportazione"
    Field.Visible = True
    Me.txtDataImportazione.Tag = Field.Name
    m_FormFields.Add Field

   'Data germinazione
    Set Field = New FormField
    Set Field.Control = Me.txtDataGerm
    Field.Name = "DataGerminazionePresunta"
    Field.Visible = True
    Me.txtDataGerm.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data germinazione effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataGermEffettiva
    Field.Name = "DataGerminazioneEffettiva"
    Field.Visible = True
    Me.txtDataGermEffettiva.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data nascita
    Set Field = New FormField
    Set Field.Control = Me.txtDataNascita
    Field.Name = "DataNascitaPresunta"
    Field.Visible = True
    Me.txtDataNascita.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data nascita effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataNascitaEffettiva
    Field.Name = "DataNascitaEffettiva"
    Field.Visible = True
    Me.txtDataNascitaEffettiva.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data trapianto
    Set Field = New FormField
    Set Field.Control = Me.txtDataTrapianto
    Field.Name = "DataTrapiantoPresunta"
    Field.Visible = True
    Me.txtDataTrapianto.Tag = Field.Name
    m_FormFields.Add Field
    
   'Data trapianto effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataTrapiantoEffettiva
    Field.Name = "DataTrapiantoEffettiva"
    Field.Visible = True
    Me.txtDataTrapiantoEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'Data ripicchettaggio
    Set Field = New FormField
    Set Field.Control = Me.txtDataRipic
    Field.Name = "DataRipicchettaggioPresunta"
    Field.Visible = True
    Me.txtDataRipic.Tag = Field.Name
    m_FormFields.Add Field

   'Data ripicchettaggio effettiva
    Set Field = New FormField
    Set Field.Control = Me.txtDataRipicEffettiva
    Field.Name = "DataRipicchettaggioEffettiva"
    Field.Visible = True
    Me.txtDataRipicEffettiva.Tag = Field.Name
    m_FormFields.Add Field

   'numero piante per Mq
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroPianteMQ
    Field.Name = "NumeroPiantaMQ"
    Field.Visible = True
    Me.txtNumeroPianteMQ.Tag = Field.Name
    m_FormFields.Add Field

   'Annotazioni di concia
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneConcia
    Field.Name = "AnnotazioniConciaSementi"
    Field.Visible = True
    Me.txtDescrizioneConcia.Tag = Field.Name
    m_FormFields.Add Field

   'ID Magazzino di carico lotto
    Set Field = New FormField
    Set Field.Control = Me.cboMagazzinoCaricoLotto
    Field.Name = "IDMagazzinoCaricoLotto"
    Field.Visible = True
    Me.cboMagazzinoCaricoLotto.Tag = Field.Name
    m_FormFields.Add Field

   'ID Causale di carico lotto
    Set Field = New FormField
    Set Field.Control = Me.cboFunzioneCaricoLotto
    Field.Name = "IDFunzioneCaricoLotto"
    Field.Visible = True
    Me.cboFunzioneCaricoLotto.Tag = Field.Name
    m_FormFields.Add Field

   'IDAnagrafica produttore
    Set Field = New FormField
    Set Field.Control = Me.CDProduttore
    Field.Name = "IDAnagraficaFornitoreAcquisto"
    Field.Visible = True
    Me.CDProduttore.Tag = Field.Name
    m_FormFields.Add Field

   'Numero passaporto di acquisto
    Set Field = New FormField
    Set Field.Control = Me.txtNPassaportoProduttore
    Field.Name = "NumeroPassaportoAcquisto"
    Field.Visible = True
    Me.txtNPassaportoProduttore.Tag = Field.Name
    m_FormFields.Add Field

   'Codice di import\Export
    Set Field = New FormField
    Set Field.Control = Me.TxtCodiceImEx
    Field.Name = "CodiceImEx"
    Field.Visible = True
    Me.TxtCodiceImEx.Tag = Field.Name
    m_FormFields.Add Field

   'Protocollo di certificazione
    Set Field = New FormField
    Set Field.Control = Me.txtProtocolloCertificazione
    Field.Name = "ProtocolloCertificazione"
    Field.Visible = True
    Me.txtProtocolloCertificazione.Tag = Field.Name
    m_FormFields.Add Field

   'Identificativo dello schema
    Set Field = New FormField
    Set Field.Control = Me.cboSchema
    Field.Name = "IDRV_PO01_Schema"
    Field.Visible = True
    Me.cboSchema.Tag = Field.Name
    m_FormFields.Add Field

    'IDUtente Inserimento
    Set Field = New FormField
    Set Field.Control = Me.cboUtenteDMTIns
    Field.Name = "IDUtenteInserimento"
    Field.Visible = True
    Me.cboUtenteDMTIns.Tag = Field.Name
    m_FormFields.Add Field
    
    'PCInserimento
    Set Field = New FormField
    Set Field.Control = Me.txtNomePCIns
    Field.Name = "PCInserimento"
    Field.Visible = True
    Me.txtNomePCIns.Tag = Field.Name
    m_FormFields.Add Field

    'UtentePCInserimento
    Set Field = New FormField
    Set Field.Control = Me.txtUtentePCIns
    Field.Name = "UtentePCInserimento"
    Field.Visible = True
    Me.txtUtentePCIns.Tag = Field.Name
    m_FormFields.Add Field

    'Data inserimento
    Set Field = New FormField
    Set Field.Control = Me.txtDataInserimento
    Field.Name = "DataInserimento"
    Field.Visible = True
    Me.txtDataInserimento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Ora inserimento
    Set Field = New FormField
    Set Field.Control = Me.txtOraInserimento
    Field.Name = "OraInserimento"
    Field.Visible = True
    Me.txtOraInserimento.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDUtente Ultima modifica
    Set Field = New FormField
    Set Field.Control = Me.cboUtenteDMTMod
    Field.Name = "IDUtenteUltimaModifica"
    Field.Visible = True
    Me.cboUtenteDMTMod.Tag = Field.Name
    m_FormFields.Add Field

    'PC Ultima modifica
    Set Field = New FormField
    Set Field.Control = Me.txtNomePCMod
    Field.Name = "PCUltimaModifica"
    Field.Visible = True
    Me.txtNomePCMod.Tag = Field.Name
    m_FormFields.Add Field

    'UtentePC Ultima modifica
    Set Field = New FormField
    Set Field.Control = Me.txtUtentePCMod
    Field.Name = "UtentePCUltimaModifica"
    Field.Visible = True
    Me.txtUtentePCMod.Tag = Field.Name
    m_FormFields.Add Field

    'Data ultima modifica
    Set Field = New FormField
    Set Field.Control = Me.txtDataModifica
    Field.Name = "DataUltimaModifica"
    Field.Visible = True
    Me.txtDataModifica.Tag = Field.Name
    m_FormFields.Add Field
    
    'Ora ultima modifica
    Set Field = New FormField
    Set Field.Control = Me.txtOraModifica
    Field.Name = "OraUltimaModifica"
    Field.Visible = True
    Me.txtOraModifica.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDRV_PO01_ClassificazioneLottoProd01
    Set Field = New FormField
    Set Field.Control = Me.cboClass1
    Field.Name = "IDRV_PO01_ClassificazioneLottoProd01"
    Field.Visible = True
    Me.cboClass1.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDRV_PO01_ClassificazioneLottoProd02
    Set Field = New FormField
    Set Field.Control = Me.cboClass2
    Field.Name = "IDRV_PO01_ClassificazioneLottoProd02"
    Field.Visible = True
    Me.cboClass2.Tag = Field.Name
    m_FormFields.Add Field
    
    'Acquistato
    Set Field = New FormField
    Set Field.Control = Me.Check1
    Field.Name = "Acquistato"
    Field.Visible = True
    Me.Check1.Tag = Field.Name
    m_FormFields.Add Field
    
    'Provvisorio
    Set Field = New FormField
    Set Field.Control = Me.Check2
    Field.Name = "Provvisorio"
    Field.Visible = True
    Me.Check2.Tag = Field.Name
    m_FormFields.Add Field
    
    'VirtualDelete
    Set Field = New FormField
    Set Field.Control = Me.chkAnnullato
    Field.Name = "VirtualDelete"
    Field.Visible = True
    Me.chkAnnullato.Tag = Field.Name
    m_FormFields.Add Field
    
    'CodiceLottoCollegato
    Set Field = New FormField
    Set Field.Control = Me.Text1
    Field.Name = "CodiceLottoCollegato"
    Field.Visible = True
    Me.Text1.Tag = Field.Name
    m_FormFields.Add Field
End Sub


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
    'Inserire qui le
    'inizializzazioni da effettuare prima dell'apertura del documento.
    
    
    'rif6 begin
    
    Dim NewLink As DmtDocManLib.Link
    
'**************************ARTICOLI DELLA VARIETA*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink = m_Document.AddDocumentsLink("RV_PO01_DettaglioLotto")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink.PrimaryKey = "IDRV_PO01_DettaglioLotto" '<-- Specifica il campo chiave primaria

    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "CodiceArticolo"
    NewLink.AddLinkColumn "Articolo"

'**************************SERRE INTERESSATE ALLA COLTURA*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink1 = m_Document.AddDocumentsLink("RV_PO01_SerraPerLotto")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink1.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink1.PrimaryKey = "IDRV_PO01_SerraPerLotto" '<-- Specifica il campo chiave primaria

    'Crea un Link LEFT JOIN sulla tabella RV_PO01_Serra
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_PO01_Serra", "RV_PO01_Serra", ltLeft, "IDRV_PO01_Serra")
    NewLink.AddLinkColumn "Codice"
    NewLink.AddLinkColumn "Descrizione"
    NewLink.AddLinkColumn "RV_POIDFeedentity"

'**************************SEMI UTILIZZATI ALLA COLTURA*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink2 = m_Document.AddDocumentsLink("RV_PO01_LottoCampagnaSemi")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink2.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink2.PrimaryKey = "IDRV_PO01_LottoCampagnaSemi" '<-- Specifica il campo chiave primaria

    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink2.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "CodiceArticolo"
    NewLink.AddLinkColumn "Articolo"
    
    'Crea un Link LEFT JOIN sulla tabella LottoArticolo
    Set NewLink = m_DocumentsLink2.AddLink("IDLottoArticolo", "LottoArticolo", ltLeft, "IDLottoArticolo")
    NewLink.AddLinkColumn "Codice"
    NewLink.AddLinkColumn "LottoArticolo"
    
'**************************VERIFICA*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink3 = m_Document.AddDocumentsLink("RV_PO01_LottoCampagnaVerifica")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink3.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink3.PrimaryKey = "IDRV_PO01_LottoCampagnaVerifica" '<-- Specifica il campo chiave primaria
    
'**************************GESTIONE SFALCI*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink4 = m_Document.AddDocumentsLink("RV_PO01_LottoCampagnaSfalcio")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink4.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink4.PrimaryKey = "IDRV_PO01_LottoCampagnaSfalcio" '<-- Specifica il campo chiave primaria

    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink4.AddLink("IDRV_PO01_Sfalcio", "RV_PO01_Sfalcio", ltLeft, "IDRV_PO01_Sfalcio")
    NewLink.AddLinkColumn "Sfalcio"
    NewLink.AddLinkColumn "Sequenza"

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

    ParametroSocio
    ParametroSeme
    ParametroTipoGestioneMovimentazione
    TipoNumerazioneLotto
    ParametroAbilitaRicFor
    GET_UTENTE_SBLOCCO TheApp.IDUser, TheApp.Branch
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
    
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
    'm_Document.Dataset.Recordset.Sort = "DataDocumento DESC, NumeroDocumento DESC"
    'Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
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
        BrwMain.LoadUserSettings
        Me.GrigliaOrdini.LoadUserSettings
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

    BrwMain.Conditions.WidthConditions = 300
    BrwMain.Conditions.WidthFields = 300
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True
    
   
    'For Each Field In m_DocType.Fields
    '    With Field
           'Vengono esclusi dai filtri i campi ID
    '        If Left(.Name, 2) <> "ID" Then
    '            Set Cond = BrwMain.Conditions.Add(.Name, .Name, m_DocType.TableName, False, False, False, ConditionType(.DBType))
    BrwMain.Conditions.Add "Group1", "Dati generali documento", ""
    BrwMain.Conditions("Group1").IsHeader = True
    
        Set Cond = BrwMain.Conditions.Add("Anagrafica", "Socio", m_DocType.TableName, True, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("CodiceSocio", "Codice socio", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("CodiceLotto", "Codice lotto di campagna", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DescrizioneLotto", "Descrizione lotto di campagna", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("CodiceImEx", "Codice Imp\Exp lotto di campagna", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("CodiceLottoCollegato", "Codice lotto collegato", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("IDRV_PO01_FamigliaProdotti", "Famiglia prodotti", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_FamigliaProdotti ORDER BY FamigliaProdotti"
            Cond.DisplayField = "FamigliaProdotti"
            Cond.KeyField = "IDRV_PO01_FamigliaProdotti"
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("IDRV_PO01_Varieta", "Varietà", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_Varieta ORDER BY Varieta"
            Cond.DisplayField = "Varieta"
            Cond.KeyField = "IDRV_PO01_Varieta"
            Cond.Indentation = 20
            
        Set Cond = BrwMain.Conditions.Add("StatoLotto", "Stato", m_DocType.TableName, False, True, , dgCondTypeText)
            Cond.Indentation = 20
        
        'Set Cond = BrwMain.Conditions.Add("IDRV_PO01_StatoLotto", "Stato", m_DocType.TableName, , , , dgCondTypeComboDB)
        '    Cond.RecordSource = "SELECT * FROM RV_PO01_StatoLotto ORDER BY StatoLotto"
        '    Cond.DisplayField = "StatoLotto"
        '    Cond.KeyField = "IDRV_PO01_StatoLotto"
         
         Set Cond = BrwMain.Conditions.Add("IDRV_PO01_TipoProduzione", "Tipo Produzione", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_TipoProduzione ORDER BY TipoProduzione"
            Cond.DisplayField = "TipoProduzione"
            Cond.KeyField = "IDRV_PO01_TipoProduzione"
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("IDRV_PO01_PeriodoCampagna", "Periodo di campagna", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_PeriodoCampagna ORDER BY DataInizio"
            Cond.DisplayField = "PeriodoCampagna"
            Cond.KeyField = "IDRV_PO01_PeriodoCampagna"
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("IDRV_PO01_ClassificazioneLottoProd01", "Classificazione 1", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_ClassificazioneLottoProd01 ORDER BY ClassificazioneLottoProd01"
            Cond.DisplayField = "ClassificazioneLottoProd01"
            Cond.KeyField = "IDRV_PO01_ClassificazioneLottoProd01"
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("IDRV_PO01_ClassificazioneLottoProd02", "Classificazione 2", m_DocType.TableName, , , , dgCondTypeComboDB)
            Cond.RecordSource = "SELECT * FROM RV_PO01_ClassificazioneLottoProd02 ORDER BY ClassificazioneLottoProd02"
            Cond.DisplayField = "ClassificazioneLottoProd02"
            Cond.KeyField = "IDRV_PO01_ClassificazioneLottoProd02"
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("Chiuso", "Chiuso", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
            Cond.FromValue = "NO"
        Set Cond = BrwMain.Conditions.Add("Acquistato", "Merce acquistata", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
        Set Cond = BrwMain.Conditions.Add("Provvisorio", "Provvisorio", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
            Cond.FromValue = "NO"
        Set Cond = BrwMain.Conditions.Add("VirtualDelete", "Annullato", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
            Cond.FromValue = "NO"
            
    BrwMain.Conditions.Add "Group2", "Date presunte o previste", ""
    BrwMain.Conditions("Group2").IsHeader = True

        Set Cond = BrwMain.Conditions.Add("DataSemina", "Data semina prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataGerminazionePresunta", "Data germinazione prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataNascitaPresunta", "Data nascita prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataTrapiantoPresunta", "Data trapianto prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataRipicchettaggioPresunta", "Data ripicchettaggio prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        
        Set Cond = BrwMain.Conditions.Add("DataInizioProduzione", "Data inizio raccolta prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataFineProduzione", "Data fine raccolta prevista", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
            
    
    BrwMain.Conditions.Add "Group3", "Date effettive", ""
    BrwMain.Conditions("Group3").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("DataSeminaEffettiva", "Data semina effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataGerminazioneEffettiva", "Data germinazione effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataNascitaEffettiva", "Data nascita effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataTrapiantoEffettiva", "Data trapianto effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataRipicchettaggioEffettiva", "Data ripicchettaggio effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        
        Set Cond = BrwMain.Conditions.Add("DataInizioProduzioneEffettiva", "Data inizio raccolta effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("DataFineProduzioneEffettiva", "Data fine raccolta effettiva", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
    BrwMain.Conditions.Add "Group4", "Altre dati", ""
    BrwMain.Conditions("Group4").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("DataImportazione", "Data importazione", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        
        
                
                
                

                
                'Non viene visualizzata la Check Intervallo
          
    '        End If
    '    End With
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
    
    'Tools-renewe
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Rinnova_lotti"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Caption = "Rinnovo" 'GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Rinnova_lotti").Description = "Avvia la procedura di rinnovo dei lotti di produzione" 'GetDescription4StatusBar("Mnu_Options")
    
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
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep18"
    BarMenu.Bands("Standard").Tools("Sep18").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdSalvaComeNuovo"
    BarMenu.Bands("Standard").Tools("cmdSalvaComeNuovo").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("cmdSalvaComeNuovo").SetPicture 0, gResource.GetBitmap(IDB_STD_ICONS16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("cmdSalvaComeNuovo").ToolTipText = "Salva come nuovo" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdSalvaComeNuovo").Description = "Salva come nuovo"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdSalvaComeNuovo").Caption = "Salva come nuovo"  'GetDescription4StatusBar("Mnu_FormView")
        
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep19"
    BarMenu.Bands("Standard").Tools("Sep19").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdGestioneSfalci"
    BarMenu.Bands("Standard").Tools("cmdGestioneSfalci").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("cmdGestioneSfalci").SetPicture 0, gResource.GetBitmap(IDB_STD_LIST16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("cmdGestioneSfalci").ToolTipText = "Gestione sfalci" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdGestioneSfalci").Description = "Gestione sfalci"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdGestioneSfalci").Caption = "Gestione sfalci"  'GetDescription4StatusBar("Mnu_FormView")
    
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep16"
    BarMenu.Bands("Standard").Tools("Sep16").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdAvviaQualita"
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").SetPicture 0, gResource.GetBitmap(IDB_ACT_ORDER_CUSTOMER_EVASION_16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").ToolTipText = "Gestione qualità" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Description = "Gestione qualità"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Caption = "Qualità"  'GetDescription4StatusBar("Mnu_FormView")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep17"
    BarMenu.Bands("Standard").Tools("Sep17").ControlType = ddTTSeparator
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
On Error GoTo ERR_ExecuteSearch
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
        'sWhere = ""
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
Exit Sub
ERR_ExecuteSearch:
    MsgBox Err.Description, vbCritical, "ExecuteSearch"
    
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
    
    m_DocType.Fields("IDFiliale").Value = m_App.Branch
    m_DocType.Fields("IDUtente").Value = m_App.IDUser
    
    
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
       
   
   
    If Me.Griglia.ColumnsHeader.Count = 0 Then
        With Me.Griglia.ColumnsHeader
            .Add "IDRV_PO01_DettaglioLotto", "IDRV_PO01_DettaglioLotto", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "CodiceArticolo", "Codice ", dgchar, True, 2000, 0, True, True, False
            .Add "Articolo", "Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "Calibro", "Calibro", dgchar, True, 2000, 0, True, True, False
            Set cl = .Add("QuantitaPresunta", "Q.tà presunta", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
    
        End With
    End If
    Me.Griglia.EnableMove = True
    
    If Me.GrigliaTerreni.ColumnsHeader.Count = 0 Then
        With Me.GrigliaTerreni.ColumnsHeader
            .Add "IDRV_PO01_SerraPerLotto", "IDRV_PO01_SerraPerLotto", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "Codice", "Codice", dgchar, True, 2000, 0, True, True, False
            .Add "Descrizione", "Descrizione", dgchar, True, 2000, 0, True, True, False
            Set cl = .Add("DimensioneMQ", "Sup. in Mq", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "DimensioneHA", "Sup. in Ha", dgchar, True, 1500, 0, True, True, False
            .Add "RV_POIDFeedentity", "ID Feed", dgchar, True, 2000, 0, True, True, False
        
        End With
    End If
    Me.GrigliaTerreni.EnableMove = True


    If Me.GrigliaSemi.ColumnsHeader.Count = 0 Then
        With Me.GrigliaSemi.ColumnsHeader
            .Add "IDRV_PO01_LottoCampagnaSemi", "IDRV_PO01_LottoCampagnaSemi", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "CodiceArticolo", "Codice articolo", dgchar, True, 2000, 0, True, True, False
            .Add "Articolo", "Descrizione articolo", dgchar, True, 2000, 0, True, True, False
            .Add "IDLottoArticolo", "IDLottoArticolo", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "Codice", "Codice lotto ", dgchar, True, 2000, 0, True, True, False
            .Add "LottoArticolo", "Descrizione lotto", dgchar, True, 2000, 0, True, True, False
            Set cl = .Add("Quantita", "Quantità", dgDouble, True, 1500, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
    
        End With
    End If
    Me.GrigliaSemi.EnableMove = True

    If Me.GrigliaVerifica.ColumnsHeader.Count = 0 Then
        With Me.GrigliaVerifica.ColumnsHeader
            .Add "IDRV_PO01_LottoCampagnaVerifica", "IDRV_PO01_LottoCampagnaVerifica", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "IDRV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "DataVerifica", "Data", dgDate, True, 2000, 0
            .Add "OraVerifica", "Ora", dgchar, True, 2000, dgAligncenter
            .Add "Operatore", "Operatore", dgchar, True, 2000, dgAlignleft
            .Add "TipoRilevazione", "Tipo rilevazione", dgchar, True, 2000, dgAlignleft
            .Add "Temperatura", "Temperatura", dgchar, True, 1200, dgAlignleft
            .Add "Annotazioni", "Annotazioni", dgchar, True, 3000, dgAlignleft
    
        End With
    End If
    Me.GrigliaVerifica.EnableMove = True

    If Me.GrigliaSfalcio.ColumnsHeader.Count = 0 Then
        With Me.GrigliaSfalcio.ColumnsHeader
            .Add "IDRV_PO01_LottoCampagnaSfalcio", "IDRV_PO01_LottoCampagnaSfalcio", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "IDRV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "IDRV_PO01_Sfalcio", "IDRV_PO01_Sfalcio", dgNumeric, False, 500, dgAlignRight, True, True, False
            .Add "Sequenza", "Sequenza", dgNumeric, True, 1500, dgAlignRight
            .Add "Sfalcio", "Descrizione sfalcio", dgchar, True, 3000, dgAlignleft
            .Add "DataPresuntaInizio", "Data presunta", dgDate, True, 2000, 0
            .Add "DataEffettivaInizio", "Data effettiva", dgDate, True, 2000, 0
            .Add "Guid", "Guid", dgchar, False, 3000, dgAlignleft
        End With
    End If
    Me.GrigliaSfalcio.EnableMove = True

'''''''''''''''''''''''''CONTROLLI STANDARD''''''''''''''''''''''''''''''''''''
    
    
    With Me.CDArticolo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDRV_PO01_TipoProdottoBio <> " & 1 & ") OR NOT (IDRV_PO01_TipoProdottoBio  IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
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


    With Me.CDSerre
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Descrizione"
        .KeyField = "IDRV_PO01_Serra"
        .TableName = "RV_PO01_Serra"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDRV_PO01_TipoProdottoBio <> " & 1 & ") OR NOT (IDRV_PO01_TipoProdottoBio  IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND RV_PO01_IDVarieta =" & Me.cboVarieta.CurrentID & " AND GestioneLotti = " & fnNormBoolean(1)
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice serra"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione serra"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fnGetTipoOggetto("RV_PO01_Serre")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Periodo di campagna
    With Me.cboPeriodoCampagna
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_PeriodoCampagna"
        .DisplayField = "PeriodoCampagna"
        .SQL = "SELECT * FROM RV_PO01_PeriodoCampagna "
        .SQL = .SQL & "WHERE IDAzienda=" & m_App.IDFirm
        .SQL = .SQL & " AND IDFiliale=" & TheApp.Branch
        .Fill
    End With

    'Famiglia prodotti
    With Me.cboFamigliaProdotti
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_FamigliaProdotti"
        .DisplayField = "FamigliaProdotti"
        .SQL = "SELECT * FROM RV_PO01_FamigliaProdotti ORDER BY FamigliaProdotti"
        .Fill
    End With

    'Stato del lotto
    With Me.cboStatoLotto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_StatoLotto"
        .DisplayField = "StatoLotto"
        .SQL = "SELECT * FROM RV_PO01_StatoLotto ORDER BY StatoLotto"
        .Fill
    End With

    'Stato del lotto
    With Me.cboTipoProduzione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_TipoProduzione"
        .DisplayField = "TipoProduzione"
        .SQL = "SELECT * FROM RV_PO01_TipoProduzione ORDER BY TipoProduzione"
        .Fill
    End With


    'Anagrafica socio
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        If IndicaRicFor = 0 Then
            .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDCategoriaAnagrafica=" & Link_TipoSocio
        Else
            .Filter = "IDAzienda = " & m_App.IDFirm
        End If
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


    With Me.CDCertificazione
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceCertificazione"
        .DescriptionField = "DescrizioneCertificazione"
        .KeyField = "IDRV_PO01_Certificazione"
        .TableName = "RV_PO01_Certificazione"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDRV_PO01_TipoProdottoBio <> " & 1 & ") OR NOT (IDRV_PO01_TipoProdottoBio  IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND RV_PO01_IDVarieta =" & Me.cboVarieta.CurrentID & " AND GestioneLotti = " & fnNormBoolean(1)
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fnGetTipoOggetto("RV_PO01_Certificazione")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    
    'Magazzino di carico lotto vegetali
    With Me.cboMagazzinoCaricoLotto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT * FROM Magazzino WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    'Causali di movimento carico vegatali
    With Me.cboFunzioneCaricoLotto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    'Anagrafica produttore
    With Me.CDProduttore
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

    With Me.CDArticoloSemi
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDRV_PO01_TipoProdottoBio <> " & 1 & ") OR NOT (IDRV_PO01_TipoProdottoBio  IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
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

    'Utente DMT inserimento
    With Me.cboUtenteDMTIns
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With

    'Utente DMT modifica
    With Me.cboUtenteDMTMod
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With

    'Utente DMT inserimento prodotto
    With Me.cboUtenteDMTInsArt
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With

    'Utente DMT modifica prodotto
    With Me.cboUtenteDMTModArt
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With

    'Sfalcio
    With Me.cboSfalcio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_Sfalcio"
        .DisplayField = "Sfalcio"
        .SQL = "SELECT * FROM RV_PO01_Sfalcio ORDER BY IDRV_PO01_Sfalcio"
        .Fill
    End With
    
    'Classificazione 01
    With Me.cboClass1
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_ClassificazioneLottoProd01"
        .DisplayField = "ClassificazioneLottoProd01"
        .SQL = "SELECT * FROM RV_PO01_ClassificazioneLottoProd01 ORDER BY ClassificazioneLottoProd01"
        .Fill
    End With
    
    'Classificazione 02
    With Me.cboClass2
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_ClassificazioneLottoProd02"
        .DisplayField = "ClassificazioneLottoProd02"
        .SQL = "SELECT * FROM RV_PO01_ClassificazioneLottoProd02 ORDER BY ClassificazioneLottoProd02"
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
On Error GoTo ERR_OnSave
Dim Field As DmtDocManLib.Field
Dim DocLink As DmtDocManLib.DocumentsLink
Dim sSQL As String
Dim NumeroInserimenti As Long
Dim NuovoDocumento As Boolean
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long

    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If
    
    If Not PermissionToSave Then
        Exit Sub
    End If
        
    NuovoDocumento = False
    
    'If GET_TIPO_CODICE = 2 Then
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'SALVA_NUMERO_LOTTO
        NuovoDocumento = True
    End If
    'End If
    
    SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
    APERTURA_FORM_CODA = False

    Me.cboUtenteDMTMod.WriteOn TheApp.IDUser
    Me.txtNomePCMod.Text = GET_NOMECOMPUTER
    Me.txtUtentePCMod.Text = GET_NOMEUTENTE
    Me.txtDataModifica.Value = Date
    Me.txtOraModifica.Text = GET_ORARIO(Now)
    
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
    
                If Field.Name = "IDArticoloLottoSemi" Then
                    Field.Value = Link_LottoSeme
                End If
                    
                If Field.Name = "IDTipoOggetto" Then
                    Field.Value = fnGetTipoOggetto("RV_PO01_CreazioneLotto")
                End If
                
                If Field.Name = "Anagrafica" Then
                    Field.Value = Me.CDSocio.Description
                End If
                
                If Field.Name = "Nome" Then
                    Field.Value = Me.txtNomeSocio.Text
                End If
                
                If Field.Name = "CodiceSocio" Then
                    Field.Value = Me.CDSocio.Code
                End If
                If Field.Name = "GuidID" Then
                    If m_Document(m_Document.PrimaryKey).Value <= 0 Then
                        Field.Value = GetGUID
                    Else
                        If Len(GUID_LOTTOCAMPAGNA) = 0 Then
                            Field.Value = GetGUID
                        Else
                            Field.Value = GUID_LOTTOCAMPAGNA
                        End If
                        
                    End If
                End If
                If Field.Name = "DataUltimoAggiornamento" Then
                    Field.Value = Now
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

    Me.txtSuperficieMQ.Value = GET_TOTALE_SUPERFICIE_LOTTO
    txtSuperficieMQ_LostFocus
    
    m_Document("DimensioneMQ").Value = Me.txtSuperficieMQ.Value
    m_Document("DimensioneHA").Value = Me.txtSuperficieHA.Text
    
    
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Me.Enabled = True
    'Me.SetFocus
    'Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
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
    
    Cn.CursorLocation = OLD_Cursor
    
    Cn.CommitTrans
    
    fnModificaNomeLotto m_Document(m_Document.PrimaryKey).Value, Me.txtCodiceLotto.Text, Me.txtDescrizioneLotto.Text
        
    If SALVA_COME_NUOVO = True Then
        If RIPORTA_PRODOTTI = 1 Then
            RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO LINK_LOTTO_CAMPAGNA_SCN, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        End If
        
        If RIPORTA_SERRE_APPEZZAMENTI = 1 Then
            RIPORTA_SERRE_DA_SALVA_COME_NUOVO LINK_LOTTO_CAMPAGNA_SCN, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        End If
        SALVA_COME_NUOVO = False
        OnSave
    End If
    
    frmAttesa.lblInfo.Caption = "CREAZIONE SFALCI..."
    CREA_SFALCI_PER_LOTTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value), DatePart("ww", Me.txtDataSemina.Text, , vbFirstFullWeek), Me.cboFamigliaProdotti.CurrentID
    
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    
    
    
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
    
    'Refresh delle variabili di stato
    m_Changed = False
    m_Search = False
    m_Saved = True
    
    'Refresh dello stato della ToolBar standard in modalità variazione
    SetStatus4Modality Modify

    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    m_DocumentsLink.Refresh
    m_DocumentsLink1.Refresh
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Enabled = True
        Me.cboMagazzinoCaricoLotto.Enabled = True
        Me.cboFunzioneCaricoLotto.Enabled = True
    Else
        'Me.txtCodiceLotto.Enabled = False
        Me.cboMagazzinoCaricoLotto.Enabled = False
        Me.cboFunzioneCaricoLotto.Enabled = False
    
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
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)
    
    
End Sub
Private Function GET_TOTALE_SUPERFICIE_LOTTO() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sum(DimensioneMQ) as TotaleSuperficie "
sSQL = sSQL & "FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_SUPERFICIE_LOTTO = 0
Else
    GET_TOTALE_SUPERFICIE_LOTTO = fnNotNullN(rs!TotaleSuperficie)
End If


rs.CloseResultset
Set rs = Nothing
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
    Dim sToRemove As String
    Dim DocLink As DmtDocManLib.DocumentsLink
    Dim TestoMessaggio As String
    Dim LINK_TESTA_DOCUMENTO As Long
    Dim CODICE_LOTTO As String
    
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
                
        If fncControlloMovimentazioneLotto = True Then
            MsgBox "Il lotto risulta movimentato nel quaderno di campagna", vbInformation, "Impossibile eliminare"
            Exit Sub
        End If
           
        If GET_CONTROLLO_ESISTENZA_MOV_RIGHE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
            TestoMessaggio = "ATTENZIONE!!!!" & vbCrLf
            TestoMessaggio = TestoMessaggio & "Impossibile eliminare il documento poichè alcune righe articolo sono state movimentate da altri documenti"
            MsgBox TestoMessaggio, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
        'If GET_ESISTENZA_PROGRAMMA(5) = True Then
        If GET_ESISTENZA_LOTTO_CONFERITO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
            TestoMessaggio = "ATTENZIONE!!!!" & vbCrLf
            TestoMessaggio = TestoMessaggio & "Questo lotto risulta conferito, pertanto è impossibile eliminare"
            MsgBox TestoMessaggio, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
        'End If
        'rif16
        
        LINK_TESTA_DOCUMENTO = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        CODICE_LOTTO = Me.txtCodiceLotto.Text
        'Cancellazione
        m_Document.DeleteDocument
        
        
        ELIMINA_MOVIMENTI_DOCUMENTO LINK_TESTA_DOCUMENTO
        
        ELIMINA_RIFERIMENTI_ORDINE CODICE_LOTTO
        

        fncEliminaRiferimentiSerre LINK_TESTA_DOCUMENTO
        
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
    
    GENERA_FILTRO_PER_TIPO_OGGETTO
    
    'Nota: utilizzo la chiamata al metodo ApplyFilter della dmtGrid piuttosto
    'che la chiamata diretta di ExecuteSearch perchè in questo modo la dmtGrid
    'può gestire internamente le conditions di ricerca.
    'Verrà generato l'evento BrwMain_OnApplyFilter()
    '
    'ExecuteSearch
    '
    BrwMain.ApplyFilter
    
    Set rsRinnovo = BrwMain.Recordset
    
        
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
       ' N' 'ext
        
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










Private Sub BrwMain_ConditionEdit(ByVal Name As String, Value As Variant)
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = TheApp.Database.Connection

If Name = "Socio" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Fornitore"

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
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica "
    sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & TheApp.IDFirm & " AND "
    sSQL = sSQL & "Anagrafica.IDCategoriaAnagrafica=" & Link_TipoSocio
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
If Name = "Lotto di campagna" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Lotto di campagna"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Codice", "CodiceLotto", 1 'STRINGTYPE
    oSearch.AddDisplayField "Descrizione", "DescrizioneLotto", 1   'STRINGTYPE
    oSearch.AddDisplayField "Chiuso", "Chiuso", 0
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDRV_PO01_LottoCampagna, CodiceLotto, DescrizioneLotto, Chiuso "
    sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
    sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("CodiceLotto")
                
    End If
End If
End Sub

Private Sub cboClass1_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboClass2_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboFamigliaProdotti_Click()
    
    Me.txtCodiceFamigliaProdotto.Text = GET_CODICE_TABELLA("RV_PO01_FamigliaProdotti", "CodiceImEx", "IDRV_PO01_FamigliaProdotti", Me.cboFamigliaProdotti.CurrentID)
    
     If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'If Len(Me.cboFamigliaProdotti.Text) > 0 Then
        '    CreaCodiceLotto
        'End If
        
        GET_CERTIFICAZIONE Me.CDSocio.KeyFieldID, Me.cboFamigliaProdotti.CurrentID, m_App.Branch
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
        
    End If
    
    
    
    If Not (BrwMain.Visible) Then Change

End Sub
Private Sub CreaCodiceLotto()
Dim Testo As String
Dim MM As String
Dim GG As String
Dim CodiceLotto As String
Dim DescrizioneLotto As String
Dim DataLotto As String
Dim IDTipoCodice As Long

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then

    IDTipoCodice = GET_TIPO_CODICE
    Select Case IDTipoCodice
    
        Case 1
            If Len(Me.cboFamigliaProdotti.Text) <= 22 Then
                Testo = Mid(Me.cboFamigliaProdotti.Text, 1, Len(Me.cboFamigliaProdotti.Text))
            Else
                Testo = Mid(Me.cboFamigliaProdotti.Text, 1, 22)
            End If
            
            If Len(Me.txtDataSemina.Text) > 0 Then
                DataLotto = Me.txtDataSemina.Text
            Else
                DataLotto = Date
            End If
            
            If Len(DatePart("m", DataLotto)) = 1 Then
                MM = "0" & DatePart("m", DataLotto)
            Else
                MM = DatePart("m", DataLotto)
            End If
            
            If Len(DatePart("d", DataLotto)) = 1 Then
                GG = "0" & DatePart("d", DataLotto)
            Else
                GG = DatePart("d", DataLotto)
            End If
            
            
            
            Me.txtCodiceLotto.Text = Testo & "_" & DatePart("yyyy", DataLotto) & MM & GG
            Me.txtDescrizioneLotto.Text = Testo & "_" & DatePart("yyyy", DataLotto) & MM & GG
            StringaLottoStd = Me.txtCodiceLotto.Text
        Case 2
            Me.txtCodiceLotto.Text = GET_NUMERO_LOTTO
            Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End Select
End If
End Sub

Private Sub cboFunzioneCaricoLotto_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboMagazzinoCaricoLotto_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboPeriodoCampagna_Click()
    Me.txtCodicePeriodoCampagna.Text = GET_CODICE_TABELLA("RV_PO01_PeriodoCampagna", "Codice", "IDRV_PO01_PeriodoCampagna", Me.cboPeriodoCampagna.CurrentID)
    Me.txtAnnoRifPeriodoCampagna.Text = GET_CODICE_TABELLA("RV_PO01_PeriodoCampagna", "AnnoDiRiferimento", "IDRV_PO01_PeriodoCampagna", Me.cboPeriodoCampagna.CurrentID)
 
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End If
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboSchema_Click()
Dim Link_schema_local
Dim Testo As String
Dim NumeroRecord As Long
Dim sSQL As String

If AggiornamentoGriglia = 0 Then
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) > 0 Then
        If Me.cboSchema.CurrentID <> fnNotNullN(m_Document("IDRV_PO01_Schema").Value) Then
            Testo = "ATTENZIONE!!!!" & vbCrLf
            Testo = Testo & "Cambiando lo schema delle serre\appezzamenti del socio, "
            Testo = Testo & "verranno eliminati tutti i dati inseriti precedentemente nella tasca 'Serre/Appezzamenti'" & vbCrLf
            Testo = Testo & "Continuare con questo comando?"
            
            If MsgBox(Testo, vbQuestion + vbYesNo, "Schema") = vbNo Then
                Me.cboSchema.WriteOn fnNotNullN(m_Document("IDRV_PO01_Schema").Value)
            Else
                NumeroRecord = Me.GrigliaTerreni.ListIndex - 1
            
                sSQL = "DELETE FROM RV_PO01_SerraPerLotto "
                sSQL = sSQL & "WHERE " & m_Document.PrimaryKey & "=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
                Cn.Execute sSQL
                
                m_DocumentsLink1.Refresh
    
                OnSave
                
                If Not (Me.GrigliaTerreni.Recordset.EOF And Me.GrigliaTerreni.Recordset.BOF) Then
                    Me.GrigliaTerreni.Recordset.Move NumeroRecord
                End If
    
            End If
            
        End If
        
    End If
End If


With Me.CDSerre
    Set .Application = TheApp
    Set .Database = TheApp.Database
    .HwndContainer = Me.hwnd
    .CodeField = "Codice"
    .DescriptionField = "Descrizione"
    .KeyField = "IDRV_PO01_Serra"
    
    If (fnNotNullN(m_Document("Provvisorio").Value) = 0) Then
        .TableName = "RV_PO01_IESchemaSerra"
        .Filter = "IDRV_PO01_Schema=" & fnNotNullN(m_Document("IDRV_PO01_Schema").Value)
    Else
        .TableName = "RV_PO01_IESerra"
        .Filter = "IDSocio=" & fnNotNullN(m_Document("IDSocio").Value)
    End If
    .MenuFunctions("EseguiGestione").Enabled = True
    .PropCodice.Caption = "Codice"
    'Caption da associare alla label del campo Descrizione
    .PropDescrizione.Caption = "Descrizione"
    'Caption da associare alla intestazione della colonna della Find per il campo Codice
    .CodeCaption4Find = "Codice serra"
    'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
    .DescriptionCaption4Find = "Descrizione serra"
    'Identificativo della Funzione Diamante per l'Esegui Gestione
    .IDExecuteFunction = fnGetTipoOggetto("RV_PO01_Serre")
    'Indica se il campo Codice è un campo numerico
    .CodeIsNumeric = False
End With




End Sub

Private Sub cboStatoLotto_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoProduzione_Click()
    Me.txtCodiceTipoProduzione.Text = GET_CODICE_TABELLA("RV_PO01_TipoProduzione", "CodiceImEx", "IDRV_PO01_TipoProduzione", Me.cboTipoProduzione.CurrentID)
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End If
    If Not (BrwMain.Visible) Then Change
End Sub





Private Sub CDArticolo_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.TxtArticolo.Text = ""
Else
    Me.TxtArticolo.Text = fnNotNull(rs!Articolo)
End If

rs.CloseResultset
Set rs = Nothing

End Sub



Private Sub CDCertificazione_ChangeElement()
On Error Resume Next
    
    If Me.CDCertificazione.KeyFieldID <> fnNotNullN(m_Document("IDRV_PO01_CertificazioneSocio").Value) Then
        Me.txtProtocolloCertificazione.Text = ""
    End If
    If Me.CDCertificazione.KeyFieldID = 0 Then Me.txtProtocolloCertificazione.Text = ""
    'Me.txtProtocolloCertificazione.Text = GET_PROTOCOLLO_CERTIFICAZIONE(Me.CDCertificazione.KeyFieldID)
    If Not (BrwMain.Visible) Then Change
    
End Sub

Private Sub CDProduttore_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Nome FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & Me.CDProduttore.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNomeFornitore.Text = ""
Else
    Me.txtNomeFornitore.Text = fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDSocio_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Nome FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & Me.CDSocio.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNomeSocio.Text = ""
Else
    Me.txtNomeSocio.Text = fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing

With Me.cboSchema
    Set .Database = TheApp.Database.Connection
    .AddFieldKey "IDRV_PO01_Schema"
    .DisplayField = "SchemaSettori"
    .SQL = "SELECT * FROM RV_PO01_Schema "
    .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
    .SQL = .SQL & " AND IDSocio=" & Me.CDSocio.KeyFieldID
    .Fill
End With



If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    GET_CERTIFICAZIONE Me.CDSocio.KeyFieldID, Me.cboFamigliaProdotti.CurrentID, m_App.Branch
    'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
    'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    Me.cboSchema.WriteOn GET_SCHEMA_PREDEFINITO_SOCIO(Me.CDSocio.KeyFieldID)
End If



If Not (BrwMain.Visible) Then Change

End Sub

Private Sub Check1_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub Check2_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAnnullato_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkChiuso_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdElabora_Click()
On Error GoTo ERR_cmdElabora_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ControlloEsistenza As Boolean
Dim Link_Testa As Long
Dim rsRighe As DmtOleDbLib.adoResultset
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire i prodotti derivati", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If Me.cboFamigliaProdotti.CurrentID = 0 Then
    MsgBox "Inserire il tipo di coltura del lotto prima di elaborare gli articoli", vbInformation, "Elaborazione dati"
    Exit Sub
End If

If Len(Me.txtCodiceLotto.Text) = 0 Then
    MsgBox "Inserire il codice del lotto prima di elaborare gli articoli", vbInformation, "Elaborazione dati"
    Exit Sub
End If
If Len(Me.txtDescrizioneLotto.Text) = 0 Then
    MsgBox "Inserire la descrizione del lotto prima di elaborare gli articoli", vbInformation, "Elaborazione dati"
    Exit Sub
End If


AggiornamentoGriglia = 1


SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
APERTURA_FORM_CODA = False


    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) > 0 Then
    sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo, GestioneLotti FROM Articolo "
    sSQL = sSQL & "WHERE RV_PO01_IDFamigliaProdotti=" & Me.cboFamigliaProdotti.CurrentID
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    
    Set rs = Cn.OpenResultset(sSQL)
    
    
    
    While Not rs.EOF
        If fnNotNullN(rs!IDArticolo) > 0 Then
            If EsistenzaArticolo(rs!IDArticolo, fnNotNullN(m_Document("IDRV_PO01_LottoCampagna").Value)) = False Then
                   
                m_DocumentsLink.NewRow
                
                m_DocumentsLink("IDArticolo").Value = rs!IDArticolo
                m_DocumentsLink("Calibro").Value = ""
                m_DocumentsLink("QuantitaPresunta").Value = 0
                m_DocumentsLink.Save
                
                
                If Abs(fnNotNullN(rs!GestioneLotti)) = 1 Then
                    fnGetLottoArticolo m_DocumentsLink("IDArticolo").Value
                    If MOVIMENTAZIONE_ARTICOLO(fnNotNullN(rs!IDArticolo), fnNotNull(rs!Articolo), fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value), 0, Date) = False Then
                        MsgBox "Impossibile movimentare la riga dell'articolo", vbCritical, "Movimentazione articolo"
                    End If
                End If
                    
            Else
                If Abs(fnNotNullN(rs!GestioneLotti)) = 1 Then
                    If GET_ESISTENZA_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, fnNotNullN(rs!IDArticolo)) = False Then
                        fnGetLottoArticolo m_DocumentsLink("IDArticolo").Value
                    End If
                End If
            End If
        End If
    
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    m_DocumentsLink.Refresh
    
    AggiornamentoGriglia = 0
    

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display(False)
    Cn.CursorLocation = OLD_Cursor
    
    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    


    OnSave
    'If Not (BrwMain.Visible) Then Change
End If
Exit Sub
ERR_cmdElabora_Click:

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)
End Sub
Private Function EsistenzaArticolo(IDArticolo As Long, IDTesta As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_DettaglioLotto "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDRV_PO01_LottoCampagna=" & IDTesta

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    EsistenzaArticolo = True
Else
    EsistenzaArticolo = False
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub cmdElimina_Quadratura_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Lotto As Long
Dim GESTIONE_LOTTI As Boolean
Dim Testo As String


If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then Exit Sub

If GET_CONTROLLO_AZIONI_UTENTE(m_DocType.ID, TheApp.IDUser, TheApp.Branch, 4) = False Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Non si hanno i diritti per eseguire questo comando"
    
    MsgBox Testo, vbCritical, "Controllo azioni utente"
    Exit Sub
    
End If

If MsgBox("Vuoi eliminare la riga del documento?", vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
sSQL = "Select GestioneLotti FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GESTIONE_LOTTI = False
Else
    GESTIONE_LOTTI = fnNotNullN(rs!GestioneLotti)
End If


If GESTIONE_LOTTI = True Then
    Link_Lotto = GET_LINK_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, Me.CDArticolo.KeyFieldID)
    If Link_Lotto > 0 Then
        If GET_ESISTENZA_MOVIMENTAZIONE_LOTTO(Me.CDArticolo.KeyFieldID, Link_Lotto, fnNotNullN(m_Document(m_Document.PrimaryKey).Value), m_DocType.ID) = True Then
            MsgBox "Impossibile eliminare la riga poichè risulta essere movimentato", vbCritical, "Elimazione dati"
            Exit Sub
        Else
            Set Mov = New DmtMovim.cMovimentazione
            Set Mov.Connection = TheApp.Database.Connection
            
            ''''''''ELIMINAZIONE MOVIMENTO'''''''''''''''''''''''''''''''''''''''
            Mov.IDTipoOggetto = m_DocType.ID
            Mov.IDOggetto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
            Mov.Field "IDValoriOggettoDettaglio", fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)
            Mov.Delete
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Set Mov = Nothing
            
            ELIMINA_LOTTO_ARTICOLO Me.CDArticolo.KeyFieldID, Me.txtCodiceLotto.Text
        End If
    End If
End If


m_DocumentsLink.Delete
    
'If Not (BrwMain.Visible) Then Change

End Sub

Private Sub cmdEliminaOrdineAssociato_Click()
On Error GoTo ERR_cmdEliminaOrdineAssociato_Click
Dim Testo As String
Dim sSQL As String

    If (Me.GrigliaOrdini.Recordset.EOF) And (Me.GrigliaOrdini.Recordset.BOF) Then Exit Sub

    If fnNotNullN(Me.GrigliaOrdini.AllColumns("IDValoriOggettoDettaglio").Value) = 0 Then Exit Sub
    
    Testo = "Sei sicuro di eliminare il riferimento all'ordine?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento riga di ordine") = vbNo Then Exit Sub
    
    sSQL = "UPDATE ValoriOggettoDettaglio0010 SET "
    sSQL = sSQL & "RV_PO01_UbicazioneLottoDiCampagna=" & fnNormString("") & ", "
    sSQL = sSQL & "RV_PO01_LottoDiCampagna=" & fnNormString("") & ", "
    sSQL = sSQL & "RV_PO01_DescrLottoDiCampagna=" & fnNormString("") & ", "
    sSQL = sSQL & "RV_PO01_DataSemina=" & fnNormDate(Null) & ", "
    sSQL = sSQL & "RV_PO01_IDLottoCampagna=0 "
    sSQL = sSQL & " WHERE RV_POLinkRiga=" & fnNotNullN(Me.GrigliaOrdini.AllColumns("RV_POLinkRiga").Value)
    
    Cn.Execute sSQL
    
    GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    GET_RIEPILOGO_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Exit Sub
ERR_cmdEliminaOrdineAssociato_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaOrdineAssociato_Click"
End Sub

Private Sub cmdEliminaSemi_Click()
Dim NumeroRecord As Long

    NumeroRecord = Me.GrigliaSemi.ListIndex - 1

    m_DocumentsLink2.Delete
    
    OnSave
    
    If Not (Me.GrigliaSemi.Recordset.EOF And Me.GrigliaSemi.Recordset.BOF) Then
        Me.GrigliaSemi.Recordset.Move NumeroRecord
    End If
    
End Sub

Private Sub cmdEliminaSfalcio_Click()
On Error GoTo ERR_cmdEliminaSfalcio_Click
Dim NumeroRecord As Long
    
    If GET_CONTROLLO_ELIMINAZIONE_SFALCIO(fnNotNullN(m_DocumentsLink4(m_DocumentsLink4.PrimaryKey).Value)) = False Then
        MsgBox "Impossibile eliminare il record perchè utilizzato in altre funzionalità", vbInformation, "Validazione dati"
        Exit Sub
    End If

    NumeroRecord = Me.GrigliaSfalcio.ListIndex - 1

    m_DocumentsLink4.Delete
    
    If Not (Me.GrigliaSfalcio.Recordset.EOF And Me.GrigliaSfalcio.Recordset.BOF) Then
        Me.GrigliaSfalcio.Recordset.Move NumeroRecord
    End If
    
Exit Sub
ERR_cmdEliminaSfalcio_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaSfalcio_Click"
End Sub

Private Sub cmdEliminaTerreno_Click()
Dim NumeroRecord As Long
Dim Testo As String

    If GET_CONTROLLO_AZIONI_UTENTE(m_DocType.ID, TheApp.IDUser, TheApp.Branch, 4) = False Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Non si hanno i diritti per eseguire questo comando"
        
        MsgBox Testo, vbCritical, "Controllo azioni utente"
        Exit Sub
        
    End If

    NumeroRecord = Me.GrigliaTerreni.ListIndex - 1

    m_DocumentsLink1.Delete
    
    
    OnSave

    AGGIORNA_ORDINI_DA_SERRE Me.txtCodiceLotto.Text, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    If Not (Me.GrigliaTerreni.Recordset.EOF And Me.GrigliaTerreni.Recordset.BOF) Then
        Me.GrigliaTerreni.Recordset.Move NumeroRecord
    End If
    
End Sub

Private Sub cmdEliminaVerifica_Click()
Dim Testo As String
Dim NumeroRecord As Long
    
    Testo = "Sei sicuro di voler eliminare la riga?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione verifica") = vbNo Then Exit Sub
    
    
    NumeroRecord = Me.GrigliaVerifica.ListIndex - 1

    m_DocumentsLink3.Delete
    
    OnSave
    
    If Not (Me.GrigliaVerifica.Recordset.EOF And Me.GrigliaVerifica.Recordset.BOF) Then
        Me.GrigliaVerifica.Recordset.Move NumeroRecord
    End If
End Sub

Private Sub cmdLottoArticoloSeme_Click()
    
    frmRicerca.Show vbModal
    
End Sub

Private Sub cmdLottoSeme_Click()
    If Me.Frame5.Visible = True Then
        Me.Frame5.Visible = False
    Else
        Me.Frame5.Visible = True
    End If
End Sub

Private Sub cmdNuovaCartella_Click()
    frmNuovaCartella.Show vbModal
    Me.DirSelezionato.Path = Me.txtPercorsoSelezionato.Text
    Me.DirSelezionato.Refresh
End Sub

Private Sub cmdNuovo_Quadratura_Click()

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire i prodotti derivati", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If Me.txtIDVarieta.Value > 0 Then
    If m_DocumentsLink.TableNew Then
        m_DocumentsLink.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink.NewRow
Else
    MsgBox "Bisogna inserire la varietà del lotto per poter gestire gli articoli", vbInformation, "Inserimento dati"
End If
    
    
Me.CDArticolo.SetFocus
End Sub

Private Sub cmdNuovoOrdiniAssociato_Click()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di inserire i collegamenti agli ordini cliente", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
    Link_LottoCampagna = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    frmRicercaOrdini.Show vbModal
    GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    OnSave
    
    GET_RIEPILOGO_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

End Sub

Private Sub cmdNuovoSemi_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire i lotti utilizzati", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If m_DocumentsLink2.TableNew Then
    m_DocumentsLink2.AbortNewRow
End If
'Crea una nuova riga vuota nel buffer
m_DocumentsLink2.NewRow

Me.CDArticoloSemi.SetFocus

End Sub

Private Sub cmdNuovoSfalcio_Click()
On Error GoTo ERR_cmdNuovoSfalcio_Click
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim AvviaProcesso As Boolean

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire lo sfalcio", vbInformation, "Salvataggio documento"
    Exit Sub
End If


''''CONTROLLO PER INSERIMENTO AUTOMATICO'''''''''''''''''''''''''''''''''''''
AvviaProcesso = False
sSQL = "SELECT AttivaGestioneSfalci FROM RV_PO01_FamigliaProdotti "
sSQL = sSQL & "WHERE IDRV_PO01_FamigliaProdotti=" & Me.cboFamigliaProdotti.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!AttivaGestioneSfalci) = 1 Then
        AvviaProcesso = True
    End If
Else
    AvviaProcesso = False
End If

If AvviaProcesso = False Then Exit Sub

If m_DocumentsLink4.TableNew Then
    m_DocumentsLink4.AbortNewRow
End If
'Crea una nuova riga vuota nel buffer
m_DocumentsLink4.NewRow

Exit Sub
ERR_cmdNuovoSfalcio_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovoSfalcio_Click"
End Sub

Private Sub cmdNuovoTerreno_Click()

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire le serre\appezzamenti", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If Me.cboFamigliaProdotti.CurrentID > 0 Then
    If m_DocumentsLink1.TableNew Then
        m_DocumentsLink1.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink1.NewRow
   
    
Else
    MsgBox "Inserire la famiglia prodotti del lotto di riferimento", vbInformation, "Salvataggio righe"
End If

Me.CDSerre.SetFocus
End Sub



Private Sub cmdNuovoVerifica_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire le verifiche", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If m_DocumentsLink3.TableNew Then
    m_DocumentsLink3.AbortNewRow
End If
'Crea una nuova riga vuota nel buffer
m_DocumentsLink3.NewRow

Me.txtDataVerifica.SetFocus
End Sub

Private Sub cmdProgressivo_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If fnNotNullN(m_Document("NumeroSbloccoLotto")) = 0 Then
    sSQL = "SELECT MAX(NumeroSbloccoLotto) as NumeroRecord "
    sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.txtProgressivo.Value = 0
    Else
        Me.txtProgressivo.Value = fnNotNullN(rs!NumeroRecord) + 1
    End If

    rs.CloseResultset
    Set rs = Nothing
End If


End Sub

Private Sub cmdProtocollo_Click()
If ((Me.CDSocio.KeyFieldID > 0) And (Me.cboFamigliaProdotti.CurrentID > 0)) Then
    frmProtocollo.Show vbModal
    Me.SetFocus
End If
End Sub

Private Sub cmdRipristina_Click()
    Me.txtPercorsoDocumentazione.Text = fnNotNull(m_Document("PercorsoDocumentazione").Value)
End Sub

Private Sub cmdSalva_Quadratura_Click()
'On Error GoTo ERR_cmdSalva_Quadratura_Click
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long
Dim sSQL As String
Dim Testo As String

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire i prodotti derivati", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If GET_CONTROLLO_AZIONI_UTENTE(m_DocType.ID, TheApp.IDUser, TheApp.Branch, 3) = False Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Non si hanno i diritti per eseguire questo comando"
    
    MsgBox Testo, vbCritical, "Controllo azioni utente"
    Exit Sub
    
End If

SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
APERTURA_FORM_CODA = False

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents


'''''''''''''''SALVATAGGIO RIGA''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    m_DocumentsLink("IDArticolo").Value = Me.CDArticolo.KeyFieldID
    m_DocumentsLink("Calibro").Value = Trim(Me.txtCalibro.Text)
    m_DocumentsLink("QuantitaPresunta").Value = Me.txtQtaPresunta.Value
    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then
        m_DocumentsLink("IDUtenteInserimento").Value = TheApp.IDUser
        m_DocumentsLink("PCInserimento").Value = GET_NOMECOMPUTER
        m_DocumentsLink("UtentePCInserimento").Value = GET_NOMEUTENTE
        m_DocumentsLink("DataInserimento").Value = Date
        m_DocumentsLink("OraInserimento").Value = GET_ORARIO(Now)
    Else
        m_DocumentsLink("IDUtenteUltimaModifica").Value = TheApp.IDUser
        m_DocumentsLink("PCUltimaModifica").Value = GET_NOMECOMPUTER
        m_DocumentsLink("UtentePCUltimaModifica").Value = GET_NOMEUTENTE
        m_DocumentsLink("DataUltimaModifica").Value = Date
        m_DocumentsLink("OraUltimaModifica").Value = GET_ORARIO(Now)
    End If



    m_DocumentsLink.Save
    
    If GET_ARTICOLO_GESTIONE_LOTTI(Me.CDArticolo.KeyFieldID) = True Then
        If GET_ESISTENZA_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, Me.CDArticolo.KeyFieldID) = False Then
            fnGetLottoArticolo Me.CDArticolo.KeyFieldID
        End If
        If MOVIMENTAZIONE_ARTICOLO(Me.CDArticolo.KeyFieldID, Me.CDArticolo.Description, fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value), Me.txtQtaPresunta.Value, Me.txtDataNascita.Text) = False Then
            MsgBox "Impossibile movimentare la riga dell'articolo", vbCritical, "Movimentazione articolo"
        End If
    End If
    
    
    m_DocumentsLink.Move Me.Griglia.ListIndex - 1
    
    'OnSave
    
    'If Not (BrwMain.Visible) Then Change
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display(False)
    Cn.CursorLocation = OLD_Cursor

    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    OnSave
    
Exit Sub
ERR_cmdSalva_Quadratura_Click:
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)
End Sub

Private Sub cmdSalvaComeNuovo_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDLottoCampagna_local As Long
Dim NumeroGiorniDifferenza As Integer


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

LINK_LOTTO_CAMPAGNA_SCN = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

frmSalvaComeNuovo.Show vbModal

If SALVA_COME_NUOVO = False Then Exit Sub


NewRecord

sSQL = "SELECT * FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & LINK_LOTTO_CAMPAGNA_SCN

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.CDSocio.Load fnNotNullN(rs!IDSocio)
    Me.cboFamigliaProdotti.WriteOn fnNotNullN(rs!IDRV_PO01_FamigliaProdotti)
    Me.txtIDVarieta.Value = fnNotNullN(rs!IDRV_PO01_Varieta)
    Me.txtVarieta.Text = GET_CODICE_TABELLA("RV_PO01_Varieta", "Varieta", "IDRV_PO01_Varieta", Me.txtIDVarieta.Value)
    
    Me.cboPeriodoCampagna.WriteOn RIPORTA_LINK_PERIODO_CAMPAGNA
    Me.cboTipoProduzione.WriteOn fnNotNullN(rs!IDRV_PO01_TipoProduzione)
    Me.cboStatoLotto.WriteOn RIPORTA_LINK_STATO_LOTTO
    
    Me.txtDataSemina.Text = DATA_SEMINA_SCN
    txtDataSemina_LostFocus
    NumeroGiorniDifferenza = DateDiff("d", fnNotNull(rs!DataSemina), DATA_SEMINA_SCN)
    
    If Me.txtDataInizioProduzione.Value = 0 Then
        Me.txtDataInizioProduzione.Text = DateAdd("d", NumeroGiorniDifferenza, fnNotNull(rs!DataInizioProduzione))
    End If
    If Me.txtDataFineProduzione.Value = 0 Then
        Me.txtDataFineProduzione.Text = DateAdd("d", NumeroGiorniDifferenza, fnNotNull(rs!DataFineProduzione))
    End If



    'Me.txtDataGerm.Text = fnNotNull(rs!DataGerminazionePresunta)
    'Me.txtDataNascita.Text = fnNotNull(rs!DataNascitaPresunta)
    'Me.txtDataTrapianto.Text = fnNotNull(rs!DataTrapiantoPresunta)
    'Me.txtDataRipic.Text = fnNotNull(rs!DataRipicchettaggioPresunta)
    
    'Me.txtDataSeminaEffettiva.Text = fnNotNull(rs!DataSeminaEffettiva)
    'Me.txtDataInizioProduzioneEffettiva.Text = fnNotNull(rs!DataInizioProduzioneEffettiva)
    'Me.txtDataFineProduzioneEffettiva.Text = fnNotNull(rs!DataFineProduzioneEffettiva)
    'M'e.txtDataGermEffettiva.Text = fnNotNull(rs!DataGerminazioneEffettiva)
    'Me.txtDataNascitaEffettiva.Text = fnNotNull(rs!DataNascitaEffettiva)
    'Me.txtDataTrapiantoEffettiva.Text = fnNotNull(rs!DataTrapiantoEffettiva)
    'Me.txtDataRipicEffettiva.Text = fnNotNull(rs!DataRipicchettaggioEffettiva)
    
    Me.txtDescrizioneLotto.Text = fnNotNull(rs!DescrizioneLotto)
    Me.CDCertificazione.Load fnNotNullN(rs!IDRV_PO01_CertificazioneSocio)
    Me.cboSchema.WriteOn fnNotNullN(rs!IDRV_PO01_Schema)
    Me.txtNumeroPianteMQ.Value = fnNotNullN(rs!NumeroPiantaMQ)
    Me.txtProtocolloCertificazione.Text = fnNotNull(rs!ProtocolloCertificazione)

End If

rs.CloseResultset
Set rs = Nothing




End Sub

Private Sub cmdSalvaSemi_Click()
On Error GoTo ERR_cmdSalva_Quadratura_Click
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long
Dim sSQL As String


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire i semi utilizzati", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If Me.CDArticoloSemi.KeyFieldID = 0 Then
    MsgBox "Inserire l'articolo", vbInformation, "Salvataggio lotti semi"
    Me.CDArticolo.SetFocus
    Exit Sub
End If


SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
APERTURA_FORM_CODA = False

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents


''''''''''''''SALVATAGGIO DELLA RIGA DELLA SERRA'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    m_DocumentsLink2("IDArticolo").Value = Me.CDArticoloSemi.KeyFieldID
    m_DocumentsLink2("IDLottoArticolo").Value = Me.txtIDLottoArticolo.Value
    m_DocumentsLink2("Quantita").Value = Me.txtQuantitaSemi.Value
    
    m_DocumentsLink2.Save
    
    m_DocumentsLink2.Move Me.GrigliaSemi.ListIndex - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display(False)
    Cn.CursorLocation = OLD_Cursor

    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    OnSave
Exit Sub
ERR_cmdSalva_Quadratura_Click:
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)




    'If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdSalvaSfalcio_Click()
On Error GoTo ERR_cmdSalvaSfalcio_Click
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire lo sfalcio", vbInformation, "Validazione dati"
    Exit Sub
End If

If Me.cboSfalcio.CurrentID = 0 Then
    MsgBox "Inserire lo sfalcio", vbInformation, "Validazione dati"
    Exit Sub
End If
If (m_DocumentsLink4(m_DocumentsLink4.PrimaryKey) <= 0) Then
    If GET_CONTROLLO_ESISTENZA_SFALCIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.cboSfalcio.CurrentID) = True Then
        MsgBox "Sfalcio già inserito per questo lotto di produzione", vbInformation, "Validazione dati"
        Exit Sub
    End If
End If

m_DocumentsLink4("IDRV_PO01_Sfalcio").Value = Me.cboSfalcio.CurrentID
If (Me.txtDataPresSfalcio.Value = 0) Then
    m_DocumentsLink4("DataPresuntaInizio").Value = Null
Else
    m_DocumentsLink4("DataPresuntaInizio").Value = Me.txtDataPresSfalcio.Value
End If
If Me.txtDataEffSfalcio.Value = 0 Then
     m_DocumentsLink4("DataEffettivaInizio").Value = Null
Else
    m_DocumentsLink4("DataEffettivaInizio").Value = Me.txtDataEffSfalcio.Value
End If

If (m_DocumentsLink4(m_DocumentsLink4.PrimaryKey) <= 0) Then
    m_DocumentsLink4("Guid").Value = GetGUID
End If
m_DocumentsLink4.Save
    
m_DocumentsLink4.Move Me.GrigliaSfalcio.ListIndex - 1
Exit Sub
ERR_cmdSalvaSfalcio_Click:
    MsgBox Err.Description, vbCritical, "cmdSalvaSfalcio_Click"
End Sub

Private Sub cmdSalvaTerreno_Click()
On Error GoTo ERR_cmdSalva_Quadratura_Click
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long
Dim sSQL As String
Dim NumeroRecord As Long
Dim Testo As String


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire le serre\appezzamenti", vbInformation, "Salvataggio documento"
    Exit Sub
End If
If GET_CONTROLLO_AZIONI_UTENTE(m_DocType.ID, TheApp.IDUser, TheApp.Branch, 3) = False Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Non si hanno i diritti per eseguire questo comando"
    
    MsgBox Testo, vbCritical, "Controllo azioni utente"
    Exit Sub
    
End If

SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
APERTURA_FORM_CODA = False

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents
    
    NumeroRecord = Me.GrigliaTerreni.ListIndex - 1
    
''''''''''''''SALVATAGGIO DELLA RIGA DELLA SERRA'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    m_DocumentsLink1("IDRV_PO01_Serra").Value = Me.CDSerre.KeyFieldID
    m_DocumentsLink1("DimensioneMQ").Value = Me.txtSuperficieMQ_Serra.Value
    m_DocumentsLink1("DimensioneHA").Value = Me.txtSuperficieHA_Serra.Text
    If (m_DocumentsLink1(m_DocumentsLink1.PrimaryKey) <= 0) Then
        m_DocumentsLink1("GuidID").Value = GetGUID
    End If
    m_DocumentsLink1.Save
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display(False)
    Cn.CursorLocation = OLD_Cursor

    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    OnSave
    
    AGGIORNA_ORDINI_DA_SERRE Me.txtCodiceLotto.Text, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    m_DocumentsLink1.Move NumeroRecord
Exit Sub
ERR_cmdSalva_Quadratura_Click:
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)




    'If Not (BrwMain.Visible) Then Change

End Sub

Private Sub cmdSalvaVerifica_Click()
On Error GoTo ERR_cmdSalva_Quadratura_Click
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long
Dim sSQL As String


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di inserire le verifiche del lotto", vbInformation, "Salvataggio documento"
    Exit Sub
End If




SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
APERTURA_FORM_CODA = False

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
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
        sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents


''''''''''''''SALVATAGGIO DELLA RIGA DELLA SERRA'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.txtDataVerifica.Value > 0 Then
        m_DocumentsLink3("DataVerifica").Value = Me.txtDataVerifica.Value
    End If
    m_DocumentsLink3("OraVerifica").Value = Me.txtOraVerifica.Text
    m_DocumentsLink3("Annotazioni").Value = Me.txtAnnotazioniVerifica.Text
    m_DocumentsLink3("Operatore").Value = Me.txtOperatoreVerifica.Text
    m_DocumentsLink3("TipoRilevazione").Value = Me.txtTipoRilevazione.Text
    m_DocumentsLink3("Temperatura").Value = Me.txtTemperaturaVerifica.Text
    
    
    
    m_DocumentsLink3.Save
    
    m_DocumentsLink3.Move Me.GrigliaVerifica.ListIndex - 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display(False)
    Cn.CursorLocation = OLD_Cursor

    ''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    OnSave
Exit Sub
ERR_cmdSalva_Quadratura_Click:
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Salvataggio verifiche lotto"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)

    'If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdSelezionaVarieta_Click()
Dim Link_Lotto As Long
On Error GoTo ERR_cmdSelezionaVarieta_Click

    frmSelezionaVarieta.Show vbModal
    
    If Me.txtIDVarieta.Value > 0 Then
        GET_DATI_VARIETA Me.txtIDVarieta.Value
    End If
    
    Me.SetFocus
Exit Sub

ERR_cmdSelezionaVarieta_Click:
    MsgBox Err.Description, vbCritical, "cmdSelezionaVarieta_Click"
End Sub

Private Function GET_DATI_VARIETA(IDVarieta As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_PO01_Varieta.IDRV_PO01_Varieta, RV_PO01_Varieta.Varieta, RV_PO01_Varieta.IDRV_PO01_FamigliaProdotti, "
sSQL = sSQL & "RV_PO01_Varieta.CodiceImEx, RV_PO01_FamigliaProdotti.FamigliaProdotti"
sSQL = sSQL & "FROM RV_PO01_Varieta LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_FamigliaProdotti ON "
sSQL = sSQL & "RV_PO01_Varieta.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti "
sSQL = sSQL & "WHERE IDRV_PO01_Varieta=" & IDVarieta

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
'    LINK_VARIETA_LOTTO_CAMPAGNA = 0
'    LINK_FAMIGLIA_LOTTO_CAMPAGNA = 0
'    LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA = 0
'    VARIETA_LOTTO_CAMPAGNA = ""
'    FAMIGLIA_LOTTO_CAMPAGNA = ""
'    TIPO_PRODUZIONE_LOTTO_CAMPAGNA = ""
'    DATA_SBLOCCO_LOTTO_CAMPAGNA = ""
Else
'    LINK_VARIETA_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_Varieta)
 '   LINK_FAMIGLIA_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_FamigliaProdotti)
 '   LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_TipoProduzione)
 '   VARIETA_LOTTO_CAMPAGNA = fnNotNull(rs!Varieta)
 '   FAMIGLIA_LOTTO_CAMPAGNA = fnNotNull(rs!FamigliaProdotti)
 '   TIPO_PRODUZIONE_LOTTO_CAMPAGNA = fnNotNull(rs!TipoProduzione)
 '   DATA_SBLOCCO_LOTTO_CAMPAGNA = fnNotNull(rs!DataSbloccoLotto)
End If

rs.CloseResultset
Set rs = Nothing


'Me.txtVarietà.Text = VARIETA_LOTTO_CAMPAGNA
'Me.txtFamiglia.Text = FAMIGLIA_LOTTO_CAMPAGNA
'Me.txtTipoProduzione.Text = TIPO_PRODUZIONE_LOTTO_CAMPAGNA
'Me.txtDataSbloccoLotto.Text = DATA_SBLOCCO_LOTTO_CAMPAGNA
End Function



Private Sub cmdSerreDisponibili_Click()
Dim Testo As String

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di inserire le serre/appezzamenti", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    If fnNotNullN(m_Document("Provvisorio").Value) = 1 Then
        MsgBox "Comando non utilizzabile quando un lotto risulta provvisorio", vbInformation, "Controllo dati"
        Exit Sub
    End If
    If GET_CONTROLLO_AZIONI_UTENTE(m_DocType.ID, TheApp.IDUser, TheApp.Branch, 3) = False Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo = "Non si hanno i diritti per eseguire questo comando"
        
        MsgBox Testo, vbCritical, "Controllo azioni utente"
        Exit Sub
    End If
    
    VarIDAzienda = m_App.IDFirm
    VarIDFiliale = m_App.Branch
    VarIDUtente = m_App.IDUser
    Link_LottoCampagna = m_Document(m_Document.PrimaryKey).Value
    Link_Socio = Me.CDSocio.KeyFieldID

    Me.cboSchema.WriteOn fnNotNullN(m_Document("IDRV_PO01_Schema").Value)

    CONFERMA_SELEZIONE_SERRE = False
    
    frmConfigurazioneSerre.Show vbModal
    
    If CONFERMA_SELEZIONE_SERRE = True Then
        Me.txtSuperficieMQ.Value = GET_TOTALE_SERRE
        txtSuperficieMQ_LostFocus
        AggiornamentoGriglia = 1
        Me.cboSchema.WriteOn Link_Schema
        AggiornamentoGriglia = 0
        m_DocumentsLink1.Refresh
        OnSave
        AGGIORNA_ORDINI_DA_SERRE Me.txtCodiceLotto.Text, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    Else
        Link_Schema = fnNotNullN(m_Document("IDRV_PO01_Schema").Value)
    End If

    
End Sub
Private Function GET_TOTALE_SERRE() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(RV_PO01_SettoreSerra.DimensioneMQ) AS Totale "
sSQL = sSQL & "FROM RV_PO01_SettoreSerra INNER JOIN "
sSQL = sSQL & "RV_PO01_SettoreSchema ON RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema = RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema INNER JOIN "
sSQL = sSQL & "RV_PO01_LottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_SerraPerLotto ON RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna = RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna ON "
sSQL = sSQL & "RV_PO01_SettoreSchema.IDRV_PO01_Schema = RV_PO01_LottoCampagna.IDRV_PO01_Schema "
sSQL = sSQL & "WHERE (RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna = " & fnNotNullN(m_Document(m_Document.PrimaryKey).Value) & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_SERRE = 0
Else
    GET_TOTALE_SERRE = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub cmdTrovaPercorso_Click()
 
    frmPercorso.Show vbModal

End Sub


Private Sub cmdUtente_Click()
If Me.Frame3.Height = 3375 Then
    Me.Frame3.Height = 6015
Else
    Me.Frame3.Height = 3375
End If
End Sub

Private Sub cmdVisUtenteMod_Click()
    If Me.FraModProdotti.Visible = False Then
        Me.Griglia.Top = 2520
        Me.Griglia.Height = 1815
        Me.FraModProdotti.Visible = True
    Else
        Me.Griglia.Top = 1200
        Me.Griglia.Height = 3135
        Me.FraModProdotti.Visible = False
    
    End If
End Sub

Private Sub DirSelezionato_Change()
    Me.FileDocumentazione.Path = Me.DirSelezionato.Path
    Me.FileDocumentazione.Refresh
    
    Me.txtPercorsoSelezionato.Text = Me.DirSelezionato.Path
End Sub

Private Sub FileDocumentazione_DblClick()
On Error GoTo ERR_FileDocumentazione_DblClick
Dim Scr_hDC As Long
Dim X
Dim msg As String
Scr_hDC = GetDesktopWindow()

X = ShellExecute(Scr_hDC, "Open", Me.FileDocumentazione.FileName, "", Me.txtPercorsoSelezionato.Text, SW_SHOWNORMAL)

If X <= 32 Then
    'There was an error
    Select Case X
        Case SE_ERR_FNF
            msg = "File not found"
        Case SE_ERR_PNF
            msg = "Path not found"
        Case SE_ERR_ACCESSDENIED
            msg = "Access denied"
        Case SE_ERR_OOM
            msg = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            msg = "DLL not found"
        Case SE_ERR_SHARE
            msg = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            msg = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            msg = "DDE Time out"
        Case SE_ERR_DDEFAIL
            msg = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            msg = "DDE busy"
        Case SE_ERR_NOASSOC
            msg = "No association for file extension"
        Case ERROR_BAD_FORMAT
            msg = "Invalid EXE file or error in EXE image"
        Case Else
            msg = "Unknown error"
    End Select
    
    MsgBox msg, vbInformation, "Apertura file"

End If

Exit Sub

ERR_FileDocumentazione_DblClick:
    MsgBox Err.Description, vbCritical, "Apertura file"

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
    DMTSplitBar1.ZOrder 0
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
    
    On Error Resume Next
    
    

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
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto, IDIvaAcquisto "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm
        
    
    
    
    'Assegnazione della query di ricerca
    oSearch.SQL = fnAnsi2Jet(sSQL)
    
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
   
    
    Set oRes = oSearch.Exec
    
    
    If Not oRes.EOF Then
        Me.CDArticolo.Load oRes!IDArticolo
        Me.CDArticolo.Code = oRes!CodiceArticolo
        Me.TxtArticolo.Text = oRes!Articolo
    End If
            
End If
    
    Set oRes = Nothing
    Set oSearch = Nothing

End Sub

Private Sub m_DocumentsLink1_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    
    
    

    If Not (m_DocumentsLink1.BOF And m_DocumentsLink1.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        Me.CDSerre.Load fnNotNullN(m_DocumentsLink1("IDRV_PO01_Serra").Value)
        Me.txtSuperficieHA_Serra.Text = fnNotNull(m_DocumentsLink1("DimensioneHA").Value)
        Me.txtSuperficieMQ_Serra.Value = fnNotNullN(m_DocumentsLink1("DimensioneMQ").Value)
        bValue = True
             
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.CDSerre.Load 0
        Me.txtSuperficieHA_Serra.Text = ""
        Me.txtSuperficieMQ_Serra.Text = 0

        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        'Me.cboCausaleQuadratura.Enabled = bValue
        
        
        Me.CDSerre.Enabled = bValue
        Me.txtSuperficieHA_Serra.Enabled = bValue
        Me.txtSuperficieMQ_Serra.Enabled = bValue
        
       
 
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoTerreno.Enabled = True
        Me.cmdSalvaTerreno.Enabled = bValue
        Me.cmdEliminaTerreno.Enabled = bValue



End Sub

Private Sub m_DocumentsLink2_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        Me.CDArticoloSemi.Load fnNotNullN(m_DocumentsLink2("IDArticolo").Value)
        Me.txtIDLottoArticolo.Value = fnNotNullN(m_DocumentsLink2("IDLottoArticolo").Value)
        Me.txtQuantitaSemi.Value = fnNotNullN(m_DocumentsLink2("Quantita").Value)
        
        
        bValue = True
             
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.CDArticoloSemi.Load 0
        Me.txtIDLottoArticolo.Value = 0
        Me.txtQuantitaSemi.Value = 0

        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    

        Me.CDArticoloSemi.Enabled = bValue
        Me.txtIDLottoArticolo.Enabled = bValue
        Me.txtQuantitaSemi.Enabled = bValue
        Me.lblLottoSeme.Enabled = bValue
       
 
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoSemi.Enabled = True
        Me.cmdSalvaSemi.Enabled = bValue
        Me.cmdEliminaSemi.Enabled = bValue

End Sub

Private Sub m_DocumentsLink3_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink3.BOF And m_DocumentsLink3.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.

        Me.txtDataVerifica.Value = fnNotNullN(m_DocumentsLink3("DataVerifica").Value)
        Me.txtOraVerifica.Text = fnNotNull(m_DocumentsLink3("OraVerifica").Value)
        Me.txtAnnotazioniVerifica.Text = fnNotNull(m_DocumentsLink3("Annotazioni").Value)
        Me.txtOperatoreVerifica.Text = fnNotNull(m_DocumentsLink3("Operatore").Value)
        Me.txtTipoRilevazione.Text = fnNotNull(m_DocumentsLink3("TipoRilevazione").Value)
        Me.txtTemperaturaVerifica.Text = fnNotNull(m_DocumentsLink3("Temperatura").Value)
        
            
        bValue = True
             
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.txtDataVerifica.Value = 0
        Me.txtOraVerifica.Text = ""
        Me.txtAnnotazioniVerifica.Text = ""
        Me.txtOperatoreVerifica.Text = ""
        Me.txtTipoRilevazione.Text = ""
        Me.txtTemperaturaVerifica.Text = ""

        bValue = False
    End If
    
        Me.txtDataVerifica.Enabled = bValue
        Me.txtOraVerifica.Enabled = bValue
        Me.txtAnnotazioniVerifica.Enabled = bValue
        Me.txtOperatoreVerifica.Enabled = bValue
        Me.txtTipoRilevazione.Enabled = bValue
        Me.txtTemperaturaVerifica.Enabled = bValue
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoVerifica.Enabled = True
        Me.cmdSalvaVerifica.Enabled = bValue
        Me.cmdEliminaVerifica.Enabled = bValue
End Sub

Private Sub m_DocumentsLink4_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink4.BOF And m_DocumentsLink4.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.

        Me.cboSfalcio.WriteOn fnNotNullN(m_DocumentsLink4("IDRV_PO01_Sfalcio").Value)
        Me.txtDataPresSfalcio.Value = fnNotNullN(m_DocumentsLink4("DataPresuntaInizio").Value)
        Me.txtDataEffSfalcio.Value = fnNotNullN(m_DocumentsLink4("DataEffettivaInizio").Value)

        
            
        bValue = True
             
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.cboSfalcio.WriteOn 0
        Me.txtDataPresSfalcio.Value = 0
        Me.txtDataEffSfalcio.Value = 0

        bValue = False
    End If
    
    Me.cboSfalcio.Enabled = bValue
    Me.txtDataPresSfalcio.Enabled = bValue
    Me.txtDataEffSfalcio.Enabled = bValue
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoSfalcio.Enabled = True
        Me.cmdSalvaSfalcio.Enabled = bValue
        Me.cmdEliminaSfalcio.Enabled = bValue
End Sub

Private Sub TxtCodiceImEx_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtCodiceLotto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtCodiceLotto_GotFocus()
    Me.txtCodiceLotto.SelStart = 0
    Me.txtCodiceLotto.SelLength = Len(Me.txtCodiceLotto.Text)
End Sub
Private Sub txtCodiceLottoSeme_Change()
    If Not (BrwMain.Visible) Then Change

End Sub

Private Sub txtDataFineProduzione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataFineProduzione_LostFocus()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End If
End Sub

Private Sub txtDataFineProduzioneEffettiva_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataGerm_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataGermEffettiva_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataImportazione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataInizioProduzione_Change()


If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataInizioProduzione_LostFocus()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End If
End Sub

Private Sub txtDataInizioProduzioneEffettiva_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataNascita_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataNascitaEffettiva_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataRipic_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataRipicEffettiva_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataSbloccoLotto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataSemina_Change()
On Error GoTo ERR_txtDataSemina_Change
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
    End If
    
    
    
    If Not (BrwMain.Visible) Then Change
Exit Sub
ERR_txtDataSemina_Change:
    MsgBox Err.Description, vbCritical, "txtDataSemina_Change"
End Sub

Private Sub txtDataSemina_LostFocus()
On Error GoTo ERR_txtDataSemina_LostFocus
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
        If Me.txtDataSemina.Value > 0 Then
            GET_CALCOLO_DATE_DOPO_SEMINA Me.cboFamigliaProdotti.CurrentID, Me.txtDataSemina.Text
        End If
    Else
        If Me.txtDataSemina <> fnNotNull(m_Document("DataSemina").Value) Then
            If MsgBox("Vuoi ricalcolare le date?", vbQuestion + vbYesNo, "Ricalcolo date per famiglia") = vbNo Then Exit Sub
            GET_CALCOLO_DATE_DOPO_SEMINA Me.cboFamigliaProdotti.CurrentID, Me.txtDataSemina.Text
        End If
    End If
Exit Sub

ERR_txtDataSemina_LostFocus:
    MsgBox Err.Description, vbCritical, "txtDataSemina_LostFocus"
End Sub

Private Sub txtDataSeminaEffettiva_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataTrapianto_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataTrapiantoEffettiva_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneConcia_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneLotto_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneLotto_GotFocus()
    Me.txtDescrizioneLotto.SelStart = 0
    Me.txtDescrizioneLotto.SelLength = Len(Me.txtDescrizioneLotto.Text)

End Sub



Private Sub txtDescrizioneLottoSeme_Change()
    If Not (BrwMain.Visible) Then Change

End Sub

Private Sub txtIDLottoArticolo_Change()
On Error GoTo ERR_txtIDLottoArticolo_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM LottoArticolo "
sSQL = sSQL & "WHERE IDLottoArticolo=" & Me.txtIDLottoArticolo.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtCodiceLottoSemeRiga.Text = ""
    Me.txtDescrizioneLottoSemeRiga.Text = ""
Else
    Me.txtCodiceLottoSemeRiga.Text = fnNotNull(rs!Codice)
    Me.txtDescrizioneLottoSemeRiga.Text = fnNotNull(rs!LottoArticolo)
End If


rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_txtIDLottoArticolo_Change:
    MsgBox Err.Description, vbCritical, "txtIDLottoArticolo_Change"
    
End Sub

Private Sub txtIDVarieta_Change()
    Me.txtCodiceVarieta.Text = GET_CODICE_TABELLA("RV_PO01_Varieta", "CodiceImEx", "IDRV_PO01_Varieta", Me.txtIDVarieta.Value)
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNPassaportoProduttore_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroPianteMQ_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPercorsoDocumentazione_Change()
On Error Resume Next
    If Me.txtPercorsoDocumentazione.Text = "" Then
        Me.FileDocumentazione.Visible = False
        Me.DirSelezionato.Visible = False
    Else
        Me.FileDocumentazione.Visible = True
        Me.DirSelezionato.Visible = True
    End If
    
    
    Me.DirSelezionato.Path = Me.txtPercorsoDocumentazione.Text
    Me.DirSelezionato.Refresh
    
    'Me.FileDocumentazione.Path = Me.txtPercorsoDocumentazione.Text

    
    If Not (BrwMain.Visible) Then Change
    
    
End Sub



Private Sub EliminaRiferimentiLottoSuperficie()
Dim sSQL As String

sSQL = "DELETE FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Cn.Execute sSQL

m_DocumentsLink.Refresh
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
    
    
    
    On Error Resume Next
        
    ParametroLinkSocio
    Link_Schema = fnNotNullN(m_Document("IDRV_PO01_Schema").Value)





    If fnNotNullN(m_Document("IDSocio").Value) > 0 Then
        
        cboFamigliaProdotti_Click
        
        With Me.CDSerre
            Set .Application = TheApp
            Set .Database = TheApp.Database
            .HwndContainer = Me.hwnd
            .CodeField = "Codice"
            .DescriptionField = "Descrizione"
            .KeyField = "IDRV_PO01_Serra"
            If (fnNotNullN(m_Document("Provvisorio").Value) = 0) Then
                .TableName = "RV_PO01_IESchemaSerra"
                .Filter = "IDRV_PO01_Schema=" & fnNotNullN(m_Document("IDRV_PO01_Schema").Value)
            Else
                .TableName = "RV_PO01_IESerra"
                .Filter = "IDSocio=" & fnNotNullN(m_Document("IDSocio").Value)
            End If
            .MenuFunctions("EseguiGestione").Enabled = True
            .PropCodice.Caption = "Codice"
            'Caption da associare alla label del campo Descrizione
            .PropDescrizione.Caption = "Descrizione"
            'Caption da associare alla intestazione della colonna della Find per il campo Codice
            .CodeCaption4Find = "Codice serra"
            'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
            .DescriptionCaption4Find = "Descrizione serra"
            'Identificativo della Funzione Diamante per l'Esegui Gestione
            .IDExecuteFunction = fnGetTipoOggetto("RV_PO01_Serre")
            'Indica se il campo Codice è un campo numerico
            .CodeIsNumeric = False
        End With
        
    End If

    GUID_LOTTOCAMPAGNA = ""
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        GET_GRIGLIA_ORDINE -1
        
    Else
        GET_GRIGLIA_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        GUID_LOTTOCAMPAGNA = fnNotNull(m_Document("GuidID").Value)
    End If
    

    Set Me.Griglia.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink.TableName).Data
    Set Me.GrigliaTerreni.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink1.TableName).Data
    Set Me.GrigliaSemi.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink2.TableName).Data
    Set Me.GrigliaVerifica.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink3.TableName).Data
    Set Me.GrigliaSfalcio.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink4.TableName).Data
    
    
    ContDelete = 0


    
    GET_RIEPILOGO_ORDINE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    

    
    Me.txtVarieta.Text = fnNotNull(m_Document("Varieta").Value)
       
    If fnNotNullN(m_Document("IDSocio").Value) > 0 Then
        GET_CERTIFICAZIONE_SOCIO fnNotNullN(m_Document("IDSocio").Value)
    End If
    
    If fnNotNullN(m_Document("IDRV_PO01_CertificazioneSocio").Value) > 0 Then
        Me.txtProtocolloCertificazione.Text = GET_PROTOCOLLO_CERTIFICAZIONE(fnNotNullN(m_Document("IDRV_PO01_CertificazioneSocio").Value))
    End If
    
    If Me.txtPercorsoDocumentazione.Text = "" Then
        Me.FileDocumentazione.Visible = False
    Else
        Me.FileDocumentazione.Visible = True
    End If
        
        

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
    
If AggiornamentoGriglia = 0 Then
    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
        Me.CDArticolo.Load fnNotNullN(m_DocumentsLink("IDArticolo").Value)
        Me.txtCalibro.Text = fnNotNull(m_DocumentsLink("Calibro").Value)
        Me.txtQtaPresunta.Value = fnNotNull(m_DocumentsLink("QuantitaPresunta").Value)
        
        Me.cboUtenteDMTInsArt.WriteOn fnNotNullN(m_DocumentsLink("IDUtenteInserimento").Value)
        Me.txtPCInsArt.Text = fnNotNull(m_DocumentsLink("PCInserimento").Value)
        Me.txtUtentePCInsArt.Text = fnNotNull(m_DocumentsLink("UtentePCInserimento").Value)
        Me.txtDataInsArt.Value = fnNotNullN(m_DocumentsLink("DataInserimento").Value)
        Me.txtOraInsArt.Text = fnNotNull(m_DocumentsLink("OraInserimento").Value)
        
        Me.cboUtenteDMTModArt.WriteOn fnNotNullN(m_DocumentsLink("IDUtenteUltimaModifica").Value)
        Me.txtPCModArt.Text = fnNotNull(m_DocumentsLink("PCUltimaModifica").Value)
        Me.txtUtentePCModArt.Text = fnNotNull(m_DocumentsLink("UtentePCUltimaModifica").Value)
        Me.txtDataModArt.Value = fnNotNullN(m_DocumentsLink("DataUltimaModifica").Value)
        Me.txtOraModArt.Text = fnNotNull(m_DocumentsLink("OraUltimaModifica").Value)
        

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
        Me.TxtArticolo.Text = ""
        Me.txtCalibro.Text = ""
        Me.txtQtaPresunta.Value = 0
        
        Me.cboUtenteDMTInsArt.WriteOn 0
        Me.txtPCInsArt.Text = ""
        Me.txtUtentePCInsArt.Text = ""
        Me.txtDataInsArt.Value = ""
        Me.txtOraInsArt.Text = ""
        
        Me.cboUtenteDMTModArt.WriteOn 0
        Me.txtPCModArt.Text = ""
        Me.txtUtentePCModArt.Text = ""
        Me.txtDataModArt.Value = ""
        Me.txtOraModArt.Text = ""
        
        
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        'Me.cboCausaleQuadratura.Enabled = bValue
        Me.CDArticolo.Enabled = bValue
        Me.TxtArticolo.Enabled = bValue
        Me.lblArticolo.Enabled = bValue
        Me.txtCalibro.Enabled = bValue
        Me.txtQtaPresunta.Enabled = bValue
       
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovo_Quadratura.Enabled = True
        Me.cmdSalva_Quadratura.Enabled = bValue
        Me.cmdElimina_Quadratura.Enabled = bValue

End If



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







Private Sub TxtArticolo_LostFocus()
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    Dim I As Long
    
 If Me.CDArticolo.Code = "" Then
    If Me.TxtArticolo.Text <> "" Then
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm
        sSQL = sSQL & " AND Articolo LIKE " & fnNormString(Me.TxtArticolo.Text & "%")

    
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
            sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto, IDIvaAcquisto "
            sSQL = sSQL & "FROM Articolo "
            sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm
            sSQL = sSQL & " AND Articolo LIKE " & fnNormString(Me.TxtArticolo.Text & "%")

            Set rs = Cn.OpenResultset(sSQL)
                Me.CDArticolo.Load rs!IDArticolo
                Me.CDArticolo.Code = rs!CodiceArticolo
                Me.TxtArticolo.Text = rs!Articolo
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
           
        End If
        
    End If
End If
    
End Sub
Private Function Parametro_TipoProdotto() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProdotto FROM RV_PO01_ParametriFiliale WHERE "
sSQL = sSQL & "IDAzienda=" & m_App.IDFirm & " AND "
sSQL = sSQL & "IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    Parametro_TipoProdotto = 0
Else
    Parametro_TipoProdotto = fnNotNullN(rs!IDProdotto)
End If
End Function
Private Function Parametro_MagazzinoProduzione() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDMagazzinoProduzione FROM RV_PO01_ParametriFiliale WHERE "
sSQL = sSQL & "IDAzienda=" & m_App.IDFirm & " AND "
sSQL = sSQL & "IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    Parametro_MagazzinoProduzione = 0
Else
    Parametro_MagazzinoProduzione = fnNotNullN(rs!IDMagazzinoProduzione)
End If
End Function

Private Function SalvataggioRighe() As Boolean
If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
    m_DocumentsLink.MoveFirst
    While Not m_DocumentsLink.EOF
        If m_DocumentsLink("IDRV_PO01_DettaglioLotto").Value < 0 Then
            If IsNull(m_DocumentsLink("IDArticolo").Value) Then
                m_DocumentsLink.DeleteRowFromBuffer
            Else
                Link_LottoArticolo = fnGetLottoArticolo(m_DocumentsLink("IDArticolo"))
            End If
        End If
            
            
        If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
            m_DocumentsLink.MoveNext
        End If
    Wend
End If



SalvataggioRighe = True

End Function

Private Function fnGetLottoArticolo(IDArticolo) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDLotto As Long

sSQL = "SELECT IDLottoArticolo FROM LottoArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo & " AND "
sSQL = sSQL & "Codice=" & fnNormString(Me.txtCodiceLotto.Text)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    IDLotto = fnGetNewKey("LottoArticolo", "IDLottoArticolo")
    
    sSQL = "INSERT INTO LottoArticolo (IDLottoArticolo, IDArticolo, Codice, LottoArticolo, Sospeso, DataScadenza) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDLotto & ", "
    sSQL = sSQL & m_DocumentsLink("IDArticolo").Value & ", "
    sSQL = sSQL & fnNormString(Me.txtCodiceLotto.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtDescrizioneLotto.Text) & ", "
    sSQL = sSQL & fnNormBoolean(0) & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataFineProduzione.Text) & ")"
    Cn.Execute sSQL
    
    fnGetLottoArticolo = IDLotto
    InserimentoLottoInMagazzino IDLotto
Else
    fnGetLottoArticolo = fnNotNullN(rs!IDLottoArticolo)
End If

rs.CloseResultset
Set rs = Nothing
    
    
    
End Function
Private Function InserimentoLottoInMagazzino(IDLotto As Long) As Boolean
'Restituisce True se tutte le operazioni sono andate a buon fine
'Altrimenti restiruisce False

    Dim sSQL As String
        
        sSQL = "INSERT INTO LottoArticoloPerMagazzino ("
        sSQL = sSQL & "IDLottoArticolo, IDMagazzino, Giacenza, DataUltimoCarico) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDLotto & ", "
        sSQL = sSQL & Me.cboMagazzinoCaricoLotto.CurrentID & ", "
        sSQL = sSQL & fnNormNumber(0) & ", "
        sSQL = sSQL & fnNormDate(Date) & ")"
        
        Cn.Execute sSQL
        InserimentoLottoInMagazzino = True
    
End Function


Private Function fnGetTipoOggetto(NomeGestore) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(NomeGestore)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function fnGetEsercizio(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio, DataInizio, DataFine FROM Esercizio WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (DataInizio <= " & fnNormDate(dData) & ")"
    sSQL = sSQL & " AND (DataFine >= " & fnNormDate(dData) & "))"
   

    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = fnNotNullN(rsEse!IDEsercizio)
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Sub TipoNumerazioneLotto()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ContLottoPeriodoCamp FROM RV_PO01_ParametriFiliale "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_TIPO_CONTATORE = fnNotNullN(rs!ContLottoPeriodoCamp)
Else
    LINK_TIPO_CONTATORE = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub ParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_PO01_ParametriFiliale WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & "))"


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSocio = fnNotNullN(rs!IDCategoriaAnagrafica)
Else
    Link_TipoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroAbilitaRicFor()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AbilitaRicercaFornitoriLottoCampagna, CodiceLottoDiProduzionePerSocio "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & 0


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IndicaRicFor = fnNotNullN(rs!AbilitaRicercaFornitoriLottoCampagna)
    LOTTO_PER_SOCIO = fnNotNullN(rs!CodiceLottoDiProduzionePerSocio)
Else
    IndicaRicFor = 0
    LOTTO_PER_SOCIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub ParametroSeme()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoSchedaSementi FROM RV_PO01_ParametriFiliale WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & "))"


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSchedaSeme = fnNotNullN(rs!IDTipoSchedaSementi)
Else
    Link_TipoSchedaSeme = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub ParametroLinkSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM RV_PO01_ParametriFiliale WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & "))"


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_Socio = fnNotNullN(rs!IDAnagrafica)
Else
    Link_Socio = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub ParametroTipoGestioneMovimentazione()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Narciso FROM RV_PO01_ParametriFiliale WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & "))"


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_TIPO_GESTIONE_MOVIMENTI = fnNotNullN(rs!Narciso)
Else
    LINK_TIPO_GESTIONE_MOVIMENTI = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function RaggruppamentoSocio(IDSocio As Long) As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsCTRL As ADODB.Recordset
Dim rsSocio As ADODB.Recordset

'Prelevo il numero dei soci nella licenza
sSQL = "SELECT NumeroSoci FROM RV_POComponenteTesta WHERE IDRV_POProgramma=" & IdentificativoProgramma
Set rsSocio = New ADODB.Recordset

rsSocio.Open sSQL, Cn.InternalConnection

If rsSocio.EOF Then
    NumeroSoci = 0
Else
    NumeroSoci = fnNotNullN(rsSocio!NumeroSoci)
End If

rsSocio.Close
Set rsSocio = Nothing




''''Controlla quanti soci sono stati inseriti nella creazione del lotto'''
sSQL = "SELECT IDSocio "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "GROUP BY IDSocio "

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic



If rs.EOF = False Then
    rs.MoveLast
    If rs.RecordCount < NumeroSoci Then
        RaggruppamentoSocio = True
    End If
    
'''Se il numero dei soci sono uguali a quelli inseriti allora si controlla se fra questi,
'''è presente quello inserito in questo record.
    If rs.RecordCount = NumeroSoci Then
        sSQL = "SELECT IDSocio "
        sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
        sSQL = sSQL & "WHERE "
        sSQL = sSQL & "IDSocio=" & IDSocio
        sSQL = sSQL & " GROUP BY IDSocio "
        
        Set rsCTRL = New ADODB.Recordset
        rsCTRL.Open sSQL, Cn.InternalConnection
        
        If rsCTRL.EOF Then
            'Se non esiste non si può salvare il documento
            RaggruppamentoSocio = False
        Else
            'Si può salvare il documento poichè fa parte del gruppo dei soci
            RaggruppamentoSocio = True
        End If
        
        rsCTRL.Close
        Set rsCTRL = Nothing
    End If
Else
    'Se non si è arrivato al numero si procede con le altre operazioni
    RaggruppamentoSocio = True
End If


rs.Close
Set rs = Nothing

    
End Function
Private Function ControlloEsistenzaCodiceLotto(CodiceLotto As String, Link_Record As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceLotto FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE CodiceLotto=" & fnNormString(CodiceLotto)
sSQL = sSQL & " AND IDRV_PO01_PeriodoCampagna=" & Me.cboPeriodoCampagna.CurrentID
sSQL = sSQL & " AND IDRV_PO01_LottoCampagna<>" & Link_Record
If LOTTO_PER_SOCIO = 1 Then
    sSQL = sSQL & " AND IDSocio=" & Me.CDSocio.KeyFieldID
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    ControlloEsistenzaCodiceLotto = False
Else
    ControlloEsistenzaCodiceLotto = True
End If

rs.CloseResultset
Set rs = Nothing
End Function



Private Sub txtProgressivo_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtProtocolloCertificazione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtSuperficieHA_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtSuperficieHAEffettiva_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtSuperficieMQ_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtSuperficieMQ_LostFocus()
Dim I As Integer
Dim X As String 'Stringa da formattare
Dim J As String '
Dim Ris As String
Const MAX_X As Integer = 10
Dim Rimanenza As Integer
Dim Valore As Long
Valore = Me.txtSuperficieMQ.Value

Rimanenza = MAX_X - Len(CStr(Valore))

For I = 1 To Rimanenza
    X = X & "0"
Next
    'Stringa da formattare
    X = X & CStr(Valore)
    J = ""
For I = 1 To Len(X)
    
    J = J & Mid(X, I, 1)
    If I < Len(X) Then
        If I Mod 2 = 0 Then
            J = J & "."
        End If
    End If
Next
Me.txtSuperficieHA.Text = J
End Sub
Private Function fncControlloMovimentazioneLotto() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_QuadernoDiCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna = " & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    fncControlloMovimentazioneLotto = False
Else
    fncControlloMovimentazioneLotto = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fncEliminaRiferimentiSerre(IDLottoCampagna As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Cn.Execute sSQL
End Sub

Private Sub txtSuperficieMQ_Serra_LostFocus()
Dim I As Integer
Dim X As String 'Stringa da formattare
Dim J As String '
Dim Ris As String
Const MAX_X As Integer = 10
Dim Rimanenza As Integer
Dim Valore As Long

Valore = Me.txtSuperficieMQ_Serra.Value

Rimanenza = MAX_X - Len(CStr(Valore))

For I = 1 To Rimanenza
    X = X & "0"
Next
    'Stringa da formattare
    X = X & CStr(Valore)
    J = ""
For I = 1 To Len(X)

    J = J & Mid(X, I, 1)
    If I < Len(X) Then
        If I Mod 2 = 0 Then
            J = J & "."
        End If
    End If
Next

Me.txtSuperficieHA_Serra.Text = J

'Me.txtSuperficieHA_Serra.Text = FORMATTA_ETTARI(Me.txtSuperficieMQ_Serra.Value)
End Sub
Private Sub fnModificaNomeLotto(IDLottoCampagna As Long, CodiceLottoCampagna As String, DescrizioneLottoCampagna As String)
Dim sSQL As String

sSQL = "UPDATE RV_PO01_SchedaPerLotto SET "
sSQL = sSQL & "CodiceLottoCampagna=" & fnNormString(CodiceLottoCampagna) & ", "
sSQL = sSQL & "DescrizioneLottoCampagna=" & fnNormString(DescrizioneLottoCampagna) & " "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Cn.Execute sSQL


End Sub
Private Function GET_TIPO_CODICE() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCodiceLotto FROM RV_PO01_ParametriFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_CODICE = 1
Else
    If fnNotNullN(rs!IDTipoCodiceLotto) = 0 Then
        GET_TIPO_CODICE = 1
    Else
        GET_TIPO_CODICE = fnNotNullN(rs!IDTipoCodiceLotto)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_LOTTO() As String
Dim Codice As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer

If LINK_TIPO_CONTATORE = 0 Then
    sSQL = "SELECT NumeroLottoDiCampagna "
    sSQL = sSQL & "FROM RV_PO01_ParametriFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Codice = 1
    Else
        Codice = fnNotNullN(rs!NumeroLottoDiCampagna)
    End If
    
    GET_NUMERO_LOTTO = ""
    For I = 6 To (Len(CStr(Codice)) + 1) Step -1
        GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & "0"
    Next I
    
    GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & CStr(Codice)
    rs.CloseResultset
    Set rs = Nothing
Else
    sSQL = "SELECT Numerazione "
    sSQL = sSQL & "FROM RV_PO01_ContatoreLotto "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_PO01_PeriodoCampagna=" & Me.cboPeriodoCampagna.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Codice = 1
    Else
        Codice = fnNotNullN(rs!Numerazione)
    End If
    
    GET_NUMERO_LOTTO = ""
    For I = 6 To (Len(CStr(Codice)) + 1) Step -1
        GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & "0"
    Next I
    
    GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & CStr(Codice)
    rs.CloseResultset
    Set rs = Nothing
End If

End Function

Private Function SALVA_NUMERO_LOTTO() As String
Dim rs As ADODB.Recordset
Dim I As Integer
Dim sSQL As String

If LINK_TIPO_CONTATORE = 0 Then
    sSQL = "SELECT NumeroLottoDiCampagna FROM RV_PO01_ParametriFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic
    
    If rs.EOF = False Then
        
        rs!NumeroLottoDiCampagna = fnNotNullN(rs!NumeroLottoDiCampagna) + 1
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
Else
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_PO01_ContatoreLotto "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_PO01_PeriodoCampagna=" & Me.cboPeriodoCampagna.CurrentID
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!IDRV_PO01_ContatoreLotto = fnGetNewKey("RV_PO01_ContatoreLotto", "IDRV_PO01_ContatoreLotto")
        rs!IDAzienda = TheApp.IDFirm
        rs!IDFiliale = TheApp.Branch
        rs!IDRV_PO01_PeriodoCampagna = Me.cboPeriodoCampagna.CurrentID
        rs!Numerazione = 2
    Else
        rs!Numerazione = fnNotNullN(rs!Numerazione) + 1
    End If
        
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
End If
End Function
Public Sub GET_CERTIFICAZIONE_SOCIO(IDAnagrafica As Long)

End Sub
Private Function GET_PROTOCOLLO_CERTIFICAZIONE(IDCertificazioneSocio As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ProtocolloCertificazione FROM RV_PO01_CertificazioneSocio "
sSQL = sSQL & "WHERE IDRV_PO01_Certificazione=" & Me.CDCertificazione.KeyFieldID
sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocio.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PROTOCOLLO_CERTIFICAZIONE = ""
Else
    GET_PROTOCOLLO_CERTIFICAZIONE = fnNotNull(rs!ProtocolloCertificazione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GENERA_FILTRO_PER_TIPO_OGGETTO()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim Filtro As String

sSQL = "DELETE FROM RV_PO01_Filtro "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_PO01_CreazioneLotto")

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

sSQL = "SELECT * FROM RV_PO01_Filtro"
Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic

rs.AddNew
    rs!IDUtente = TheApp.IDUser
    rs!IDAzienda = TheApp.IDFirm
    rs!IDTipoOggetto = fnGetTipoOggetto("RV_PO01_CreazioneLotto")
    rs!Filtro = Filtro
rs.Update

rs.Close
Set rs = Nothing

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

Private Sub GET_CONTROLLO_LICENZA()
Dim sSQL As String
Dim Codice_Diamante As String
Dim Codice_Prodotto_Calcolato  As String
Dim Codice_Attivazione As String

Dim Partita_Iva_Licenza As String

Codice_Diamante = GET_CODICE_DIAMANTE
Partita_Iva_Licenza = GET_PARTITA_IVA
Codice_Attivazione = GET_CODICE_SBLOCCO_ATTIVAZIONE


Codice_Prodotto_Calcolato = GET_CODICE_SBLOCCO(Codice_Diamante, Partita_Iva_Licenza, "01")

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
sSQL = sSQL & "WHERE IDRV_POProgramma=2"

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
sSQL = sSQL & "WHERE IDRV_POProgramma=2"

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

sSQL = "SELECT Count(IDRV_PO01_LottoCampagna) As NumeroInserimenti "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INSERIMENTI = 0
Else
    GET_NUMERO_INSERIMENTI = fnNotNullN(rs!NumeroInserimenti)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtSuperficieMQEffettiva_Change()
 If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtSuperficieMQEffettiva_LostFocus()
Dim I As Integer
Dim X As String 'Stringa da formattare
Dim J As String '
Dim Ris As String
Const MAX_X As Integer = 10
Dim Rimanenza As Integer
Dim Valore As Long
Valore = Me.txtSuperficieMQEffettiva.Value

Rimanenza = MAX_X - Len(CStr(Valore))

For I = 1 To Rimanenza
    X = X & "0"
Next
    'Stringa da formattare
    X = X & CStr(Valore)
    J = ""
For I = 1 To Len(X)
    
    J = J & Mid(X, I, 1)
    If I < Len(X) Then
        If I Mod 2 = 0 Then
            J = J & "."
        End If
    End If
Next
Me.txtSuperficieHAEffettiva.Text = J

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
Dim OLD_Cursor As Long
Dim CodiceLotto_OLD As String
Dim ProgressivoLotto As Long

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
            Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
            
            If (Me.txtCodiceLotto.Text = "") And (FLAG_AGGIUNGI_PROG_LOTTO = 0) Then
                Me.txtCodiceLotto.Text = GET_NUMERO_LOTTO
                
                m_Document("CodiceLotto").Value = Me.txtCodiceLotto.Text
                
                If Me.txtDescrizioneLotto.Text = "" Then
                    Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
                    m_Document("DescrizioneLotto").Value = Me.txtCodiceLotto.Text
                End If
                
                SALVA_NUMERO_LOTTO
                
            Else
                Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
                
                If FLAG_AGGIUNGI_PROG_LOTTO = 1 Then
                    CodiceLotto_OLD = Me.txtCodiceLotto.Text
                    If LINK_TIPO_POSIZIONE_PROG_LOTTO = 2 Then
                        Me.txtCodiceLotto.Text = GET_NUMERO_LOTTO & Me.txtCodiceLotto.Text
                    Else
                        Me.txtCodiceLotto.Text = Me.txtCodiceLotto.Text & GET_NUMERO_LOTTO
                    End If
                    
                    m_Document("CodiceLotto").Value = Me.txtCodiceLotto.Text
                    
                    If (Me.txtDescrizioneLotto.Text = CodiceLotto_OLD) Or (Me.txtDescrizioneLotto.Text = "") Then
                        Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
                        m_Document("DescrizioneLotto").Value = Me.txtCodiceLotto.Text
                    End If
                    
                    SALVA_NUMERO_LOTTO
                Else
                    m_Document("CodiceLotto").Value = Me.txtCodiceLotto.Text
                End If
            End If
            If ControlloEsistenzaCodiceLotto(Me.txtCodiceLotto.Text, m_Document(m_Document.PrimaryKey).Value) = True Then
                MsgBox "Il codice lotto è esistente", vbCritical, "Inserimento dati"
                Me.txtCodiceLotto.Text = ""
                m_Document("CodiceLotto").Value = ""
                GET_NUMERO_DOCUMENTO = -1
            Exit Function
            End If
        
            
            DoEvents
        
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
Exit Function

ERR_GET_NUMERO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "Errore coda"
    GET_NUMERO_DOCUMENTO = -1
    Unload frmCoda
End Function

Private Function GET_NUMERO_SBLOCCO() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT MAX(NumeroSbloccoLotto) as NumeroRecord "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_SBLOCCO = 0
Else
    GET_NUMERO_SBLOCCO = fnNotNullN(rs!NumeroRecord) + 1
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_MAGAZZINO_PARAMETRI() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDMagazzinoCaricoLottoVeg FROM RV_PO01_ParametriFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_MAGAZZINO_PARAMETRI = 0
Else
    GET_LINK_MAGAZZINO_PARAMETRI = fnNotNullN(rs!IDMagazzinoCaricoLottoVeg)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CAUSALE_PARAMETRI() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzioneCaricoLottoVeg FROM RV_PO01_ParametriFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CAUSALE_PARAMETRI = 0
Else
    GET_LINK_CAUSALE_PARAMETRI = fnNotNullN(rs!IDFunzioneCaricoLottoVeg)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_ESISTENZA_LOTTO_ARTICOLO(CodiceLotto As String, IDArticolo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM LottoArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND Codice=" & fnNormString(CodiceLotto)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_LOTTO_ARTICOLO = False
Else
    GET_ESISTENZA_LOTTO_ARTICOLO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_LOTTO_ARTICOLO(CodiceLotto As String, IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLottoArticolo FROM LottoArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND Codice=" & fnNormString(CodiceLotto)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LOTTO_ARTICOLO = 0
Else
    GET_LINK_LOTTO_ARTICOLO = fnNotNullN(rs!IDLottoArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ARTICOLO_GESTIONE_LOTTI(IDArticolo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT GestioneLotti FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & m_App.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_GESTIONE_LOTTI = False
Else
    GET_ARTICOLO_GESTIONE_LOTTI = fnNotNullN(rs!GestioneLotti)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function MOVIMENTAZIONE_ARTICOLO(IDArticolo As Long, Articolo As String, IDValoriOggettoDettaglio As Long, Quantita As Double, DataDocumento As String) As Boolean
On Error GoTo ERR_MOVIMENTAZIONE_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Set Mov = New DmtMovim.cMovimentazione
Set Mov.Connection = TheApp.Database.Connection

''''''''ELIMINAZIONE MOVIMENTO'''''''''''''''''''''''''''''''''''''''
'Mov.IDTipoOggetto = m_DocType.ID
'Mov.IDOggetto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
'Mov.Field "IDValoriOggettoDettaglio", IDValoriOggettoDettaglio

sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & m_DocType.ID
sSQL = sSQL & " AND IDOggetto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Cn.BeginTrans
    Mov.Delete fnNotNullN(rs!IDMovimento)
    Cn.CommitTrans
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rs.CloseResultset
Set rs = Nothing

If Quantita <= 0 Then
    Set Mov = Nothing
    MOVIMENTAZIONE_ARTICOLO = True
    Exit Function
End If

If Len(fnNotNull(DataDocumento)) = 0 Then DataDocumento = Date

Mov.DataMovimento = DataDocumento
Mov.FattoreDiConversione = Null
Mov.GestioneLotti = True
Mov.IDEsercizio = fnGetEsercizio(DataDocumento)
Mov.IDTipoOggetto = m_DocType.ID
Mov.IDOggetto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Mov.IDFunzione = Me.cboFunzioneCaricoLotto.CurrentID
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Me.cboMagazzinoCaricoLotto.CurrentID
Mov.IDMagazzinoUscita = Me.cboMagazzinoCaricoLotto.CurrentID
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", GET_LINK_UM(IDArticolo)
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Quantita
Mov.Field "Importo", 0
Mov.Field "DataDocumento", DataDocumento
Mov.Field "Oggetto", TheApp.FunctionName
Mov.Field "IDTipoMovimento", 1
Mov.Field "IDValoriOggettoDettaglio", IDValoriOggettoDettaglio
Mov.Field "IDLottoArticolo", GET_LINK_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, IDArticolo)
Mov.Field "TipoRiga", trcNessuno

Cn.BeginTrans
MOVIMENTAZIONE_ARTICOLO = Mov.Insert
Cn.CommitTrans
Set Mov = Nothing
Exit Function
ERR_MOVIMENTAZIONE_ARTICOLO:
    On Error GoTo ERR_ERROR_END
    MOVIMENTAZIONE_ARTICOLO = False
    Cn.RollbackTrans
    Exit Function
ERR_ERROR_END:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    MOVIMENTAZIONE_ARTICOLO = False

End Function
Private Function GET_LINK_UM(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraAcquisto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM = 0
Else
    GET_LINK_UM = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_MOVIMENTAZIONE_LOTTO(IDArticolo As Long, IDLottoArticolo As Long, IDOggetto As Long, IDTipoOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDLottoArticolo=" & IDLottoArticolo
sSQL = sSQL & " AND IDTipoOggetto<>" & IDTipoOggetto
sSQL = sSQL & " AND IDOggetto<>" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_MOVIMENTAZIONE_LOTTO = False
Else
    GET_ESISTENZA_MOVIMENTAZIONE_LOTTO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_ESISTENZA_MOV_RIGHE(IDLottoCampagna As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Lotto As Long

GET_CONTROLLO_ESISTENZA_MOV_RIGHE = False

sSQL = "SELECT RV_PO01_DettaglioLotto.IDArticolo, RV_PO01_DettaglioLotto.IDRV_PO01_DettaglioLotto, Articolo.IDAzienda, "
sSQL = sSQL & "RV_PO01_DettaglioLotto.IDRV_PO01_LottoCampagna , Articolo.GestioneLotti "
sSQL = sSQL & "FROM RV_PO01_DettaglioLotto INNER JOIN "
sSQL = sSQL & "Articolo ON RV_PO01_DettaglioLotto.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE RV_PO01_DettaglioLotto.IDRV_PO01_LottoCampagna=" & IDLottoCampagna
sSQL = sSQL & " AND Articolo.GestioneLotti=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Link_Lotto = GET_LINK_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, fnNotNullN(rs!IDArticolo))
    If Link_Lotto > 0 Then
        If GET_CONTROLLO_ESISTENZA_MOV_RIGHE = False Then
            If GET_ESISTENZA_MOVIMENTAZIONE_LOTTO(fnNotNullN(rs!IDArticolo), Link_Lotto, IDLottoCampagna, m_DocType.ID) = True Then
                GET_CONTROLLO_ESISTENZA_MOV_RIGHE = True
            End If
        End If
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Function ELIMINA_MOVIMENTI_DOCUMENTO(IDLottoCampagna As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Lotto As Long

ELIMINA_MOVIMENTI_DOCUMENTO = False

sSQL = "SELECT RV_PO01_DettaglioLotto.IDArticolo, RV_PO01_DettaglioLotto.IDRV_PO01_DettaglioLotto, Articolo.IDAzienda, "
sSQL = sSQL & "RV_PO01_DettaglioLotto.IDRV_PO01_LottoCampagna , Articolo.GestioneLotti "
sSQL = sSQL & "FROM RV_PO01_DettaglioLotto INNER JOIN "
sSQL = sSQL & "Articolo ON RV_PO01_DettaglioLotto.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE RV_PO01_DettaglioLotto.IDRV_PO01_LottoCampagna=" & IDLottoCampagna
sSQL = sSQL & " AND Articolo.GestioneLotti=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Link_Lotto = GET_LINK_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, fnNotNullN(rs!IDArticolo))
    
    If Link_Lotto > 0 Then
        Set Mov = New DmtMovim.cMovimentazione
        Set Mov.Connection = TheApp.Database.Connection
        
        ''''''''ELIMINAZIONE MOVIMENTO'''''''''''''''''''''''''''''''''''''''
        Mov.IDTipoOggetto = m_DocType.ID
        Mov.IDOggetto = IDLottoCampagna
        Mov.Field "IDValoriOggettoDettaglio", fnNotNullN(rs!IDRV_PO01_DettaglioLotto)
        'Cn.BeginTrans
        Mov.Delete
        'Cn.CommitTrans
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Set Mov = Nothing
        
        ELIMINA_LOTTO_ARTICOLO fnNotNullN(rs!IDArticolo), Me.txtCodiceLotto.Text
    End If
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_LOTTO_ARTICOLO(IDArticolo As Long, LottoArticolo As String)
Dim sSQL As String
Dim LINK_MAGAZZINO As Long
Dim Link_Lotto As Long

Link_Lotto = GET_LINK_LOTTO_ARTICOLO(Me.txtCodiceLotto.Text, IDArticolo)

If Link_Lotto > 0 Then
    '''ELIMAZIONE DEL LOTTO'''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM LottoArticolo "
    sSQL = sSQL & "WHERE IDLottoArticolo=" & Link_Lotto
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''ELIMINAZIONE LOTTO MAGAZZINO''''''''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM LottoArticoloPerMagazzino "
    sSQL = sSQL & "WHERE IDLottoArticolo=" & Link_Lotto
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

End Sub
Private Function GET_ESISTENZA_PROGRAMMA(IDProgramma As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POProgramma FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & IDProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_PROGRAMMA = False
Else
    GET_ESISTENZA_PROGRAMMA = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ESISTENZA_LOTTO_CONFERITO(IDLottoCampagna As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDLottoDiCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_LOTTO_CONFERITO = False
Else
    GET_ESISTENZA_LOTTO_CONFERITO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_CERTIFICAZIONE(IDAnagraficaSocio As Long, IDFamigliaProdotti As Long, IDFiliale As Long)
On Error GoTo ERR_GET_CERTIFICAZIONE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_CERTIFICAZIONE As Long
Dim PROTOCOLLO_CERTIFICAZIONE As String

LINK_CERTIFICAZIONE = 0
PROTOCOLLO_CERTIFICAZIONE = ""

'''''''''Controllo del predefinito per famiglia per socio'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PO01_CertificazioneSocioFamiglia "
sSQL = sSQL & "WHERE IDRV_PO01_FamigliaProdotti=" & IDFamigliaProdotti
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaSocio
sSQL = sSQL & " AND Predefinito=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    LINK_CERTIFICAZIONE = fnNotNullN(rs!IDRV_PO01_Certificazione)
    PROTOCOLLO_CERTIFICAZIONE = fnNotNull(rs!ProtocolloCertificazione)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If LINK_CERTIFICAZIONE = 0 Then
    '''''''''Controllo del predefinito per famiglia per socio'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM RV_PO01_CertificazioneSocio "
    sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagraficaSocio
    sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        LINK_CERTIFICAZIONE = fnNotNullN(rs!IDRV_PO01_Certificazione)
        PROTOCOLLO_CERTIFICAZIONE = fnNotNull(rs!ProtocolloCertificazione)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If
If LINK_CERTIFICAZIONE = 0 Then
    '''''''''Controllo del predefinito per famiglia per socio'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM RV_PO01_ParametriFiliale "
    sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        If fnNotNullN(rs!IDRV_PO01_Certificazione) > 0 Then
            If (Abs(fnNotNullN(rs!ProtocolloCertificazionePredefinito))) > 0 Then
                LINK_CERTIFICAZIONE = fnNotNullN(rs!IDRV_PO01_Certificazione)
                PROTOCOLLO_CERTIFICAZIONE = fnNotNull(rs!ProtocolloCertificazione)
            End If
        End If
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

Me.CDCertificazione.Load LINK_CERTIFICAZIONE
Me.txtProtocolloCertificazione.Text = PROTOCOLLO_CERTIFICAZIONE

Exit Sub
ERR_GET_CERTIFICAZIONE:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    Me.CDCertificazione.Load 0
    Me.txtProtocolloCertificazione.Text = ""
    
End Sub
Private Function GET_CODICE_LOTTO_CAMPAGNA() As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim CodiceLotto As String

sSQL = "SELECT RV_PO01_ConfigLottoCampagnaRighe.IDRV_PO01_ConfigLottoCampagnaRighe, "
sSQL = sSQL & "RV_PO01_ConfigLottoCampagnaRighe.IDRV_PO01_ConfigLottoCampagna, RV_PO01_ConfigLottoCampagnaRighe.IDRV_PO01_CampiLottoCampagna, "
sSQL = sSQL & "RV_PO01_ConfigLottoCampagnaRighe.Lunghezza, RV_PO01_ConfigLottoCampagnaRighe.Progressivo, "
sSQL = sSQL & "RV_PO01_ConfigLottoCampagnaRighe.Testo "
sSQL = sSQL & "FROM RV_PO01_ConfigLottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_ConfigLottoCampagnaRighe ON "
sSQL = sSQL & "RV_PO01_ConfigLottoCampagna.IDRV_PO01_ConfigLottoCampagna = RV_PO01_ConfigLottoCampagnaRighe.IDRV_PO01_ConfigLottoCampagna "
sSQL = sSQL & "WHERE RV_PO01_ConfigLottoCampagna.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " ORDER BY RV_PO01_ConfigLottoCampagnaRighe.Progressivo"

CodiceLotto = ""

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Select Case fnNotNullN(rs!IDRV_PO01_CampiLottoCampagna)
        Case 1 'TESTO
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(fnNotNull(rs!Testo), fnNotNullN(rs!Lunghezza), False, "")
        Case 2 'Codice socio
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.CDSocio.Code, fnNotNullN(rs!Lunghezza), True, "0")
        Case 3 'Anagrafica socio
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.CDSocio.Description, fnNotNullN(rs!Lunghezza), False, "")
        Case 4 'Codice famiglia prodotto
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtCodiceFamigliaProdotto.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 5 'Descrizione famiglia prodotto
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.cboFamigliaProdotti.Text, fnNotNullN(rs!Lunghezza), False, "")
        Case 6 'Codice varietà prodotto
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtCodiceVarieta.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 7 'Anno della data di semina presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Year(Me.txtDataSemina.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 8 'Anno della data di semina presunta ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("yy", Me.txtDataSemina.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 9 'Mese della data di semina presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Month(Me.txtDataSemina.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 10 'Giorno della data di semina presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Day(Me.txtDataSemina.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 11 'Giorno della Settimana della data di semina presunta
             CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Weekday(Me.txtDataSemina.Text, vbMonday), fnNotNullN(rs!Lunghezza), True, "0")
        Case 12 'Anno della data di inserimento
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Year(Date), fnNotNullN(rs!Lunghezza), True, "0")
        Case 13 'Anno della data di inserimento ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("yy", Date), fnNotNullN(rs!Lunghezza), True, "0")
        Case 14 'Mese della data di inserimento
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Month(Date), fnNotNullN(rs!Lunghezza), True, "0")
        Case 15 'Giorno della data di inserimento
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Day(Date), fnNotNullN(rs!Lunghezza), True, "0")
        Case 16 'Settimana della data di inserimento
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Weekday(Date, vbMonday), fnNotNullN(rs!Lunghezza), True, "0")
        Case 17 'Codice del tipo produzione
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtCodiceTipoProduzione.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 18 'Descrizione del tipo di produzione
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.cboTipoProduzione.Text, fnNotNullN(rs!Lunghezza), False, "")
        Case 19 'Anno di riferimento del periodo di campagna
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtAnnoRifPeriodoCampagna.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 20 'Codice del periodo di campagna
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtCodicePeriodoCampagna.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 21 'Anno della data di inizio produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Year(Me.txtDataInizioProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 22 'Anno della data di inizio produzione presunta ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("yy", Me.txtDataInizioProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 23 'Mese della data di inizio produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Month(Me.txtDataInizioProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 24 'Giorno della data di inizio produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Day(Me.txtDataInizioProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 25 'Settimana della data di inizio produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Weekday(Me.txtDataInizioProduzione.Text, vbMonday), fnNotNullN(rs!Lunghezza), True, "0")
        Case 26 'Anno della data di fine produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Year(Me.txtDataFineProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 27 'Anno della data di fine produzione presunta ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("yy", Me.txtDataFineProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 28 'Mese della data di fine produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Month(Me.txtDataFineProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 29 'Giorno della data di fine produzione presunta ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Day(Me.txtDataFineProduzione.Text), fnNotNullN(rs!Lunghezza), True, "0")
        Case 30 'Settimana della data di fine produzione presunta ridotta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Weekday(Me.txtDataFineProduzione.Text, vbMonday), fnNotNullN(rs!Lunghezza), True, "0")
        Case 31 'Descrizione della varietà del prodotto
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtVarieta.Text, fnNotNullN(rs!Lunghezza), True, "0")
        Case 32 'Settimana dell'anno della data di semina presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("ww", Me.txtDataSemina.Text, , vbFirstFullWeek), fnNotNullN(rs!Lunghezza), True, "0")
        Case 33 'Settimana dell'anno della data di inserimento
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("ww", Date, , vbFirstFullWeek), fnNotNullN(rs!Lunghezza), True, "0")
        Case 34 'Settimana dell'anno della data di inizio produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("ww", Me.txtDataInizioProduzione.Text, , vbFirstFullWeek), fnNotNullN(rs!Lunghezza), True, "0")
        Case 35 'Settimana dell'anno della data di fine produzione presunta
            CodiceLotto = CodiceLotto & FORMATTA_CAMPO(DatePart("ww", Me.txtDataFineProduzione.Text, , vbFirstFullWeek), fnNotNullN(rs!Lunghezza), True, "0")
        Case 36 'Descrizione del lotto di campagna
             CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.txtDescrizioneLotto.Text, fnNotNullN(rs!Lunghezza), True, "")
        Case 37 'Codice Import/Export del lotto di campagna
             CodiceLotto = CodiceLotto & FORMATTA_CAMPO(Me.TxtCodiceImEx.Text, fnNotNullN(rs!Lunghezza), True, "")
    End Select
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

GET_CODICE_LOTTO_CAMPAGNA = CodiceLotto

End Function
Private Function FORMATTA_CAMPO(Valore As String, Lunghezza As Long, filler As Boolean, StringaFiller As String) As String
Dim LunghezzaCampo As Long
Dim I As Long

LunghezzaCampo = Len(Valore)
FORMATTA_CAMPO = ""
If filler = True Then
    Select Case (LunghezzaCampo - Lunghezza)
        Case Is = 0
            FORMATTA_CAMPO = Valore
        Case Is > 0
            FORMATTA_CAMPO = Mid(Valore, 1, Lunghezza)
        Case Is < 0
            FORMATTA_CAMPO = Valore
            For I = LunghezzaCampo To Lunghezza - 1
                FORMATTA_CAMPO = StringaFiller & FORMATTA_CAMPO
            Next
    
    End Select

Else
    If LunghezzaCampo >= Lunghezza Then
        FORMATTA_CAMPO = Mid(Valore, 1, Lunghezza)
    Else
        FORMATTA_CAMPO = Valore
    End If
End If



End Function

Private Function GET_CODICE_TABELLA(Tabella As String, Campo As String, CampoWhere As String, ValoreCampoWhere As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & Campo & " FROM " & Tabella
sSQL = sSQL & " WHERE " & CampoWhere & "=" & fnNotNullN(ValoreCampoWhere)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_TABELLA = ""
Else
    GET_CODICE_TABELLA = fnNotNull(rs.adoColumns(Campo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtVarieta_Change()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    'Me.txtCodiceLotto.Text = GET_CODICE_LOTTO_CAMPAGNA
    'Me.txtDescrizioneLotto.Text = Me.txtCodiceLotto.Text
End If
End Sub
Private Sub GET_PARAMETRI_PROGRESSIVO_LOTTO(IDFiliale As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_ConfigLottoCampagna "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    FLAG_AGGIUNGI_PROG_LOTTO = 0
    LINK_TIPO_POSIZIONE_PROG_LOTTO = 2
Else
    FLAG_AGGIUNGI_PROG_LOTTO = fnNotNullN(rs!UtilizzaProgressivoLotto)
    LINK_TIPO_POSIZIONE_PROG_LOTTO = fnNotNullN(rs!IDRV_PO01_LatoProgressivoLotto)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_SCHEMA_PREDEFINITO_SOCIO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_Schema "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDSocio=" & IDAnagrafica
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_SCHEMA_PREDEFINITO_SOCIO = 0
Else
    GET_SCHEMA_PREDEFINITO_SOCIO = fnNotNullN(rs!IDRV_PO01_Schema)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CALCOLO_DATE_DOPO_SEMINA(IDFamiglia As Long, DataSemina As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumSettimaSemina As Long


NumSettimaSemina = DatePart("ww", DataSemina, , vbFirstFullWeek)

sSQL = "SELECT * FROM RV_PO01_FamigliaProdottiRighe "
sSQL = sSQL & " WHERE IDRV_PO01_FamigliaProdotti=" & IDFamiglia
sSQL = sSQL & " AND NumeroSettimana=" & NumSettimaSemina

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
    Me.txtDataGerm.Value = 0
    Me.txtDataNascita.Value = 0
    Me.txtDataRipic.Value = 0
    Me.txtDataTrapianto.Value = 0
    Me.txtDataInizioProduzione.Value = 0
    Me.txtDataFineProduzione.Value = 0
    
Else
    If fnNotNullN(rs!NumeroGiorniGerminazioneSemina) > 0 Then
        Me.txtDataGerm.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniGerminazioneSemina), DataSemina)
    End If
    If fnNotNullN(rs!NumeroGiorniNascitaSemina) > 0 Then
        Me.txtDataNascita.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniNascitaSemina), DataSemina)
    End If
    If fnNotNullN(rs!NumeroGiorniRipicchettaggioSemina) > 0 Then
        Me.txtDataRipic.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniRipicchettaggioSemina), DataSemina)
    End If
    If fnNotNullN(rs!NumeroGiorniTrapiantoSemina) > 0 Then
        Me.txtDataTrapianto.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniTrapiantoSemina), DataSemina)
    End If
    If fnNotNullN(rs!NumeroGiorniInizioProduzione) > 0 Then
        Me.txtDataInizioProduzione.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniInizioProduzione), DataSemina)
    End If
    If fnNotNullN(rs!NumeroGiorniFineProduzione) > 0 Then
        Me.txtDataFineProduzione.Value = DateAdd("d", fnNotNullN(rs!NumeroGiorniFineProduzione), DataSemina)
    End If

End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_GRIGLIA_ORDINE(IDLottoCampagna As Long)
On Error GoTo ERR_GetGrigliaLavorazione
Dim sSQL As String
Dim cl As dgColumnHeader
Dim OLDCursor As Long

sSQL = "SELECT ValoriOggettoDettaglio0010.IDValoriOggettoDettaglio, ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, "
sSQL = sSQL & "ValoriOggettoDettaglio0010.Link_Art_articolo, ValoriOggettoDettaglio0010.Art_codice, ValoriOggettoDettaglio0010.Art_descrizione, "
sSQL = sSQL & "ValoriOggettoDettaglio0010.Art_quantita_totale, ValoriOggettoPerTipo000F.Doc_data, ValoriOggettoPerTipo000F.Doc_numero, "
sSQL = sSQL & "ValoriOggettoPerTipo000F.Link_Nom_anagrafica, ValoriOggettoPerTipo000F.Nom_nome, ValoriOggettoPerTipo000F.Nom_codice, "
sSQL = sSQL & "ValoriOggettoPerTipo000F.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo000F.Doc_ordine_chiuso, "
sSQL = sSQL & "ValoriOggettoDettaglio0010.RV_PO01_UbiLottoDiCampagna, ValoriOggettoDettaglio0010.RV_PO01_LottoDiCampagna, "
sSQL = sSQL & "ValoriOggettoDettaglio0010.RV_POLinkRiga "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000F ON ValoriOggettoDettaglio0010.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto "
'sSQL = sSQL & "SELECT * FROM RV_PO01_IETrovaOrdinePerLotto "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_PO01_IDLottoCampagna=" & IDLottoCampagna
sSQL = sSQL & " AND RV_POTipoRiga=1"
    
OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

If Not (rsGrigliaOrd Is Nothing) Then
    If rsGrigliaOrd.State > 0 Then
        rsGrigliaOrd.Close
    End If
    
    Set rsGrigliaOrd = Nothing
End If

Set rsGrigliaOrd = New ADODB.Recordset
rsGrigliaOrd.Open sSQL, Cn.InternalConnection

    

    With Me.GrigliaOrdini
        .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Link_nom_anagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Doc_data", "Data documento", dgDate, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Doc_numero", "Numero documento", dgchar, True, 1800, dgAlignRight
        .ColumnsHeader.Add "Nom_codice", "Codice", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Nom_nome", "Nome", dgchar, True, 1800, dgAlignleft
        
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Art_descrizione", "Articolo", dgchar, True, 1800, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Quantità", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set .Recordset = rsGrigliaOrd
        .LoadUserSettings
        .Refresh
    End With

Cn.CursorLocation = OLDCursor


Exit Function
ERR_GetGrigliaLavorazione:
    MsgBox Err.Description, vbCritical, "Griglia ricerca"
End Function
Private Sub AGGIORNA_ORDINI_DA_SERRE(CodiceLotto As String, IDLottoCampagna As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim SERRE As String

sSQL = "SELECT ValoriOggettoDettaglio0010.IDValoriOggettoDettaglio "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000F ON ValoriOggettoDettaglio0010.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoDettaglio0010.RV_PO01_IDLottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    SERRE = GET_STRINGA_SERRE(IDLottoCampagna)
    
    sSQL = "UPDATE ValoriOggettoDettaglio0010 SET "
    sSQL = sSQL & "RV_PO01_UbiLottoDiCampagna=" & fnNormString(SERRE) & ", "
    sSQL = sSQL & "RV_PO01_IDLottoCampagna=" & IDLottoCampagna & ", "
    sSQL = sSQL & "RV_PO01_LottoDiCampagna=" & fnNormString(CodiceLotto)
    sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rs!IDValoriOggettoDettaglio)
    Cn.Execute sSQL
    
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_STRINGA_SERRE(IDLottoCampagna) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_STRINGA_SERRE = ""

sSQL = "SELECT RV_PO01_SerraPerLotto.IDRV_PO01_Serra, RV_PO01_Serra.Codice, RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna "
sSQL = sSQL & "FROM RV_PO01_SerraPerLotto INNER JOIN "
sSQL = sSQL & "RV_PO01_Serra ON RV_PO01_SerraPerLotto.IDRV_PO01_Serra = RV_PO01_Serra.IDRV_PO01_Serra "
sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_STRINGA_SERRE = GET_STRINGA_SERRE & fnNotNull(rs!Codice) & "|"
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

GET_STRINGA_SERRE = Mid(GET_STRINGA_SERRE, 1, 250)

End Function
Private Sub ELIMINA_RIFERIMENTI_ORDINE(CodiceLotto As String)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0010 SET "
sSQL = sSQL & "RV_PO01_UbicazioneLottoDiCampagna=" & fnNormString("") & ", "
sSQL = sSQL & "RV_PO01_LottoDiCampagna=" & fnNormString("")
sSQL = sSQL & "WHERE RV_PO01_LottoDiCampagna=" & fnNormString(CodiceLotto)
Cn.Execute sSQL
End Sub
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
Private Sub GET_RIEPILOGO_ORDINE(IDLottoCampagna As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(art_quantita_totale) AS Quantita, "
sSQL = sSQL & "SUM(Art_qta_evasa_totale) AS QuantitaEvasa, "
sSQL = sSQL & "SUM(Art_qta_evad_totale) AS QuantitaDaEvadere "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE RV_PO01_IDLottoCampagna=" & IDLottoCampagna
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtQtaOrdinata.Value = 0
    Me.txtQtaEvasa.Value = 0
    Me.txtQtaDaEvadere.Value = 0
Else
    Me.txtQtaOrdinata.Value = fnNotNullN(rs!Quantita)
    Me.txtQtaEvasa.Value = fnNotNullN(rs!QuantitaEvasa)
    Me.txtQtaDaEvadere.Value = fnNotNullN(rs!QuantitaDaEvadere)
End If


rs.CloseResultset
Set rs = Nothing
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
Private Sub RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO(IDLottoCampagna As Long, IDLottoCampagnaNew As Long)
On Error GoTo ERR_RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_PO01_DettaglioLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PO01_DettaglioLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagnaNew

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PO01_DettaglioLotto = fnGetNewKey("RV_PO01_DettaglioLotto", "IDRV_PO01_DettaglioLotto")
        rsNew!IDRV_PO01_LottoCampagna = IDLottoCampagnaNew
        rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsNew!Calibro = fnNotNull(rs!Calibro)
        rsNew!QuantitaPresunta = fnNotNull(rs!QuantitaPresunta)
        rsNew("IDUtenteInserimento").Value = TheApp.IDUser
        rsNew("PCInserimento").Value = GET_NOMECOMPUTER
        rsNew("UtentePCInserimento").Value = GET_NOMEUTENTE
        rsNew("DataInserimento").Value = Date
        rsNew("OraInserimento").Value = GET_ORARIO(Now)
        rsNew("IDUtenteUltimaModifica").Value = TheApp.IDUser
        rsNew("PCUltimaModifica").Value = GET_NOMECOMPUTER
        rsNew("UtentePCUltimaModifica").Value = GET_NOMEUTENTE
        rsNew("DataUltimaModifica").Value = Date
        rsNew("OraUltimaModifica").Value = GET_ORARIO(Now)
        
        
    rsNew.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO:
    MsgBox Err.Description, vbCritical, "RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO"
End Sub
Private Sub RIPORTA_SERRE_DA_SALVA_COME_NUOVO(IDLottoCampagna As Long, IDLottoCampagnaNew As Long)
On Error GoTo ERR_RIPORTA_SERRE_DA_SALVA_COME_NUOVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagnaNew

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PO01_SerraPerLotto = fnGetNewKey("RV_PO01_SerraPerLotto", "IDRV_PO01_SerraPerLotto")
        rsNew!IDRV_PO01_LottoCampagna = IDLottoCampagnaNew
        rsNew!IDRV_PO01_Serra = fnNotNullN(rs!IDRV_PO01_Serra)
        rsNew!DimensioneMq = fnNotNullN(rs!DimensioneMq)
        rsNew!DimensioneHA = fnNotNull(rs!DimensioneHA)
        
    rsNew.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_RIPORTA_SERRE_DA_SALVA_COME_NUOVO:
    MsgBox Err.Description, vbCritical, "RIPORTA_SERRE_DA_SALVA_COME_NUOVO"
End Sub

Private Sub GET_UTENTE_SBLOCCO(IDUtente As Long, IDFiliale As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer
Dim UtenteSblocco As Boolean

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POUtentiSbloccoLottoCampagna "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    UtenteSblocco = False
Else
    UtenteSblocco = True
End If

rs.CloseResultset
Set rs = Nothing

txtDataSbloccoLotto.Enabled = UtenteSblocco
txtProgressivo.Enabled = UtenteSblocco
cmdProgressivo.Enabled = UtenteSblocco
chkChiuso.Enabled = UtenteSblocco

End Sub
Private Function GET_CONTROLLO_AZIONI_UTENTE(IDTipoOggetto As Long, IDUtente As Long, IDFiliale As Long, IDAzione As Long) As Boolean
Dim sem As Semaforo.dmtSemaphore

Set sem = New Semaforo.dmtSemaphore
Set sem.Database = TheApp.Database.Connection

GET_CONTROLLO_AZIONI_UTENTE = sem.IsAuthorized(IDTipoOggetto, IDUtente, IDFiliale, IDAzione)

Set sem = Nothing

End Function
Private Sub GET_DATI_PER_RINNOVO()
On Error GoTo ERR_GET_DATI_PER_RINNOVO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Integer
Dim DataInizioPeriodo As String
Dim DataFinePeriodo As String

frmSalvaComeNuovo.Show vbModal

If SALVA_COME_NUOVO = True Then
    
    fraAttesa.Top = (Me.Height / 2) - (Me.fraAttesa.Height / 2)
    fraAttesa.Left = (Me.Width / 2) - (Me.fraAttesa.Width / 2)
    
    
    fraAttesa.Visible = True
    
    Me.Enabled = False
    Screen.MousePointer = 11
    
    DoEvents
    
    GET_PARAMETRI_PERIODO_CAMPAGNA RIPORTA_LINK_PERIODO_CAMPAGNA, DataInizioPeriodo, DataFinePeriodo
    
    
    
    CREA_RECORDSET
    
    sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagna "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
    sSQL = sSQL & GET_SQL_PER_RINNOVO
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, Cn.InternalConnection
    
    rsRinnovo.Open , , adOpenKeyset, adLockBatchOptimistic
    
    NUMERO_RINNOVO = 0
    While Not rs.EOF
        rsRinnovo.AddNew
            For I = 0 To rs.Fields.Count - 1
                Select Case rs.Fields(I).Name
                    Case "DataInizioProduzione"
                        If Len(fnNotNull(rs!DataInizioProduzione)) > 0 Then
                            rsRinnovo.Fields(rs.Fields(I).Name).Value = GET_DATA_PRODUZIONE(rs!DataInizioProduzione, DataInizioPeriodo, fnNotNullN(rs!IDRV_PO01_Varieta))
                            rsRinnovo.Fields(rs.Fields(I).Name).Value = GET_DATA_PRODUZIONE_VARIETA(fnNotNull(rsRinnovo.Fields(rs.Fields(I).Name).Value), fnNotNullN(rs!IDRV_PO01_Varieta), 1)
                        End If
                    Case "DataFineProduzione"
                        If Len(fnNotNull(rs!DataFineProduzione)) > 0 Then
                            rsRinnovo.Fields(rs.Fields(I).Name).Value = GET_DATA_PRODUZIONE(rs!DataFineProduzione, DataFinePeriodo, fnNotNullN(rs!IDRV_PO01_Varieta))
                            rsRinnovo.Fields(rs.Fields(I).Name).Value = GET_DATA_PRODUZIONE_VARIETA(fnNotNull(rsRinnovo.Fields(rs.Fields(I).Name).Value), fnNotNullN(rs!IDRV_PO01_Varieta), 2)
                        End If
                    Case "IDRV_PO01_PeriodoCampagna"
                        rsRinnovo.Fields(rs.Fields(I).Name).Value = RIPORTA_LINK_PERIODO_CAMPAGNA
                    Case "PeriodoCampagna"
                        
                    Case "IDRV_PO01_StatoLotto"
                        rsRinnovo.Fields(rs.Fields(I).Name).Value = RIPORTA_LINK_STATO_LOTTO
                    Case "StatoLotto"
                        
                    Case Else
                        rsRinnovo.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                End Select
            Next
            rsRinnovo!Rinnova = 1
        rsRinnovo.Update
        NUMERO_RINNOVO = NUMERO_RINNOVO + 1
    rs.MoveNext
    Wend
    Screen.MousePointer = 0
    fraAttesa.Visible = False
    Me.Enabled = True
    
    frmRinnovo.Show vbModal
    
    ExecuteSearch
    
End If
Exit Sub
ERR_GET_DATI_PER_RINNOVO:
    MsgBox Err.Description, vbCritical, "GET_DATI_PER_RINNOVO"
    
    On Error Resume Next
    fraAttesa.Visible = False
    Me.Enabled = False
    Screen.MousePointer = 0
    
End Sub
Private Function GET_SQL_PER_RINNOVO() As String
On Error GoTo ERR_GET_SQL_PER_RINNOVO
Dim Field As DmtDocManLib.Field
Dim Cond As DmtGridCtl.dgCondition
Dim sWhere As String

    GET_SQL_PER_RINNOVO = ""
    
    'Comunica all'oggetto DocType i valori da usare per la ricerca
    For Each Cond In BrwMain.Conditions
        If Cond.IsHeader = False Then
            Select Case Cond.ConditionType
                
                'Condizione boolean
                Case dgCondTypeBoolean
                    If (Cond.FromValue = "SI") Then
                        GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " = " & fnNormBoolean(1)
                    End If
                    
                    If (Cond.FromValue = "NO") Then
                        GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " = " & fnNormBoolean(0)
                    End If

                    
                'Condizione associata ad una combo box
                Case dgCondTypeComboDB
                    If fnNotNullN(BrwMain.Conditions(Cond.FieldName).FromValueID) > 0 Then
                      GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " = " & BrwMain.Conditions(Cond.FieldName).FromValueID
                    End If
                'Condizione di tipo text, numeric, data, time
                Case dgCondTypeText
                    If Cond.RangeChecked = True Then
                        If Not IsEmpty(Cond.FromValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " >= " & fnNormString(Cond.FromValue)
                        End If
                        If Not IsEmpty(Cond.ToValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " <= " & fnNormString(Cond.ToValue)
                        End If
                    Else
                        If Not IsEmpty(Cond.FromValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " LIKE " & fnNormString(Cond.FromValue & "%")
                        End If
                    End If
                Case dgCondTypeNumber
                        If Not IsEmpty(Cond.FromValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " >= " & fnNormNumber(Cond.FromValue)
                        End If
                        If Not IsEmpty(Cond.ToValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " <= " & fnNormNumber(Cond.ToValue)
                        End If
                
                Case dgCondTypeDate
                    If Cond.RangeChecked = True Then
                        If Not IsEmpty(Cond.FromValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " >= " & fnNormDate(Cond.FromValue)
                        End If
                        If Not IsEmpty(Cond.ToValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " <= " & fnNormDate(Cond.ToValue)
                        End If

                    Else
                        If Not IsEmpty(Cond.FromValue) Then
                            GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " = " & fnNormDate(Cond.FromValue)
                        End If
                    End If
                

              
                'Altre condizioni
                Case Else
                    If Not IsEmpty(Cond.FromValue) Then
                        GET_SQL_PER_RINNOVO = GET_SQL_PER_RINNOVO & " AND " & Cond.FieldName & " = " & Cond.FromValue
                    End If

            End Select
        End If
    Next Cond

        

Exit Function
ERR_GET_SQL_PER_RINNOVO:
    MsgBox Err.Description, vbCritical, "GET_SQL_PER_RINNOVO"
End Function


Private Sub CREA_RECORDSET()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Integer


If Not (rsRinnovo Is Nothing) Then
    If rsRinnovo.State > 0 Then
        rsRinnovo.Close
    End If
    Set rsRinnovo = Nothing
End If

Set rsRinnovo = New ADODB.Recordset
rsRinnovo.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=0"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

For I = 0 To rs.Fields.Count - 1
    
    Select Case rs.Fields(I).Type
    
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsRinnovo.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            
        Case adInteger
            rsRinnovo.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
       
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsRinnovo.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes

        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsRinnovo.Fields.Append rs.Fields(I).Name, adSmallInt, , rs.Fields(I).Attributes
        
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsRinnovo.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    
    End Select
    
Next

rsRinnovo.Fields.Append "Rinnova", adSmallInt, , adFldIsNullable

rs.Close
Set rs = Nothing
End Sub
Private Function GET_PARAMETRI_PERIODO_CAMPAGNA(IDPeriodoCampagnaNew As Long, DataInizio As String, DataFine As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_PeriodoCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_PeriodoCampagna=" & IDPeriodoCampagnaNew

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    DataInizio = fnNotNull(rs!DataInizio)
    DataFine = fnNotNull(rs!DataFine)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_DATA_PRODUZIONE(DataProduzione As String, DataPeriodo As String, IDVarieta As Long)
Dim Giorno As String
Dim Mese As String
Dim Anno As String



Giorno = Day(DataProduzione)
Mese = Month(DataProduzione)
Anno = Year(DataPeriodo)

If Len(Giorno) = 1 Then Giorno = "0" & Giorno
If Len(Mese) = 1 Then Mese = "0" & Mese

GET_DATA_PRODUZIONE = Giorno & "/" & Mese & "/" & Anno


End Function
Private Function GET_DATA_PRODUZIONE_VARIETA(DataProduzione As String, IDVarieta As Long, Tipo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaVarieta As Long
Dim Giorno As Long
Dim Mese As Long
Dim Anno As String

Dim GiornoProd As String
Dim MeseProd As String

GET_DATA_PRODUZIONE_VARIETA = DataProduzione

sSQL = "SELECT * FROM RV_PO01_Varieta "
sSQL = sSQL & "WHERE IDRV_PO01_Varieta=" & IDVarieta

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If Tipo = 1 Then
        Giorno = fnNotNullN(rs!GiornoInizioMat)
        Mese = fnNotNullN(rs!MeseInizioMat)
    End If
    
    If Tipo = 2 Then
        Giorno = fnNotNullN(rs!GiornoFineMat)
        Mese = fnNotNullN(rs!MeseFineMat)
    End If
    
End If

rs.CloseResultset
Set rs = Nothing


If ((Giorno > 0) And (Mese > 0)) Then
    MeseProd = CStr(Mese)
    GiornoProd = CStr(Giorno)
    
    If Len(MeseProd) = 1 Then MeseProd = "0" & MeseProd
    If Len(GiornoProd) = 1 Then GiornoProd = "0" & GiornoProd
    
    
    GET_DATA_PRODUZIONE_VARIETA = GiornoProd & "/" & MeseProd & "/" & Year(DataProduzione)
End If

End Function
Private Function FORMATTA_ETTARI(ValoreMQ As Double) As String
Dim ValoreEttaro As Double
Dim ArrayEttaro() As String
Dim ParteIntera As String
Dim ParteDecimale As String
Dim ParteDecimaleFormattata As String
Dim Numero As Integer
Dim NumeroZeri As String

Dim I As Long

ValoreEttaro = ValoreMQ / 10000
ParteDecimale = ""

ArrayEttaro = Split(ValoreEttaro, ",")

Numero = UBound(ArrayEttaro)

If Len(fnNotNull(ArrayEttaro(0))) > 0 Then
    ParteIntera = fnRoundDown(ArrayEttaro(0))
    ParteIntera = FormatNumber(ParteIntera, 0)
    
End If
If Numero = 1 Then
    ParteDecimale = ArrayEttaro(1)
End If

NumeroZeri = ""
For I = 4 To (Len(CStr(ParteDecimale)) + 1) Step -1
    NumeroZeri = NumeroZeri & "0"
Next

ParteDecimale = ParteDecimale & NumeroZeri
ParteDecimaleFormattata = ""

For I = 1 To Len(ParteDecimale)
    If I = 3 Then
        ParteDecimaleFormattata = ParteDecimaleFormattata & "."
        'ParteDecimaleFormattata = ParteDecimaleFormattata & Mid(ParteDecimale, I, 1)
    End If

    ParteDecimaleFormattata = ParteDecimaleFormattata & Mid(ParteDecimale, I, 1)
Next

FORMATTA_ETTARI = CStr(ParteIntera) & "." & ParteDecimaleFormattata

End Function
Private Sub cmdGestioneQualitaLista()
On Error GoTo ERR_cmdGestioneQualita_Click

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "2"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "1"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.CDSocio.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "cmdGestioneQualitaLista_Click"
End Sub
Private Sub cmdGestioneQualita()
On Error GoTo ERR_cmdGestioneQualita_Click

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "1"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "1"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.CDSocio.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "cmdGestioneQualita_Click"
End Sub
Private Sub CREA_SFALCI_PER_LOTTO(IDLottoCampagna As Long, NumeroSettimana As Long, IDFamiglia As Long)
On Error GoTo ERR_CREA_SFALCI_PER_LOTTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaProcesso As Boolean
Dim NumeroGiorniPrimoSfalcio As Long
Dim NumeroGiornoDalPrimoSfalcio As Long
Dim rsNew As ADODB.Recordset
Dim NumeroEla As Long
Dim DataPresunta As String
Dim ID As Long

NumeroGiorniPrimoSfalcio = 0
NumeroGiornoDalPrimoSfalcio = 0

AvviaProcesso = False

''''CONTROLLO PER INSERIMENTO AUTOMATICO'''''''''''''''''''''''''''''''''''''

sSQL = "SELECT AttivaGestioneSfalci FROM RV_PO01_FamigliaProdotti "
sSQL = sSQL & "WHERE IDRV_PO01_FamigliaProdotti=" & IDFamiglia

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!AttivaGestioneSfalci) = 1 Then
        AvviaProcesso = True
    End If
Else
    AvviaProcesso = False
End If

If AvviaProcesso = False Then Exit Sub

AvviaProcesso = False
sSQL = "SELECT * FROM RV_PO01_LottoCampagnaSfalcio "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    AvviaProcesso = True
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


If AvviaProcesso = False Then Exit Sub

sSQL = "SELECT * FROM RV_PO01_FamigliaProdottiRighe "
sSQL = sSQL & " WHERE IDRV_PO01_FamigliaProdotti=" & IDFamiglia
sSQL = sSQL & " AND NumeroSettimana=" & NumeroSettimana

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    NumeroGiorniPrimoSfalcio = fnNotNullN(rs!NumeroGiorniPrimoSfalcio)
    NumeroGiornoDalPrimoSfalcio = fnNotNullN(rs!NumeroGiorniProssimiSfalci)
End If

rs.CloseResultset
Set rs = Nothing

sSQL = "SELECT * FROM RV_PO01_Sfalcio "
sSQL = sSQL & " WHERE Dismesso=0 "
sSQL = sSQL & " ORDER BY Sequenza "

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PO01_LottoCampagnaSfalcio "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

NumeroEla = 1
DataPresunta = Me.txtDataSemina.Text
ID = fnGetNewKey("RV_PO01_LottoCampagnaSfalcio", "IDRV_PO01_LottoCampagnaSfalcio")
While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PO01_LottoCampagnaSfalcio = ID
        rsNew!Guid = GetGUID
        rsNew!IDRV_PO01_LottoCampagna = IDLottoCampagna
        rsNew!IDRV_PO01_Sfalcio = fnNotNullN(rs!IDRV_PO01_Sfalcio)
        If NumeroEla = 1 Then
            rsNew!DataPresuntaInizio = DateAdd("d", NumeroGiorniPrimoSfalcio, DataPresunta)
        Else
            rsNew!DataPresuntaInizio = DateAdd("d", NumeroGiornoDalPrimoSfalcio, DataPresunta)
        End If
        DataPresunta = rsNew!DataPresuntaInizio
    rsNew.Update
    NumeroEla = NumeroEla + 1
    ID = ID + 1
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing

m_DocumentsLink4.Refresh

Exit Sub
ERR_CREA_SFALCI_PER_LOTTO:
    MsgBox Err.Description, vbCritical, "CREA_SFALCI_PER_LOTTO"
End Sub
Private Function GET_CONTROLLO_ESISTENZA_SFALCIO(IDLottoCampagna As Long, IDSfalcio As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_ESISTENZA_SFALCIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_ESISTENZA_SFALCIO = False

sSQL = "SELECT * FROM RV_PO01_LottoCampagnaSfalcio "
sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna
sSQL = sSQL & " AND IDRV_PO01_Sfalcio=" & IDSfalcio

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_ESISTENZA_SFALCIO = True
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_ESISTENZA_SFALCIO:
    GET_CONTROLLO_ESISTENZA_SFALCIO = True
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_ESISTENZA_SFALCIO"
    
End Function
Private Function GET_CONTROLLO_ELIMINAZIONE_SFALCIO(ID As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_ESISTENZA_SFALCIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_ELIMINAZIONE_SFALCIO = True

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POCaricoMerceRighe "
sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagnaSfalcio=" & ID

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_ELIMINAZIONE_SFALCIO = False
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_ESISTENZA_SFALCIO:
    GET_CONTROLLO_ELIMINAZIONE_SFALCIO = True
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_ELIMINAZIONE_SFALCIO"
    
End Function
