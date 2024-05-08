VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAssegnazioneMerce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assegnazione merce"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAssegnazioneMerce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8955
      ScaleWidth      =   15675
      TabIndex        =   17
      Top             =   0
      Width           =   15735
      Begin VB.Frame fraRighe 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   15495
         Begin TabDlg.SSTab SSTab1 
            Height          =   8535
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   15285
            _ExtentX        =   26961
            _ExtentY        =   15055
            _Version        =   393216
            Tabs            =   1
            TabsPerRow      =   1
            TabHeight       =   520
            BackColor       =   0
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "frmAssegnazioneMerce.frx":4781A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label2(11)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label2(8)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label2(10)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label2(9)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblImballo"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label2(5)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label2(4)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label2(3)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label2(2)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label2(1)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label2(0)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "lblArticolo"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label2(6)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label7(0)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label8(0)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Label8(1)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label2(24)"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "Label2(29)"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Label2(28)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "Label2(14)"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "Label2(15)"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "txtPesoPerCollo"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txtQuantitaPerCollo"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "txtCostoConfezione"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "chkConfermaDaUtentePrim"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "chkTracciaImballoGestPrim"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "chkTracciaImballoGest"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "chkConfermaDaUtente"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "txtNumeroConfImballo"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "txtTaraConfImballo"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "CDImballoPrimario"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "fraScarti"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "chkScarti"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "ProgressBar1"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "cboUMImballo"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "txtColli"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).Control(36)=   "txtPesoLordo"
            Tab(0).Control(36).Enabled=   0   'False
            Tab(0).Control(37)=   "txtPesoNetto"
            Tab(0).Control(37).Enabled=   0   'False
            Tab(0).Control(38)=   "txtTara"
            Tab(0).Control(38).Enabled=   0   'False
            Tab(0).Control(39)=   "txtPezzi"
            Tab(0).Control(39).Enabled=   0   'False
            Tab(0).Control(40)=   "txtQta_UM"
            Tab(0).Control(40).Enabled=   0   'False
            Tab(0).Control(41)=   "CDCodiceImballo"
            Tab(0).Control(41).Enabled=   0   'False
            Tab(0).Control(42)=   "cboUM"
            Tab(0).Control(42).Enabled=   0   'False
            Tab(0).Control(43)=   "txtTaraUnitaria"
            Tab(0).Control(43).Enabled=   0   'False
            Tab(0).Control(44)=   "txtDataLavorazione"
            Tab(0).Control(44).Enabled=   0   'False
            Tab(0).Control(45)=   "CDArticolo"
            Tab(0).Control(45).Enabled=   0   'False
            Tab(0).Control(46)=   "CboTipoLavorazione"
            Tab(0).Control(46).Enabled=   0   'False
            Tab(0).Control(47)=   "cboCalibro"
            Tab(0).Control(47).Enabled=   0   'False
            Tab(0).Control(48)=   "cboTipoCategoria"
            Tab(0).Control(48).Enabled=   0   'False
            Tab(0).Control(49)=   "TxtArticolo"
            Tab(0).Control(49).Enabled=   0   'False
            Tab(0).Control(50)=   "txtLottoVendita"
            Tab(0).Control(50).Enabled=   0   'False
            Tab(0).Control(51)=   "txtImballo"
            Tab(0).Control(51).Enabled=   0   'False
            Tab(0).Control(52)=   "cmdElaborazione"
            Tab(0).Control(52).Enabled=   0   'False
            Tab(0).Control(53)=   "FraCliente"
            Tab(0).Control(53).Enabled=   0   'False
            Tab(0).Control(54)=   "List1"
            Tab(0).Control(54).Enabled=   0   'False
            Tab(0).Control(55)=   "FraOrdineCliente"
            Tab(0).Control(55).Enabled=   0   'False
            Tab(0).Control(56)=   "txtIDProcessoIVGamma"
            Tab(0).Control(56).Enabled=   0   'False
            Tab(0).Control(57)=   "txtOraLav"
            Tab(0).Control(57).Enabled=   0   'False
            Tab(0).Control(58)=   "FraPedana"
            Tab(0).Control(58).Enabled=   0   'False
            Tab(0).Control(59)=   "cmdSelezionaPedana"
            Tab(0).Control(59).Enabled=   0   'False
            Tab(0).Control(60)=   "cmdNuovaPedana"
            Tab(0).Control(60).Enabled=   0   'False
            Tab(0).Control(61)=   "cmdAssegnazioneMerce"
            Tab(0).Control(61).Enabled=   0   'False
            Tab(0).ControlCount=   62
            Begin VB.CommandButton cmdAssegnazioneMerce 
               Height          =   315
               Left            =   7800
               Picture         =   "frmAssegnazioneMerce.frx":47836
               Style           =   1  'Graphical
               TabIndex        =   201
               Top             =   60
               Width           =   375
            End
            Begin VB.CommandButton cmdNuovaPedana 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4920
               Picture         =   "frmAssegnazioneMerce.frx":47DC0
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Nuova pedana"
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton cmdSelezionaPedana 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4080
               Picture         =   "frmAssegnazioneMerce.frx":4834A
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Trova pedana"
               Top             =   120
               Width           =   735
            End
            Begin VB.Frame FraPedana 
               Caption         =   "PEDANA"
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
               TabIndex        =   57
               Top             =   360
               Width           =   5655
               Begin VB.TextBox txtCodicePedana 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   0
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   2415
               End
               Begin VB.TextBox txtArticoloInLinguaPred 
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   63
                  Top             =   1440
                  Width           =   3615
               End
               Begin VB.TextBox txtCalibroInLiguaPred 
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   62
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.TextBox txtCategoriaInLinguaPred 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   61
                  Top             =   2040
                  Width           =   1695
               End
               Begin VB.TextBox txtCodicePostazione 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   60
                  Top             =   2640
                  Width           =   2415
               End
               Begin VB.TextBox txtUtenteMacchina 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   59
                  Top             =   3240
                  Width           =   2415
               End
               Begin VB.TextBox txtNomeMacchina 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   58
                  Top             =   3240
                  Width           =   2895
               End
               Begin DMTLblLinkCtl.LabelLink LabelLink2 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   64
                  Top             =   285
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  Caption         =   "Pedana"
                  Name            =   "LabelLink"
               End
               Begin DMTDataCmb.DMTCombo cboTipoPedana 
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   1
                  TabStop         =   0   'False
                  Top             =   480
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
               Begin DMTEDITNUMLib.dmtNumber txtIDPedana 
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   65
                  Top             =   2040
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   503
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboLinguaPred 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1440
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
               Begin DMTDataCmb.DMTCombo cboUtente 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   67
                  Top             =   2640
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin VB.Label Label2 
                  Caption         =   "Tipo di pedana"
                  Height          =   255
                  Index           =   7
                  Left            =   2640
                  TabIndex        =   77
                  Top             =   285
                  Width           =   2895
               End
               Begin VB.Label Label7 
                  Caption         =   "Lingua predefinita"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   76
                  Top             =   1200
                  Width           =   1695
               End
               Begin VB.Line Line5 
                  X1              =   120
                  X2              =   5520
                  Y1              =   1080
                  Y2              =   1080
               End
               Begin VB.Label Label7 
                  Caption         =   "Utente"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   75
                  Top             =   2400
                  Width           =   2415
               End
               Begin VB.Label Label11 
                  Caption         =   "Descrizione articolo in lingua predefinita"
                  Height          =   255
                  Index           =   3
                  Left            =   1920
                  TabIndex        =   74
                  Top             =   1200
                  Width           =   3615
               End
               Begin VB.Label Label11 
                  Caption         =   "Calibro in lingua "
                  Height          =   255
                  Index           =   4
                  Left            =   1920
                  TabIndex        =   73
                  Top             =   1800
                  Width           =   1455
               End
               Begin VB.Label Label11 
                  Caption         =   "Categoria in lingua "
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   72
                  Top             =   1800
                  Width           =   1695
               End
               Begin VB.Label Label7 
                  Caption         =   "Codice postazione"
                  Height          =   255
                  Index           =   6
                  Left            =   3120
                  TabIndex        =   71
                  Top             =   2400
                  Width           =   2415
               End
               Begin VB.Label Label5 
                  Caption         =   "Utente macchina"
                  Height          =   255
                  Index           =   1
                  Left            =   3120
                  TabIndex        =   70
                  Top             =   3000
                  Width           =   2295
               End
               Begin VB.Label Label5 
                  Caption         =   "Nome macchina"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   69
                  Top             =   3000
                  Width           =   2655
               End
               Begin VB.Label Label11 
                  Caption         =   "Identificativo pedana"
                  Height          =   255
                  Index           =   6
                  Left            =   3720
                  TabIndex        =   68
                  Top             =   1800
                  Width           =   1815
               End
            End
            Begin VB.TextBox txtOraLav 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   179
               TabStop         =   0   'False
               Top             =   1965
               Width           =   975
            End
            Begin DMTEDITNUMLib.dmtNumber txtIDProcessoIVGamma 
               Height          =   255
               Left            =   5640
               TabIndex        =   178
               Top             =   120
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
            Begin VB.Frame FraOrdineCliente 
               Caption         =   "ORDINE CLIENTE"
               Enabled         =   0   'False
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
               Height          =   1095
               Left            =   5880
               TabIndex        =   162
               Top             =   120
               Width           =   9255
               Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   6
                  Top             =   435
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DmtCodDescCtl.DmtCodDesc cdCliente 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   195
                  Width           =   5295
                  _ExtentX        =   9340
                  _ExtentY        =   1085
                  PropCodice      =   $"frmAssegnazioneMerce.frx":488D4
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmAssegnazioneMerce.frx":48922
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmAssegnazioneMerce.frx":48974
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
                  Left            =   5520
                  TabIndex        =   5
                  Top             =   435
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNListaPrelievo 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   193
                  Top             =   435
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin VB.Label Label5 
                  Caption         =   "N° lista"
                  Height          =   255
                  Index           =   6
                  Left            =   8280
                  TabIndex        =   194
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label5 
                  Caption         =   "Numero "
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   165
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "Data ordine"
                  Height          =   255
                  Index           =   0
                  Left            =   5520
                  TabIndex        =   164
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label lblStatoOrdine 
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
                  Height          =   255
                  Left            =   120
                  TabIndex        =   163
                  Top             =   780
                  Width           =   8775
               End
            End
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               Height          =   4515
               Left            =   120
               TabIndex        =   167
               Top             =   3720
               Width           =   7815
            End
            Begin VB.Frame FraCliente 
               Caption         =   "ALTRI DATI CLIENTE"
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
               Height          =   5535
               Left            =   8040
               TabIndex        =   41
               Top             =   2880
               Width           =   7095
               Begin VB.TextBox txtRaggrOrd 
                  Height          =   315
                  Left            =   4080
                  TabIndex        =   196
                  Top             =   5040
                  Width           =   2895
               End
               Begin VB.TextBox txtLottoCliente 
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
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   3975
               End
               Begin VB.TextBox txtAltreAnnotazioniCliente 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   3480
                  Width           =   6855
               End
               Begin VB.CheckBox chkPrezzoInclusoImballo 
                  Caption         =   "Incluso imballo"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   4680
                  Width           =   1695
               End
               Begin VB.TextBox txtCodiceABarreArticoloCliente 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   2280
                  Width           =   2415
               End
               Begin VB.TextBox txtCodiceGSICliente 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.TextBox txtDescrizioneArticoloInLinguaCliente 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   4815
               End
               Begin VB.TextBox txtCategoriaInLingua 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtCalibroInLingua 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtCodificaArticoloCodiceABarre 
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   43
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox txtCodiceABarreImballoCliente 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   2280
                  Width           =   2415
               End
               Begin VB.TextBox txtCodificaImballoCodiceABarre 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   42
                  Top             =   2280
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox txtCodiceAssociatoPressoCliente 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   2880
                  Width           =   1935
               End
               Begin VB.TextBox txtCodificaCodicePedanaCliente 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2160
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   2880
                  Width           =   2895
               End
               Begin DMTDataCmb.DMTCombo cboLinguaCliente 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin DMTEDITNUMLib.dmtNumber txtCodicePedanaCliente 
                  Height          =   315
                  Left            =   5040
                  TabIndex        =   28
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   168
                  Top             =   4320
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
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
               Begin DMTDataCmb.DMTCombo cboListino 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   169
                  Top             =   4320
                  Width           =   2295
                  _ExtentX        =   4048
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
               Begin DMTEDITNUMLib.dmtNumber txtSconto2 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   170
                  Top             =   4320
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecimalPlaces   =   1
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtSconto1 
                  Height          =   315
                  Left            =   4080
                  TabIndex        =   171
                  Top             =   4320
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecimalPlaces   =   1
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioImballo 
                  Height          =   315
                  Left            =   5520
                  TabIndex        =   172
                  Top             =   4320
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
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
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdineCliente 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   189
                  Top             =   1650
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
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
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdinePadre 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   191
                  Top             =   2280
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
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDRigaOrdine 
                  Height          =   315
                  Left            =   480
                  TabIndex        =   195
                  Top             =   5040
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  DecimalPlaces   =   0
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioListino 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   198
                  Top             =   5040
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
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
               Begin VB.Label Label4 
                  Caption         =   "ID"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   200
                  Top             =   5040
                  Width           =   375
               End
               Begin VB.Label Label4 
                  Caption         =   "Prezzo listino"
                  Height          =   255
                  Index           =   16
                  Left            =   2520
                  TabIndex        =   199
                  Top             =   4800
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  Caption         =   "Raggruppamento ordine"
                  Height          =   255
                  Index           =   20
                  Left            =   4080
                  TabIndex        =   197
                  Top             =   4800
                  Width           =   2655
               End
               Begin VB.Label Label4 
                  Caption         =   "ID Padre"
                  Height          =   255
                  Index           =   1
                  Left            =   5280
                  TabIndex        =   192
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Line Line6 
                  X1              =   5160
                  X2              =   5160
                  Y1              =   1440
                  Y2              =   3240
               End
               Begin VB.Label Label4 
                  Caption         =   "ID Ord. cliente"
                  Height          =   255
                  Index           =   2
                  Left            =   5280
                  TabIndex        =   190
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Caption         =   "Listino"
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   177
                  Top             =   4080
                  Width           =   1695
               End
               Begin VB.Label Label4 
                  Caption         =   "Imp. Uni. Art."
                  Height          =   255
                  Index           =   15
                  Left            =   2520
                  TabIndex        =   176
                  Top             =   4080
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Sc.1 %"
                  Height          =   255
                  Index           =   17
                  Left            =   4080
                  TabIndex        =   175
                  Top             =   4080
                  Width           =   615
               End
               Begin VB.Label Label4 
                  Caption         =   "Sc.2 %"
                  Height          =   255
                  Index           =   18
                  Left            =   4800
                  TabIndex        =   174
                  Top             =   4080
                  Width           =   615
               End
               Begin VB.Label Label4 
                  Caption         =   "Imp. Uni. Imb"
                  Height          =   255
                  Index           =   19
                  Left            =   5520
                  TabIndex        =   173
                  Top             =   4080
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  Caption         =   "Lotto cliente"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   2415
               End
               Begin VB.Label Label7 
                  Caption         =   "Altre annotazioni per il  cliente"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   55
                  Top             =   3240
                  Width           =   3735
               End
               Begin VB.Label Label7 
                  Caption         =   "Lingua cliente"
                  Height          =   255
                  Index           =   3
                  Left            =   4200
                  TabIndex        =   54
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label Label6 
                  Caption         =   "Codice a barre articolo"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   53
                  Top             =   2040
                  Width           =   2415
               End
               Begin VB.Label Label6 
                  Caption         =   "Pedana cliente"
                  Height          =   255
                  Index           =   1
                  Left            =   5040
                  TabIndex        =   52
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.Label Label6 
                  Caption         =   "Codice G.S.I."
                  Height          =   255
                  Index           =   2
                  Left            =   3720
                  TabIndex        =   51
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.Label Label11 
                  Caption         =   "Descrizione articolo in lingua"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   50
                  Top             =   840
                  Width           =   4575
               End
               Begin VB.Label Label11 
                  Caption         =   "Categoria in lingua "
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.Label Label11 
                  Caption         =   "Calibro in lingua "
                  Height          =   255
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   48
                  Top             =   1440
                  Width           =   1455
               End
               Begin VB.Label Label6 
                  Caption         =   "Codice a barre imballo"
                  Height          =   255
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   47
                  Top             =   2040
                  Width           =   2415
               End
               Begin VB.Label Label6 
                  Caption         =   "Codice associato"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   46
                  Top             =   2640
                  Width           =   1935
               End
               Begin VB.Label Label6 
                  Caption         =   "Codice pedana"
                  Height          =   255
                  Index           =   5
                  Left            =   2160
                  TabIndex        =   45
                  Top             =   2640
                  Width           =   2295
               End
            End
            Begin VB.CommandButton cmdElaborazione 
               Caption         =   "ELABORA"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   38
               Top             =   2880
               Width           =   7815
            End
            Begin VB.TextBox txtImballo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   1440
               Width           =   3015
            End
            Begin VB.TextBox txtLottoVendita 
               Height          =   315
               Left            =   8040
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   2520
               Width           =   3855
            End
            Begin VB.TextBox TxtArticolo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1440
               Width           =   3615
            End
            Begin DMTDataCmb.DMTCombo cboTipoCategoria 
               Height          =   315
               Left            =   4920
               TabIndex        =   16
               Top             =   1965
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
            Begin DMTDataCmb.DMTCombo cboCalibro 
               Height          =   315
               Left            =   6480
               TabIndex        =   44
               Top             =   1965
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
            Begin DMTDataCmb.DMTCombo CboTipoLavorazione 
               Height          =   315
               Left            =   2880
               TabIndex        =   15
               Top             =   1965
               Width           =   1935
               _ExtentX        =   3413
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
            Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
               Height          =   615
               Left            =   120
               TabIndex        =   7
               Top             =   1170
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   1085
               PropCodice      =   $"frmAssegnazioneMerce.frx":489CE
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmAssegnazioneMerce.frx":48A1D
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmAssegnazioneMerce.frx":48A75
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
            Begin DMTDATETIMELib.dmtDate txtDataLavorazione 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   1965
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Appearance      =   1
            End
            Begin DMTEDITNUMLib.dmtNumber txtTaraUnitaria 
               Height          =   315
               Left            =   13680
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Enabled         =   0   'False
               Appearance      =   1
               UseSeparator    =   -1  'True
               DecimalPlaces   =   5
               DecFinalZeros   =   -1  'True
            End
            Begin DMTDataCmb.DMTCombo cboUM 
               Height          =   315
               Left            =   5640
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   1440
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
               TabIndex        =   10
               Top             =   1170
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   1085
               PropCodice      =   $"frmAssegnazioneMerce.frx":48ACF
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmAssegnazioneMerce.frx":48B26
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmAssegnazioneMerce.frx":48B7D
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
            Begin DMTEDITNUMLib.dmtNumber txtQta_UM 
               Height          =   315
               Left            =   6720
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Left            =   5400
               TabIndex        =   22
               Top             =   2520
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   253
               Text            =   "0"
               BackColor       =   16777215
               Appearance      =   1
               UseSeparator    =   -1  'True
               DecFinalZeros   =   -1  'True
               AllowEmpty      =   0   'False
            End
            Begin DMTEDITNUMLib.dmtNumber txtTara 
               Height          =   315
               Left            =   2760
               TabIndex        =   20
               Top             =   2520
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Left            =   4080
               TabIndex        =   21
               Top             =   2520
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Left            =   1440
               TabIndex        =   19
               Top             =   2520
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               TabIndex        =   18
               Top             =   2505
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   253
               Text            =   "0"
               BackColor       =   16777215
               Appearance      =   1
               UseSeparator    =   -1  'True
               DecFinalZeros   =   -1  'True
               AllowEmpty      =   0   'False
            End
            Begin DMTDataCmb.DMTCombo cboUMImballo 
               Height          =   315
               Left            =   12000
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1575
               _ExtentX        =   2778
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
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   255
               Left            =   120
               TabIndex        =   166
               Top             =   3360
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.CheckBox chkScarti 
               Caption         =   "Inserisci differenza nella quadratura"
               Enabled         =   0   'False
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
               Height          =   195
               Left            =   120
               TabIndex        =   188
               Top             =   120
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.Frame fraScarti 
               Caption         =   "QUADRATURA"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   1335
               Left            =   120
               TabIndex        =   181
               Top             =   2880
               Visible         =   0   'False
               Width           =   7815
               Begin VB.CommandButton cmdArticoliDerScarti 
                  Height          =   300
                  Left            =   5340
                  Picture         =   "frmAssegnazioneMerce.frx":48BD7
                  Style           =   1  'Graphical
                  TabIndex        =   187
                  ToolTipText     =   "Articoli derivati per quadratura"
                  Top             =   180
                  Width           =   375
               End
               Begin DmtCodDescCtl.DmtCodDesc CDArticoloScarto 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   182
                  Top             =   210
                  Width           =   5655
                  _ExtentX        =   9975
                  _ExtentY        =   1085
                  PropCodice      =   $"frmAssegnazioneMerce.frx":49161
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmAssegnazioneMerce.frx":491B0
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmAssegnazioneMerce.frx":4921E
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
               Begin DMTDataCmb.DMTCombo CboTipoLavScarti 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   183
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
               Begin DMTDataCmb.DMTCombo cboUMScarti 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   185
                  TabStop         =   0   'False
                  Top             =   465
                  Width           =   1815
                  _ExtentX        =   3201
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
                  Caption         =   "Unità di misura"
                  Height          =   255
                  Index           =   13
                  Left            =   5880
                  TabIndex        =   186
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Label2 
                  Caption         =   "Tipo lavorazione"
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   184
                  Top             =   765
                  Width           =   2175
               End
            End
            Begin DmtCodDescCtl.DmtCodDesc CDImballoPrimario 
               Height          =   615
               Left            =   8040
               TabIndex        =   202
               Top             =   1710
               Width           =   4890
               _ExtentX        =   8625
               _ExtentY        =   1085
               PropCodice      =   $"frmAssegnazioneMerce.frx":49278
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmAssegnazioneMerce.frx":492C7
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmAssegnazioneMerce.frx":4932F
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
            Begin DMTEDITNUMLib.dmtNumber txtTaraConfImballo 
               Height          =   315
               Left            =   14040
               TabIndex        =   203
               Top             =   1965
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
            Begin DMTEDITNUMLib.dmtNumber txtNumeroConfImballo 
               Height          =   315
               Left            =   12960
               TabIndex        =   204
               Top             =   1965
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
               DecimalPlaces   =   0
               DecFinalZeros   =   -1  'True
               AllowEmpty      =   0   'False
            End
            Begin VB.CheckBox chkConfermaDaUtente 
               Caption         =   "Conferma da utente"
               Height          =   255
               Left            =   3480
               TabIndex        =   207
               Top             =   4680
               Width           =   2535
            End
            Begin VB.CheckBox chkTracciaImballoGest 
               Caption         =   "Traccia"
               Height          =   255
               Left            =   240
               TabIndex        =   208
               Top             =   4680
               Width           =   975
            End
            Begin VB.CheckBox chkTracciaImballoGestPrim 
               Caption         =   "Traccia"
               Height          =   255
               Left            =   240
               TabIndex        =   209
               Top             =   5040
               Width           =   975
            End
            Begin VB.CheckBox chkConfermaDaUtentePrim 
               Caption         =   "Conferma da utente"
               Height          =   255
               Left            =   3480
               TabIndex        =   210
               Top             =   5040
               Width           =   2535
            End
            Begin DMTEDITNUMLib.dmtNumber txtCostoConfezione 
               Height          =   315
               Left            =   14040
               TabIndex        =   211
               Top             =   2640
               Visible         =   0   'False
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
            Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerCollo 
               Height          =   315
               Left            =   12000
               TabIndex        =   212
               Top             =   2520
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
               DecimalPlaces   =   0
               DecFinalZeros   =   -1  'True
               AllowEmpty      =   0   'False
            End
            Begin DMTEDITNUMLib.dmtNumber txtPesoPerCollo 
               Height          =   315
               Left            =   13800
               TabIndex        =   213
               Top             =   2520
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
            Begin VB.Label Label2 
               Caption         =   "Q.tà pezzi per collo"
               Height          =   255
               Index           =   15
               Left            =   12000
               TabIndex        =   215
               ToolTipText     =   "Numero di confezioni in un imballo"
               Top             =   2295
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Peso per collo"
               Height          =   255
               Index           =   14
               Left            =   13800
               TabIndex        =   214
               ToolTipText     =   "Numero di confezioni in un imballo"
               Top             =   2295
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Tara confez."
               Height          =   255
               Index           =   28
               Left            =   14040
               TabIndex        =   206
               ToolTipText     =   "Tara unitario confezione"
               Top             =   1755
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "N° confez."
               Height          =   255
               Index           =   29
               Left            =   12960
               TabIndex        =   205
               ToolTipText     =   "Numero di confezioni in un imballo"
               Top             =   1755
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Ora"
               Height          =   255
               Index           =   24
               Left            =   1800
               TabIndex        =   180
               Top             =   1755
               Width           =   975
            End
            Begin VB.Label Label8 
               Caption         =   "Tipo categoria"
               Height          =   255
               Index           =   1
               Left            =   4920
               TabIndex        =   93
               Top             =   1755
               Width           =   1455
            End
            Begin VB.Label Label8 
               Caption         =   "Calibro"
               Height          =   255
               Index           =   0
               Left            =   6480
               TabIndex        =   92
               Top             =   1755
               Width           =   1455
            End
            Begin VB.Label Label7 
               Caption         =   "Lotto articolo di vendita"
               Height          =   255
               Index           =   0
               Left            =   8040
               TabIndex        =   91
               Top             =   2295
               Width           =   3855
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo lavorazione"
               Height          =   255
               Index           =   6
               Left            =   2880
               TabIndex        =   90
               Top             =   1755
               Width           =   2055
            End
            Begin VB.Label lblArticolo 
               Caption         =   "Articolo"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1920
               TabIndex        =   89
               Top             =   1215
               Width           =   3615
            End
            Begin VB.Label Label2 
               Caption         =   "Colli"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   88
               Top             =   2295
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Peso lordo"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   87
               Top             =   2295
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Pezzi"
               Height          =   255
               Index           =   2
               Left            =   5400
               TabIndex        =   86
               Top             =   2295
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Tara "
               Height          =   255
               Index           =   3
               Left            =   2760
               TabIndex        =   85
               Top             =   2295
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Peso netto"
               Height          =   255
               Index           =   4
               Left            =   4080
               TabIndex        =   84
               Top             =   2295
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Q.tà"
               Height          =   255
               Index           =   5
               Left            =   6720
               TabIndex        =   83
               Top             =   2295
               Width           =   1215
            End
            Begin VB.Label lblImballo 
               Caption         =   "Imballo"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8880
               TabIndex        =   82
               Top             =   1200
               Width           =   2535
            End
            Begin VB.Label Label2 
               Caption         =   "Unità di misura"
               Height          =   255
               Index           =   9
               Left            =   5640
               TabIndex        =   81
               Top             =   1215
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Tara unitaria"
               Height          =   255
               Index           =   10
               Left            =   13680
               TabIndex        =   80
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Data lavorazione"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   79
               Top             =   1755
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Unità di misura"
               Height          =   255
               Index           =   11
               Left            =   12000
               TabIndex        =   78
               Top             =   1215
               Width           =   1575
            End
         End
         Begin VB.Frame FraConferimento 
            Caption         =   "ARTICOLO DI CONFERIMENTO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   5520
            TabIndex        =   111
            Top             =   480
            Width           =   9015
            Begin VB.TextBox txtArticoloImballo_Conferimento 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5160
               TabIndex        =   131
               Top             =   960
               Width           =   3735
            End
            Begin VB.TextBox txtArticolo_Conf 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2520
               TabIndex        =   130
               Top             =   480
               Width           =   3255
            End
            Begin VB.TextBox txtCodiceArticolo_Conf 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   129
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox txtCodiceLotto_Conf 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5880
               TabIndex        =   128
               Top             =   480
               Width           =   3015
            End
            Begin VB.TextBox txtCodiceImballo_Conferimento 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3600
               TabIndex        =   127
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtUnitaDiMisura_Conferimento 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   126
               Top             =   1440
               Width           =   1695
            End
            Begin VB.CommandButton cmdApriFinestraConferimento 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   8620
               Picture         =   "frmAssegnazioneMerce.frx":49389
               Style           =   1  'Graphical
               TabIndex        =   125
               ToolTipText     =   "Nuova ordine"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtLottoEntrata 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   124
               Top             =   960
               Width           =   3375
            End
            Begin VB.TextBox txtRegioneProvenienza 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   123
               Top             =   2880
               Width           =   2655
            End
            Begin VB.TextBox txtNazioneProvenienza 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2880
               TabIndex        =   122
               Top             =   2880
               Width           =   2655
            End
            Begin VB.TextBox txtComuneProvenienza 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   5640
               TabIndex        =   121
               Top             =   2880
               Width           =   2655
            End
            Begin VB.TextBox txtProvinciaProvenienza 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   8400
               TabIndex        =   120
               Top             =   2880
               Width           =   495
            End
            Begin VB.TextBox txtFamigliaLottoCampagna 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   119
               Top             =   3840
               Width           =   2175
            End
            Begin VB.TextBox txtVarietaLottoCampagna 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   118
               Top             =   3840
               Width           =   2175
            End
            Begin VB.TextBox txtDataDiSblocco 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   6960
               TabIndex        =   117
               Top             =   3840
               Width           =   1935
            End
            Begin VB.TextBox txtTipoProduzione 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4680
               TabIndex        =   116
               Top             =   3840
               Width           =   2175
            End
            Begin VB.TextBox txtCodiceAssociatoPredefinito 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   115
               Top             =   4920
               Width           =   2055
            End
            Begin VB.TextBox txtBNDOO 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2280
               TabIndex        =   114
               Top             =   4920
               Width           =   2055
            End
            Begin DMTEDITNUMLib.dmtNumber txtIDImballoConferito 
               Height          =   255
               Left            =   960
               TabIndex        =   112
               Top             =   1920
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTEDITNUMLib.dmtNumber txtIDArticoloConferito 
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   1920
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTEDITNUMLib.dmtNumber txtQtaDifferenza 
               Height          =   285
               Left            =   6480
               TabIndex        =   132
               Top             =   1920
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   253
               Text            =   "0"
               ForeColor       =   65535
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
            Begin DMTEDITNUMLib.dmtNumber txtQtaQuadrata 
               Height          =   285
               Left            =   4800
               TabIndex        =   133
               Top             =   1440
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   253
               Text            =   "0"
               ForeColor       =   65535
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
            Begin DMTEDITNUMLib.dmtNumber txtQtaVenduta 
               Height          =   285
               Left            =   6480
               TabIndex        =   134
               Top             =   1440
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   253
               Text            =   "0"
               ForeColor       =   65535
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
            Begin DMTEDITNUMLib.dmtNumber txtDisponibiltaLottoCaricato 
               Height          =   285
               Left            =   3120
               TabIndex        =   135
               Top             =   1440
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   253
               Text            =   "0"
               ForeColor       =   65535
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
            Begin DMTEDITNUMLib.dmtNumber txtNumeroColliCaricati 
               Height          =   285
               Left            =   120
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   503
               _StockProps     =   253
               Text            =   "0"
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   1
               Enabled         =   0   'False
               AllowEmpty      =   0   'False
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "ALTRI DATI PREDEFINITI PER FILIALE"
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
               Left            =   120
               TabIndex        =   161
               Top             =   4320
               Width           =   8775
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "CARATTERISTICHE LOTTO DI CAMPAGNA"
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
               Left            =   120
               TabIndex        =   160
               Top             =   3360
               Width           =   8775
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "PROVENIENZA LOTTO DI CAMPAGNA"
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
               TabIndex        =   159
               Top             =   2400
               Width           =   8775
            End
            Begin VB.Label Label3 
               Caption         =   "Imballo"
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   158
               Top             =   765
               Width           =   3735
            End
            Begin VB.Label Label3 
               Caption         =   "Codice articolo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   157
               Top             =   280
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Articolo"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   156
               Top             =   280
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Lotto di campagna"
               Height          =   255
               Index           =   2
               Left            =   5880
               TabIndex        =   155
               Top             =   280
               Width           =   2295
            End
            Begin VB.Label Label3 
               Caption         =   "Colli caricati"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   154
               Top             =   1245
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Lotto di entrata (lotto di conferimento)"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   153
               Top             =   765
               Width           =   3375
            End
            Begin VB.Label Label3 
               Caption         =   "Q.tà conferita"
               Height          =   255
               Index           =   9
               Left            =   3120
               TabIndex        =   152
               Top             =   1245
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Differenza"
               Height          =   255
               Index           =   16
               Left            =   6480
               TabIndex        =   151
               Top             =   1725
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Q.tà Assegnata"
               Height          =   255
               Index           =   17
               Left            =   6480
               TabIndex        =   150
               Top             =   1245
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   "Q.tà Quadrata"
               Height          =   255
               Index           =   18
               Left            =   4800
               TabIndex        =   149
               Top             =   1245
               Width           =   1575
            End
            Begin VB.Image Image1 
               Height          =   240
               Left            =   8040
               Picture         =   "frmAssegnazioneMerce.frx":49913
               ToolTipText     =   "ATTENZIONE: ERRORE DI QUADRATURA"
               Top             =   1920
               Width           =   240
            End
            Begin VB.Label Label3 
               Caption         =   "Codice imballo"
               Height          =   255
               Index           =   5
               Left            =   3600
               TabIndex        =   148
               Top             =   765
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Unità di misura"
               Height          =   255
               Index           =   4
               Left            =   1320
               TabIndex        =   147
               Top             =   1245
               Width           =   1575
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   8880
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Label Label10 
               Caption         =   "Regione"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   146
               Top             =   2640
               Width           =   2655
            End
            Begin VB.Label Label10 
               Caption         =   "Nazione"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   145
               Top             =   2640
               Width           =   2655
            End
            Begin VB.Label Label10 
               Caption         =   "Comune"
               Height          =   255
               Index           =   2
               Left            =   5640
               TabIndex        =   144
               Top             =   2640
               Width           =   2655
            End
            Begin VB.Label Label10 
               Caption         =   "Prov."
               Height          =   255
               Index           =   3
               Left            =   8400
               TabIndex        =   143
               Top             =   2640
               Width           =   495
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   8880
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Label Label10 
               Caption         =   "Famiglia"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   142
               Top             =   3600
               Width           =   2655
            End
            Begin VB.Label Label10 
               Caption         =   "Varietà"
               Height          =   255
               Index           =   5
               Left            =   2400
               TabIndex        =   141
               Top             =   3600
               Width           =   2655
            End
            Begin VB.Label Label10 
               Caption         =   "Data di sblocco"
               Height          =   255
               Index           =   6
               Left            =   6960
               TabIndex        =   140
               Top             =   3600
               Width           =   1935
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   8880
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Label Label10 
               Caption         =   "Tipo di produzione"
               Height          =   255
               Index           =   7
               Left            =   4680
               TabIndex        =   139
               Top             =   3600
               Width           =   2175
            End
            Begin VB.Label Label1 
               Caption         =   "Codice associato pred."
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   138
               Top             =   4680
               Width           =   2055
            End
            Begin VB.Label Label1 
               Caption         =   "B.N.D.O.O."
               Height          =   255
               Index           =   13
               Left            =   2280
               TabIndex        =   137
               Top             =   4680
               Width           =   2055
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "DOCUMENTO DI CONFERIMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   360
            TabIndex        =   95
            Top             =   720
            Width           =   5655
            Begin VB.TextBox TxtIDAnagrafica 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4200
               TabIndex        =   99
               Top             =   120
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtNomeSocio 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   380
               Width           =   1215
            End
            Begin VB.TextBox txtSocio 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   720
               TabIndex        =   97
               Top             =   380
               Width           =   3495
            End
            Begin VB.TextBox txtCodiceSocio 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   96
               Top             =   380
               Width           =   615
            End
            Begin DMTDATETIMELib.dmtDate txtDataDocumento 
               Height          =   315
               Left            =   4200
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   1400
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   253
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   1
               Enabled         =   0   'False
            End
            Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
               Height          =   315
               Left            =   3000
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   1400
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   253
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   1
               Enabled         =   0   'False
            End
            Begin DMTDataCmb.DMTCombo CboMagazzinoVend 
               Height          =   315
               Left            =   3000
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   900
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               BackColor       =   -2147483633
               Enabled         =   0   'False
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
            Begin DMTDataCmb.DMTCombo cboMagazzinoConf 
               Height          =   315
               Left            =   120
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   900
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   582
               BackColor       =   -2147483633
               Enabled         =   0   'False
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
            Begin DMTDataCmb.DMTCombo cboSezionale 
               Height          =   315
               Left            =   120
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   1395
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   582
               BackColor       =   -2147483633
               Enabled         =   0   'False
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
            Begin VB.Label Label1 
               Caption         =   "Socio"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   110
               Top             =   195
               Width           =   5295
            End
            Begin VB.Label Label1 
               Caption         =   "Magazzino di vendita"
               Height          =   255
               Index           =   4
               Left            =   3000
               TabIndex        =   109
               Top             =   690
               Width           =   2415
            End
            Begin VB.Label Label1 
               Caption         =   "Magazzino di conferimento"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   108
               Top             =   690
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "Sezionale"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   107
               Top             =   1200
               Width           =   2535
            End
            Begin VB.Label Label1 
               Caption         =   "Data Doc."
               Height          =   255
               Index           =   1
               Left            =   4200
               TabIndex        =   106
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "N° Doc."
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   105
               Top             =   1200
               Width           =   975
            End
         End
         Begin VB.Label lblInfoStatus 
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
            Left            =   120
            TabIndex        =   94
            Top             =   3120
            Width           =   9255
         End
      End
   End
End
Attribute VB_Name = "frmAssegnazioneMerce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Link_TipoImballo As Long
Private Link_TipoSpezzatura As Long

Private MAX_QUANTITA As Double
Private MAX_QUANTITA_COLLI As Double
Private MAX_QUANTITA_CONFEZ As Double


Private IDAssegnazionePrimaria As Long
Private Link_ProcessoIVGamma As Long

Private Link_Oggetto_Ordine As Long
Private Link_UnitaDiMisura_Coop As Long
Private Link_UnitaDiMisura_Coop_Q As Long
Private Link_TipoProdotto_Q As Long
Private Link_TipoScarto As Long
Private Link_TipoCaloPeso As Long
Private Link_TipoAumentoPeso As Long
Private Link_TipoQuadratura As Long

Private Link_Pedana As Long
Private Link_CaricoMerceRighe As Long

Public Link_Arrontondamento As Long
Public PESO_LORDO_ARTICOLO As Double
Public TIPO_COMPORTAMENTO_LAVORAZIONE As Long
Public TIPO_PESO_ARTICOLO As Long

Private LINK_ANAGRAFICA_SOCIO As Long
Private CODICE_SOCIO As String
Private ANAGRAFICA_SOCIO As String
Private NOME_SOCIO As String
Private DATA_CONFERIMENTO As String
Private NUMERO_CONFERIMENTO As String

Private LINK_ARTICOLO_CONFERITO As Long
Private CODICE_ARTICOLO_CONFERITO As String
Private ARTICOLO_CONFERITO As String
Private LINK_UNITA_DI_MISURA_ARTICOLO_CONFERITO As String
Private CODICE_LOTTO_ENTRATA As String
Private CODICE_LOTTO_CAMPAGNA As String
Private Link_UnitaDiMisura_Coop_Conferimento As Long


Private Link_TipoProdotto As Long
Private mov As DmtMovim.cMovimentazione
Private MAX_CARATTERI_PEDANA_CLIENTE  As Long

Private QUANTITA_PER_COLLO As Double
Private Moltiplicatore As Double
Private bloading_Form As Boolean

Private LINK_SITO_PER_ANAGRAFICA_ORDINE As Long
Private LINK_LISTINO_CLIENTE_ORDINE As Long


Private LINK_PEDANA_LAVORAZIONE As Long
Private CODICE_PEDANA_LAVORAZIONE As String


Private Link_OggettoOrdinePrec As Long
Private Link_ProcessoLavorazione As Long
Private Link_ProcessoLavorazioneRighe As Long
Private Link_LineaProduzione As Long
Private Link_CaricoMerceRighePrelievi As Long
Private Link_TipoUtilizzoLinea As Long
Private Link_LottoCampagnaSuLotto As Long
Private IsPreConferimento As Long


'VARIABILI PER LA CONFIGURAZIONE DELL'OGGETTO DOCUMENTO PER PRELEVARE I PREZZI
Private ObjDoc As DmtDocs.cDocument
Private sTabellaTestataLocal As String
Private sTabellaDettaglioLocal As String
Private sTabellaIVALocal As String
Private sTabellaScadenzeLocal As String

Public rsLottoImballoPrim As ADODB.Recordset
Public rsLottoImballo As ADODB.Recordset
Public rsKIT As ADODB.Recordset
Public Function PermessoSalvataggio() As Boolean
PermessoSalvataggio = True

If Me.txtIDPedana.Value = 0 Then
    MsgBox "La pedana non esiste o non è stata inserita", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtCodicePedana.SetFocus
    Exit Function
End If

If Me.txtQta_UM.Value <= 0 Then
    MsgBox "La quantità deve essere maggiore di zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtQta_UM.SetFocus
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
End Function

Private Sub cboUMScarti_Click()
    Link_UnitaDiMisura_Coop_Q = fnGetUMCoop(Me.cboUMScarti.CurrentID)
End Sub
Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub CDArticolo_ChangeElement()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

    sSQL = "SELECT RV_POQuantitaPerCollo, RV_POMoltiplicatore, PesoNetto, RV_POIDTipoPesoArticolo "
    sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        QUANTITA_PER_COLLO = fnNotNullN(rs!RV_POQuantitaPerCollo)
        Moltiplicatore = fnNotNullN(rs!RV_POMoltiplicatore)
        If fnNotNullN(rs!PesoNetto) > 0 Then
            PESO_LORDO_ARTICOLO = fnNotNullN(rs!PesoNetto)
        End If
        If fnNotNullN(rs!RV_POIDTipoPesoArticolo) > 0 Then TIPO_PESO_ARTICOLO = fnNotNullN(rs!RV_POIDTipoPesoArticolo)
    End If


rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CDArticoloScarto_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraVendita, IDTipoProdotto, RV_POIDTipoLavorazione "
sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & Me.CDArticoloScarto.KeyFieldID

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.cboUMScarti.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
    Link_TipoProdotto_Q = fnNotNullN(rs!IDTipoProdotto)
    Me.CboTipoLavScarti.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub cdCliente_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
''''''''''''''''''LINGUA DESCRIZIONE ARTICOLI''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDLinguaDescrizioneArticoli "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & Me.cdCliente.KeyFieldID

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboLinguaCliente.WriteOn GET_LINK_LINGUA_PREDEFINITA
Else
    If fnNotNullN(rs!IDLinguaDescrizioneArticoli) > 0 Then
        Me.cboLinguaCliente.WriteOn fnNotNullN(rs!IDLinguaDescrizioneArticoli)
    Else
        Me.cboLinguaCliente.WriteOn GET_LINK_LINGUA_PREDEFINITA
    End If
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

GET_INFO_CLIENTE Me.cdCliente.KeyFieldID
Me.chkPrezzoInclusoImballo.Value = GET_PREZZO_IMBALLO_INCLUSO_2(Me.CDArticolo.KeyFieldID, Me.txtIDOrdineCliente.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.cdCliente.KeyFieldID, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID)
GET_IMBALLO_CODICE_A_BARRE_CLIENTE Me.cdCliente.KeyFieldID, Me.CDCodiceImballo.KeyFieldID
GET_ARTICOLO_CODICE_A_BARRE_CLIENTE Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID

Me.txtDescrizioneArticoloInLinguaCliente.Text = GET_DESCRIZIONE_IN_LINGUA(Me.cboLinguaCliente.CurrentID, Me.CDArticolo.KeyFieldID)
Me.txtCategoriaInLingua.Text = GET_DESCRIZIONE_CATEGORIA_IN_LINGUA(Me.cboLinguaCliente.CurrentID, Me.cboTipoCategoria.CurrentID)
Me.txtCalibroInLingua.Text = GET_DESCRIZIONE_CALIBRO_IN_LINGUA(Me.cboLinguaCliente.CurrentID, Me.cboCalibro.CurrentID)



If Me.cdCliente.KeyFieldID = 0 Then
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
End If
End Sub

Private Sub CDCodiceImballo_ChangeElement()
    If Me.CDCodiceImballo.KeyFieldID > 0 Then
        Me.txtImballo.Text = Me.CDCodiceImballo.Description
        Me.cboUMImballo.WriteOn GET_LINK_UM_ARTICOLO(Me.CDCodiceImballo.KeyFieldID)
                
        If Me.txtTaraUnitaria.Value = 0 Then
            Me.txtTaraUnitaria.Value = fnGetTaraImballo
        End If
        'Me.chkPrezzoInclusoImballo.Value = GET_PREZZO_IMBALLO_INCLUSO(Me.CDCodiceImballo.KeyFieldID, Me.cdCliente.KeyFieldID)
        Me.chkPrezzoInclusoImballo.Value = GET_PREZZO_IMBALLO_INCLUSO_2(Me.CDArticolo.KeyFieldID, Me.txtIDOrdineCliente.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.cdCliente.KeyFieldID, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID)
    Else
        Me.txtImballo.Text = ""
        Me.cboUMImballo.WriteOn 0
        Me.txtTaraUnitaria.Value = 0
    End If

End Sub
Private Function fnGetTaraImballo() As Double
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Tara FROM Articolo WHERE "
    sSQL = sSQL & "IDArticolo = " & Me.CDCodiceImballo.KeyFieldID
    
    Set rs = CnDMT.OpenResultset(sSQL)
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

Private Sub chkScarti_Click()
    If Me.chkScarti.Value = vbChecked Then
        Me.FraPedana.Enabled = False
        'Me.FraOrdineCliente.Enabled = False
        Me.FraCliente.Enabled = False
         
        Link_Pedana = LINK_PEDANA_LAVORAZIONE
        Me.txtIDPedana.Value = LINK_PEDANA_LAVORAZIONE
        Me.txtCodicePedana.Text = CODICE_PEDANA_LAVORAZIONE
        Me.cboTipoPedana.WriteOn GET_TIPO_PEDANA(Me.txtIDPedana.Text)
    
    Else
        Me.FraPedana.Enabled = True
        'Me.FraOrdineCliente.Enabled = True
        Me.FraCliente.Enabled = True
        
        Link_Pedana = 0
        Me.txtIDPedana.Value = 0
        Me.txtCodicePedana.Text = 0
        Me.cboTipoPedana.WriteOn 0
        
    End If
End Sub

Private Sub cmdArticoliDerScarti_Click()
    LINK_CONF_RIGA_PER_SCARTI = Link_CaricoMerceRighe
    frmArticoliDerivatiQ.Show vbModal
    LINK_CONF_RIGA_PER_SCARTI = 0
End Sub

Private Sub cmdAssegnazioneMerce_Click()
    frmSelezionaRigaOrd.Show vbModal
End Sub
Private Sub cmdElaborazione_Click()
On Error GoTo ERR_cmdElaborazione_Click
Dim IDNuovaAssegnazione As Long
Dim Quantita_Rimasta As Long
Dim Colli_Rimasti As Long

If Me.chkScarti.Value = vbUnchecked Then
    If MAX_QUANTITA < Me.txtQta_UM.Value Then
        MsgBox "La quantità da elaborare non deve essere maggiore di " & MAX_QUANTITA, vbCritical, "Assegnazione merce"
        Exit Sub
    End If
    If MAX_QUANTITA_COLLI < Me.txtColli.Value Then
        MsgBox "La quantità di colli da elaborare non deve essere maggiore di " & MAX_QUANTITA_COLLI, vbCritical, "Assegnazione merce"
        Exit Sub
    End If
End If

If Me.chkScarti.Value = vbChecked Then
    If Me.CDArticoloScarto.KeyFieldID = 0 Then
        MsgBox "Inserire l'articolo di quadratura", vbCritical, "Assegnazione merce"
        Exit Sub
    End If
    
    If Me.cboUMScarti.CurrentID = 0 Then
        MsgBox "L'unità di misura per l'articolo di quadratura non è stato configurato", vbCritical, "Assegnazione merce"
        Exit Sub
    End If
    
    If (Link_TipoProdotto_Q <> Link_TipoCaloPeso) And (Link_TipoProdotto_Q <> Link_TipoAumentoPeso) And (Link_TipoProdotto_Q <> Link_TipoScarto) Then
        MsgBox "Il prodotto selezionato non è un prodotto di quadratura", vbCritical, "Assegnazione merce"
        Exit Sub
    End If
End If


Me.List1.Clear

If PermessoSalvataggio = True Then
    
    SCRIVI_CODA 0, fnGetTipoOggetto("RV_POAssegnazioneMerce")

    Screen.MousePointer = 11
    
'''''''''''''''''''''''''CREAZIONE NUOVA ASSEGNAZIONE MERCE''''''''''''''''''''''''''''''''''
    DoEvents
    IDNuovaAssegnazione = fnGetNewKey("RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce")
    
    Me.List1.AddItem "Creazione nuova lavorazione merce"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents
        
    fnInserisciNuovaAssegnazione IDNuovaAssegnazione
    
    Me.List1.AddItem "Creazione nuova lavorazione merce avvenuta con successo"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents
    
'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''
    
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    
    CnDMT.Execute sSQL '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDNuovaAssegnazione > 0 Then
    Me.List1.AddItem "*******Movimentazione dell'assegnazione merce creata (" & IDNuovaAssegnazione & ")********"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents
    
    MOVIMENTAZIONE_RIGA_LAVORAZIONE Link_CaricoMerceRighe, IDNuovaAssegnazione, True
    
    fnAggiornaAssegnazione IDAssegnazionePrimaria
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Max
    
    Screen.MousePointer = 0
    If Not (ObjDoc Is Nothing) Then Set ObjDoc = Nothing
    Unload Me
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If
Exit Sub
ERR_cmdElaborazione_Click:
    If Not (ObjDoc Is Nothing) Then Set ObjDoc = Nothing
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Creazione nuova assegnazione"
    
End Sub
Private Sub fnInserisciNuovaAssegnazione(IDAssegnazione As Long, Optional CodiceLottoVendita As String)
On Error GoTo ERR_fnInserisciNuovaAssegnazione
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDListinoDefault As Long

If Me.txtIDRigaOrdine.Value = 0 Then
    Me.chkPrezzoInclusoImballo.Value = GET_PREZZO_IMBALLO_INCLUSO_2(Me.CDArticolo.KeyFieldID, Me.txtIDOrdinePadre.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.cdCliente.KeyFieldID, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID)
End If

Set rs = New ADODB.Recordset
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=0"
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
        
    rs.AddNew
        If ((Me.txtIDRigaOrdine.Value = 0) And (SPACCATURA_MERCE_VERSO = 0)) Then
            If PREZZI_ARTICOLI_DA_ORDINE = 0 Then
                GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtQta_UM.Value, PREZZI_ARTICOLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID
                GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtQta_UM.Value, PREZZI_IMBALLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID
            Else
                If (GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(Me.CDArticolo.KeyFieldID, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value) = False) Then
                    GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtQta_UM.Value, PREZZI_ARTICOLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID
                    GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtQta_UM.Value, PREZZI_IMBALLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID
                Else
                    If RETURN_SEL_PREZZO_IMB_DA_ORD = 0 Then
                        GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtQta_UM.Value, PREZZI_IMBALLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, Me.txtIDOrdinePadre.Value, Me.txtRaggrOrd.Text, Me.cboCalibro.CurrentID, Me.cboTipoCategoria.CurrentID
                    End If
                End If
            End If
        End If
        
        rs.Fields("IDRV_POAssegnazioneMerce").Value = IDAssegnazione
        rs.Fields("IDRV_POCaricoMerceRighe").Value = Link_CaricoMerceRighe
        
        rs.Fields("DataDocumento").Value = Me.txtDataLavorazione.Text
        rs.Fields("IDArticolo").Value = Me.CDArticolo.KeyFieldID
        rs.Fields("CodiceArticolo").Value = Me.CDArticolo.Code
        rs.Fields("Articolo").Value = Me.TxtArticolo.Text
        rs.Fields("IDUnitaDiMisuraCoop").Value = Link_UnitaDiMisura_Coop
        rs.Fields("IDUnitaDiMisura").Value = Me.cboUM.CurrentID
        rs.Fields("Colli").Value = Me.txtColli.Value
        rs.Fields("PesoLordo").Value = Me.txtPesoLordo.Value
        rs.Fields("PesoNetto").Value = Me.txtPesoNetto.Value
        rs.Fields("TaraUnitaria").Value = Me.txtTaraUnitaria.Value
        rs.Fields("Tara").Value = Me.txtTara.Value
        rs.Fields("Pezzi").Value = Me.txtPezzi.Value
        rs.Fields("Qta_UM").Value = Me.txtQta_UM.Value
        rs.Fields("IDImballoVendita").Value = Me.CDCodiceImballo.KeyFieldID
        rs.Fields("CodiceImballoVendita").Value = Me.CDCodiceImballo.Code
        rs.Fields("ImballoVendita").Value = Me.txtImballo.Text
        rs.Fields("IDTipoLavorazione").Value = Me.CboTipoLavorazione.CurrentID
        rs.Fields("IDRV_POCalibro").Value = Me.cboCalibro.CurrentID
        rs.Fields("IDRV_POTipoCategoria").Value = Me.cboTipoCategoria.CurrentID
        rs.Fields("IDCliente").Value = Me.cdCliente.KeyFieldID
        
        rs.Fields("IDOggettoOrdine").Value = Me.txtIDOrdineCliente.Value
        rs.Fields("NumeroOrdine").Value = Me.txtNumeroOrdine.Value
        rs.Fields("DataOrdine").Value = Me.txtDataOrdine.Value
        rs.Fields("NumeroListaPrelievo").Value = Me.txtNListaPrelievo.Value
        rs.Fields("IDOggettoOrdinePadre").Value = Me.txtIDOrdinePadre.Value
        
        rs.Fields("IDRV_POPedana").Value = Me.txtIDPedana.Value
        rs.Fields("CodicePedana").Value = Me.txtCodicePedana.Text
        rs.Fields("CodiceLottoVendita").Value = Me.txtLottoVendita.Text
        
        rs.Fields("IDAnagraficaSocio").Value = LINK_ANAGRAFICA_SOCIO
        rs.Fields("CodiceSocio").Value = CODICE_SOCIO
        rs.Fields("AnagraficaSocio").Value = ANAGRAFICA_SOCIO
        rs.Fields("NomeSocio").Value = NOME_SOCIO
        If IsDate(DATA_CONFERIMENTO) Then
            rs.Fields("DataConferimento").Value = DATA_CONFERIMENTO
        End If
        rs.Fields("NumeroConferimento").Value = fnNotNullN(NUMERO_CONFERIMENTO)
        
        rs.Fields("LottoCliente").Value = Me.txtLottoCliente.Text
        rs.Fields("AltreAnnotazioniPerCliente").Value = Me.txtAltreAnnotazioniCliente.Text
        rs.Fields("IDUtente").Value = Me.cboUtente.CurrentID
        rs.Fields("CodiceUtente").Value = Me.txtCodicePostazione.Text
        rs.Fields("NumeroPedanaCliente").Value = Me.txtCodicePedanaCliente.Value
        rs.Fields("MerceInclusoImballo").Value = Me.chkPrezzoInclusoImballo.Value
        rs.Fields("UtentePC").Value = Me.txtNomeMacchina.Text
        rs.Fields("NomePC").Value = Me.txtUtenteMacchina.Text
        rs.Fields("IDLinguaPredefinita").Value = Me.cboLinguaPred.CurrentID
        rs.Fields("IDLinguaCliente").Value = Me.cboLinguaCliente.CurrentID
        rs.Fields("CodicePedanaCliente").Value = Me.txtCodificaCodicePedanaCliente.Text
        rs.Fields("CodiceGSI").Value = Me.txtCodiceGSICliente.Text
        rs.Fields("CodiceAssociatoPressoCliente").Value = Me.txtCodiceAssociatoPressoCliente.Text
        rs.Fields("CodiceABarreArticoloCliente").Value = Me.txtCodificaArticoloCodiceABarre.Text
        rs.Fields("DescrizioneCodiceABarreArticoloCliente").Value = Me.txtCodiceABarreArticoloCliente.Text
        rs.Fields("CodiceABarreImballoCliente").Value = Me.txtCodificaImballoCodiceABarre.Text
        rs.Fields("DescrizioneCodiceABarreImballoCliente").Value = Me.txtCodiceABarreImballoCliente.Text
        rs.Fields("DescrizioneArticoloInLinguaPred").Value = Me.txtArticoloInLinguaPred.Text
        rs.Fields("DescrizioneCalibroInLinguaPred").Value = Me.txtCalibroInLiguaPred.Text
        rs.Fields("DescrizioneCategoriaInLinguaPred").Value = Me.txtCategoriaInLinguaPred.Text
        rs.Fields("DescrizioneArticoloInLinguaCliente").Value = Me.txtDescrizioneArticoloInLinguaCliente.Text
        rs.Fields("DescrizioneCalibroInLinguaCliente").Value = Me.txtCalibroInLingua.Text
        rs.Fields("DescrizioneCategoriaInLinguaCliente").Value = Me.txtCategoriaInLingua.Text
        rs.Fields("CodiceABarreArticoloPred").Value = GET_CODICEABARRE(Me.CDArticolo.KeyFieldID)
        rs.Fields("DescrizioneCodiceABarreArticoloPred").Value = GET_DESCRIZIONECODICEABARRE(Me.CDArticolo.KeyFieldID)
        rs.Fields("CodiceABarreImballoPred").Value = GET_CODICEABARRE(Me.CDCodiceImballo.KeyFieldID)
        rs.Fields("DescrizioneCodiceABarreImballoPred").Value = GET_DESCRIZIONECODICEABARRE(Me.CDCodiceImballo.KeyFieldID)
        rs.Fields("LinguaPredefinita").Value = Me.cboLinguaPred.Text
        rs.Fields("LinguaCliente").Value = Me.cboLinguaCliente.Text
        rs.Fields("Link_Ordinamento").Value = GET_LINK_ORDINAMENTO("RV_POAssegnazioneMerce", "IDRV_POCaricoMerceRighe", Link_CaricoMerceRighe)
        rs!ImportoUnitarioArticolo = Me.txtImportoUnitarioArticolo.Value 'GET_PREZZO_UNITARIO_ARTICOLO(Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, Me.txtIDOrdineCliente.Value, PREZZI_ARTICOLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, rs, LINK_SITO_PER_ANAGRAFICA_ORDINE, Me.cdCliente.KeyFieldID)
        rs!Sconto1 = Me.txtSconto1.Value
        rs!Sconto2 = Me.txtSconto2.Value
        rs!ImportoUnitarioImballo = Me.txtImportoUnitarioImballo.Value ' GET_PREZZO_IMBALLO_2(Me.CDArticolo.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA, LINK_LISTINO_IMBALLI_AZIENDA, Me.txtIDOrdineCliente.Value, PREZZI_IMBALLI_DA_ORDINE, Me.txtIDPedana.Value, Me.CDCodiceImballo.KeyFieldID, LINK_SITO_PER_ANAGRAFICA_ORDINE, Me.cdCliente.KeyFieldID)
        rs!IDRV_POProcessoIVGamma = Me.txtIDProcessoIVGamma.Value
        rs!OraLavorazione = Me.txtOraLav.Text
        
        rs!NotaRigaOrdRaggr = Me.txtRaggrOrd.Text
        rs!RV_POImportoUnitarioListino = Me.txtImportoUnitarioListino.Value
        rs!IDValoriOggettoDettaglioRigaOrd = Me.txtIDRigaOrdine.Value
        
        rs("IDArticoloImballoPrimario").Value = Me.CDImballoPrimario.KeyFieldID
        rs("TaraConfezioneImballo").Value = Me.txtTaraConfImballo.Value
        rs("NumeroConfezioniPerImballo").Value = Me.txtNumeroConfImballo.Value
        rs("CostoConfezioneImballo").Value = txtCostoConfezione.Value
        
        
        Me.txtTaraConfImballo.Value = fnNotNullN(rs("TaraConfezioneImballo").Value)
        Me.txtNumeroConfImballo.Value = fnNotNullN(rs("NumeroConfezioniPerImballo").Value)
        
        rs!TracciaImballo = Me.chkTracciaImballoGest.Value
        rs!ConfermaDaUtente = Me.chkConfermaDaUtente.Value
        
        rs!TracciaImballoPrim = Me.chkTracciaImballoGestPrim.Value
        rs!ConfermaDaUtentePrim = Me.chkConfermaDaUtentePrim.Value
        rs!QuantitaPerCollo = QUANTITA_PER_COLLO
        rs!PesoPerCollo = PESO_LORDO_ARTICOLO
        rs!MoltiplicatorePerCollo = Moltiplicatore
        
        rs!IDRV_POProcessoIVGamma = Link_ProcessoIVGamma
        rs!IDOggettoOrdinePrec = Link_OggettoOrdinePrec
        rs!IDRV_POProcessoLavorazione = Link_ProcessoLavorazione
        rs!IDRV_POProcessoLavorazioneRighe = Link_ProcessoLavorazioneRighe
        rs!IDRV_POLineaProduzione = Link_LineaProduzione
        rs!IDRV_POCaricoMerceRighePrelievi = Link_CaricoMerceRighePrelievi
        rs!IDRV_POTipoUtilizzoLinea = Link_TipoUtilizzoLinea
        rs!IDRV_PO01_LottoCampagna = Link_LottoCampagnaSuLotto
        rs!PreConferimento = IsPreConferimento
                
    rs.Update

rs.Close
Set rs = Nothing

AGGIORNA_TIPO_PEDANA Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, 0

CREA_RECORDSET_KIT Me.CDArticolo.KeyFieldID, Me.CDCodiceImballo.KeyFieldID, Me.CDImballoPrimario.KeyFieldID, IDAssegnazione, LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO

SALVA_KIT IDAssegnazione

Exit Sub
ERR_fnInserisciNuovaAssegnazione:
    MsgBox Err.Description, vbCritical, "Inserimento nuova assegnazione"
    IDAssegnazione = 0
End Sub

Private Sub fnAggiornaAssegnazione(IDAssegnazione As Long)
On Error GoTo ERR_fnAggiornaAssegnazione
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim IDLavorazioneQuadratura As Long

If Me.chkScarti.Value = Unchecked Then
    If (((MAX_QUANTITA - Me.txtQta_UM.Value) <= 0) And ((MAX_QUANTITA_COLLI - Me.txtColli.Value) <= 0)) Then
    
        sSQL = "DELETE FROM RV_POAssegnazioneMerce "
        sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
        CnDMT.Execute sSQL
    
    
        sSQL = "DELETE FROM RV_POAssegnazioneMerceImbPrim "
        sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
        CnDMT.Execute sSQL
    
    
        Me.List1.AddItem "Movimentazione dell'assegnazione merce primaria (" & IDAssegnazione & ")"
        Me.List1.ListIndex = Me.List1.ListCount - 1
        DoEvents
        

        
        MOVIMENTAZIONE_RIGA_LAVORAZIONE Link_CaricoMerceRighe, IDAssegnazione, False
        
        
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
        rs.Fields("Colli").Value = rs.Fields("Colli").Value - Me.txtColli.Value
        rs.Fields("PesoLordo").Value = rs.Fields("PesoLordo").Value - Me.txtPesoLordo.Value
        rs.Fields("PesoNetto").Value = rs.Fields("PesoNetto").Value - Me.txtPesoNetto.Value
        rs.Fields("Tara").Value = rs.Fields("Tara").Value - Me.txtTara.Value
        If rs.Fields("Pezzi").Value = 0 Then
            rs.Fields("Pezzi").Value = 0
        Else
            rs.Fields("Pezzi").Value = rs.Fields("Pezzi").Value - Me.txtPezzi.Value
        End If
        rs.Fields("Qta_UM").Value = rs.Fields("Qta_UM").Value - Me.txtQta_UM.Value
    rs.Update
    
    
    CREA_RECORDSET_KIT fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!IDArticoloImballoPrimario), IDAssegnazione, IDAssegnazione

    SALVA_KIT IDAssegnazione
    
    rs.Close
    Set rs = Nothing
    
    Me.List1.AddItem "*******Movimentazione dell'assegnazione merce primaria (" & IDAssegnazione & ")********"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents
    
    MOVIMENTAZIONE_RIGA_LAVORAZIONE Link_CaricoMerceRighe, IDAssegnazione, True
    
Else
    IDLavorazioneQuadratura = 0
    If ((MAX_QUANTITA - Me.txtQta_UM.Value) <> 0) Then
        
        Set rs = New ADODB.Recordset
        sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
        sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
        
        rs.Open sSQL, CnDMT.InternalConnection
        
        If Not rs.EOF Then
            sSQL = "SELECT * FROM RV_POLavorazione "
            sSQL = sSQL & "WHERE IDRV_POLavorazione=0"
            
            Set rsNew = New ADODB.Recordset
            rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
            
            rsNew.AddNew
                rsNew.Fields("IDRV_POLavorazione").Value = fnGetNewKey("RV_POLavorazione", "IDRV_POLavorazione")
                rsNew.Fields("IDRV_POCaricoMerceRighe").Value = Link_CaricoMerceRighe
                rsNew.Fields("IDArticolo").Value = Me.CDArticoloScarto.KeyFieldID
                rsNew.Fields("CodiceArticolo").Value = Me.CDArticoloScarto.Code
                rsNew.Fields("Articolo").Value = Me.CDArticoloScarto.Description
                rsNew.Fields("IDUnitaDiMisuraCoop").Value = Link_UnitaDiMisura_Coop_Q
                rsNew.Fields("IDUnitaDiMisura").Value = Me.cboUMScarti.CurrentID
                
                rsNew.Fields("IDImballoVendita").Value = Me.CDCodiceImballo.KeyFieldID
                rsNew.Fields("CodiceImballoVendita").Value = Me.CDCodiceImballo.Code
                rsNew.Fields("ImballoVendita").Value = Me.txtImballo.Text
                rsNew.Fields("DataDocumento").Value = Date
                rsNew.Fields("OraLavorazione").Value = GET_ORARIO(Now)
                rsNew.Fields("IDTipoLavorazione").Value = Me.CboTipoLavScarti.CurrentID
                
                rsNew.Fields("Colli").Value = Abs(rs.Fields("Colli").Value - Me.txtColli.Value)
                rsNew.Fields("PesoLordo").Value = Abs(rs.Fields("PesoLordo").Value - Me.txtPesoLordo.Value)
                rsNew.Fields("PesoNetto").Value = Abs(rs.Fields("PesoNetto").Value - Me.txtPesoNetto.Value)
                rsNew.Fields("Tara").Value = Abs(rs.Fields("tara").Value - Me.txtTara.Value)
                If rs.Fields("Pezzi").Value = 0 Then
                    rsNew.Fields("Pezzi").Value = 0
                Else
                    rsNew.Fields("Pezzi").Value = Abs(rs.Fields("Pezzi").Value - Me.txtPezzi.Value)
                End If
                Select Case Link_UnitaDiMisura_Coop_Q
                    Case 1
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("Colli").Value
                    Case 2
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("PesoLordo").Value
                    Case 3
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("PesoNetto").Value
                    Case 4
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("Tara").Value
                    Case 5
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("Pezzi").Value
                    Case Else
                        rsNew.Fields("Qta_UM").Value = rsNew.Fields("PesoNetto").Value
                End Select
            rsNew.Update
            IDLavorazioneQuadratura = fnNotNullN(rsNew!IDRV_POLavorazione)
            rsNew.Close
            Set rsNew = Nothing
        End If
        
        rs.Close
        Set rs = Nothing
        
        If IDLavorazioneQuadratura > 0 Then
            MOVIMENTAZIONE_RIGA_QUADRATURA IDLavorazioneQuadratura
        End If
        
    End If
        
    sSQL = "DELETE FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
    CnDMT.Execute sSQL


    sSQL = "DELETE FROM RV_POAssegnazioneMerceImbPrim "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione
    CnDMT.Execute sSQL


    Me.List1.AddItem "Movimentazione dell'assegnazione merce primaria (" & IDAssegnazione & ")"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents
    
    MOVIMENTAZIONE_RIGA_LAVORAZIONE Link_CaricoMerceRighe, IDAssegnazione, False

End If
Exit Sub
ERR_fnAggiornaAssegnazione:
    MsgBox Err.Description, vbCritical, "Aggiorna assegnazione"
End Sub
Private Sub cmdNuovaPedana_Click()

If Me.txtIDPedana.Value > 0 Then
    If MsgBox("Questa riga ha già assegnata una pedana." & vbCrLf & "Vuoi crearla una nuova?", vbYesNo + vbQuestion, "Creazione nuova pedana") = vbNo Then Exit Sub
End If

Me.txtCodicePedana.Text = GetNumeroPedana(DatePart("yyyy", Date))
Me.txtIDPedana.Value = Link_Pedana
Me.cboTipoPedana.WriteOn GET_TIPO_PEDANA(Link_Pedana)

End Sub
Private Sub cmdSelezionaPedana_Click()
    WHERE_TROVA_PEDANA = 2
    frmTrovaPedana.Show vbModal
    Me.cboTipoPedana.WriteOn GET_TIPO_PEDANA(Me.txtIDPedana.Text)
End Sub
Private Sub Form_Activate()
    If bloading_Form = False Then bloading_Form = True
End Sub
Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    bloading_Form = False
    ParametroImballo
    ParametroPesoArticolo
    ParametroTipoArrotondamento
    ParametroTipoScarto
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroTipoSpezzatura

    InitControlli
    
    GET_DATI_ASSEGNAZIONE
    
    BLOCCA_PER_SEMAFORO Link_CaricoMerceRighe
    
    MAX_QUANTITA = Me.txtQta_UM.Value
    MAX_QUANTITA_COLLI = Me.txtColli.Value
    MAX_QUANTITA_CONFEZ = Me.txtColli.Value * Me.txtNumeroConfImballo.Value
    
    GET_DATI_CONFERIMENTO Link_CaricoMerceRighe
    
    Me.chkScarti.Value = vbUnchecked
    


    'If COMANDO_SPEZZATUTA = 1 Then
    '    Me.fraScarti.Enabled = False
    '    Me.chkPrezzoInclusoImballo.Value = vbUnchecked
    'Else
    '    If COMANDO_RIPESATURA = 1 Then
    '        Me.chkScarti.Value = vbChecked
    '        Me.fraScarti.Enabled = True
    '    Else
    '        Me.chkScarti.Value = vbUnchecked
    '        Me.fraScarti.Enabled = False
    '    End If
    'End If
    
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub InitControlli()
On Error GoTo ERR_InitControlli
     With Me.CDArticolo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
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

     With Me.CDArticoloScarto
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
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
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Imballo
    With Me.CDImballoPrimario
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
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    
 
     'Magazzino di conferimento
    With Me.cboMagazzinoConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .Sql = "SELECT * FROM Magazzino WHERE IDFiliale=" & TheApp.Branch
        .Fill
    End With
    
    'Magazzino di vendita
    With Me.CboMagazzinoVend
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .Sql = "SELECT * FROM Magazzino WHERE IDFiliale=" & TheApp.Branch
        .Fill
    End With
 
     With Me.CboTipoLavorazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .Sql = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With
   
     With Me.CboTipoLavScarti
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .Sql = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With
    'Unita di misura
    With Me.cboUM
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .Sql = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    
    'Unita di misura scarto
    With Me.cboUMScarti
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .Sql = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    'Unita di misura di imballo
    With Me.cboUMImballo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .Sql = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    
    'Cliente per ordine
     With Me.cdCliente
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIETipoAnagraficaCliente"
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
        'IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    With Me.cboCalibro
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POCalibro"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Calibro"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM RV_POCalibro"
        .Sql = .Sql & " ORDER BY Calibro"
    End With
  
    With Me.cboTipoCategoria
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POTipoCategoria"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoCategoria"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM RV_POTipoCategoria"
        .Sql = .Sql & " ORDER BY TipoCategoria"
    End With

   With Me.cboUtente
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDUtente"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Utente"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM Utente"
        .Sql = .Sql & " ORDER BY Utente"
    End With
    
    With Me.cboLinguaCliente
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDLinguaDescrizioneArticolo"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "LinguaDescrizioneArticolo"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM LinguaDescrizioneArticolo"
        .Sql = .Sql & " ORDER BY LinguaDescrizioneArticolo"
    End With

    With Me.cboLinguaPred
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDLinguaDescrizioneArticolo"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "LinguaDescrizioneArticolo"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM LinguaDescrizioneArticolo"
        .Sql = .Sql & " ORDER BY LinguaDescrizioneArticolo"
    End With

    With Me.cboTipoPedana
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POTipoPedana"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoPedana"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .Sql = "SELECT * FROM RV_POTipoPedana "
        .Sql = .Sql & "WHERE IDAzienda=" & TheApp.IDFirm
        .Sql = .Sql & " ORDER BY TipoPedana"
    End With


    GET_CONFIGURAZIONE_DOCUMENTO

Exit Sub
ERR_InitControlli:
    MsgBox Err.Description, vbCritical, "InitControlli"
End Sub
Private Sub ParametroTipoCaloPeso()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCaloPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
Else
    Link_TipoCaloPeso = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub ParametroTipoAumentoPeso()
On Error GoTo ERR_ParametroTipoAumentoPeso
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoAumentoPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
Else
    Link_TipoAumentoPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ParametroTipoAumentoPeso:
    MsgBox Err.Description, vbCritical, "ParametroTipoAumentoPeso"
End Sub
Private Sub ParametroTipoScarto()
On Error GoTo ERR_ParametroTipoScarto
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoScarto FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoScarto = fnNotNullN(rs!IDTipoScarto)
Else
    Link_TipoScarto = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ParametroTipoScarto:
    MsgBox Err.Description, vbCritical, "ParametroTipoScarto"
    
End Sub

Private Sub ParametroImballo()
On Error GoTo ERR_ParametroImballo
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoImballo = fnNotNullN(rs!IDTipoImballo)
Else
    Link_TipoImballo = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ParametroImballo:
    MsgBox Err.Description, vbCritical, "ParametroImballo"
End Sub
Private Sub ParametroTipoSpezzatura()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoSpaccaturaLavorazione FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSpezzatura = fnNotNullN(rs!IDRV_POTipoSpaccaturaLavorazione)
Else
    Link_TipoSpezzatura = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_DATI_ASSEGNAZIONE()
On Error GoTo ERR_GET_DATI_ASSEGNAZIONE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtIDPedana.Value = 0
    Me.txtCodicePedana.Text = ""
    Me.cboTipoPedana.WriteOn 0
    
    IDAssegnazionePrimaria = fnNotNullN(rs!IDRV_POAssegnazioneMerce)
    Link_CaricoMerceRighe = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Link_ProcessoIVGamma = fnNotNullN(rs!IDRV_POProcessoIVGamma)
    
    Link_TipoProdotto = GET_TIPO_PRODOTTO_ARTICOLO(fnNotNullN(rs("IDArticolo").Value))
    Me.CDArticolo.Load fnNotNullN(rs("IDArticolo").Value)
    Me.CDArticolo.Code = fnNotNull(rs("CodiceArticolo").Value)
    Me.TxtArticolo.Text = fnNotNull(rs("Articolo").Value)
    Me.cboUM.WriteOn fnNotNullN(rs!IDUnitaDiMisura)
    Link_UnitaDiMisura_Coop = fnNotNullN(rs("IDUnitaDiMisuraCoop").Value)

    Me.CDCodiceImballo.Load fnNotNullN(rs("IDImballoVendita").Value)
    Me.CDCodiceImballo.Code = fnNotNull(rs("CodiceImballoVendita").Value)
    Me.txtTaraUnitaria.Value = fnNotNullN(rs("TaraUnitaria").Value)
    
    Me.CDImballoPrimario.Load fnNotNullN(rs("IDArticoloImballoPrimario").Value)
    Me.txtTaraConfImballo.Value = fnNotNullN(rs("TaraConfezioneImballo").Value)
    Me.txtNumeroConfImballo.Value = fnNotNullN(rs("NumeroConfezioniPerImballo").Value)
    Me.txtCostoConfezione.Value = fnNotNullN(rs("CostoConfezioneImballo").Value)
    
    Me.txtColli.Value = fnNotNullN(rs("Colli").Value)
    Me.txtPesoLordo.Value = fnNotNullN(rs("PesoLordo").Value)
    Me.txtPesoNetto.Value = fnNotNullN(rs("PesoNetto").Value)
    
    Me.txtTara.Value = fnNotNullN(rs("Tara").Value)
    Me.txtPezzi.Value = fnNotNullN(rs("Pezzi").Value)
    Me.txtQta_UM.Value = fnNotNullN(rs("Qta_UM").Value)
    
    
    'Me.txtImballo.Text = fnNotNull(rs("ImballoVendita").Value)
    Me.txtDataLavorazione.Value = fnNotNull(rs("DataDocumento").Value)
    Me.txtOraLav.Text = fnNotNull(rs("OraLavorazione").Value)
    Me.CboTipoLavorazione.WriteOn fnNotNullN(rs("IDTipoLavorazione").Value)
    Me.cboTipoCategoria.WriteOn fnNotNullN(rs("IDRV_POTipoCategoria").Value)
    Me.cboCalibro.WriteOn fnNotNullN(rs("IDRV_POCalibro").Value)
    Me.cboLinguaPred.WriteOn fnNotNullN(rs("IDLinguaPredefinita").Value)
    
    Me.cdCliente.Load LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO
    Me.txtIDOrdineCliente.Value = LINK_ORDINE_MERCE_PER_SMISTAMENTO
    Me.txtNumeroOrdine.Value = NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO
    Me.txtDataOrdine.Text = DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO
    Me.txtIDOrdinePadre.Value = LINK_ORDINE_PADRE_MERCE_PER_SMISTAMENTO
    Me.txtNListaPrelievo.Value = NUMERO_LISTA_ORDINE_MERCE_PER_SMISTAMENTO
    
    Me.cboLinguaCliente.WriteOn fnNotNullN(rs("IDLinguaCliente").Value)
    Me.txtCodicePedanaCliente.Value = fnNotNullN(rs("NumeroPedanaCliente").Value)
    
    Me.txtLottoVendita.Text = fnNotNull(rs("CodiceLottoVendita").Value)
    Me.txtLottoCliente.Text = fnNotNull(rs("LottoCliente").Value)
    Me.txtAltreAnnotazioniCliente.Text = fnNotNull(rs("AltreAnnotazioniPerCliente").Value)
    
    Me.cboUtente.WriteOn fnNotNullN(rs("IDUtente").Value)
    Me.txtCodicePostazione.Text = fnNotNull(rs("CodiceUtente").Value)
    
    Me.txtCodicePedanaCliente.Value = fnNotNullN(rs("NumeroPedanaCliente").Value)
    Me.chkPrezzoInclusoImballo.Value = Abs(fnNotNullN(rs("MerceInclusoImballo").Value))
    Me.txtNomeMacchina.Text = GET_NOMECOMPUTER
    Me.txtUtenteMacchina.Text = GET_NOMEUTENTE
    
    Me.txtCodiceGSICliente.Text = fnNotNull(rs("CodiceGSI").Value)
    Me.txtCodiceAssociatoPressoCliente.Text = fnNotNull(rs("CodiceAssociatoPressoCliente").Value)
    Me.txtCodificaArticoloCodiceABarre.Text = fnNotNull(rs("CodiceABarreArticoloCliente").Value)
    
    Me.txtCodiceABarreArticoloCliente.Text = fnNotNull(rs("DescrizioneCodiceABarreArticoloCliente").Value)
    Me.txtCodificaImballoCodiceABarre.Text = fnNotNull(rs("CodiceABarreImballoCliente").Value)
    Me.txtCodiceABarreImballoCliente.Text = fnNotNull(rs("DescrizioneCodiceABarreImballoCliente").Value)
    Me.txtArticoloInLinguaPred.Text = fnNotNull(rs("DescrizioneArticoloInLinguaPred").Value)
    Me.txtCalibroInLiguaPred.Text = fnNotNull(rs("DescrizioneCalibroInLinguaPred").Value)
    Me.txtCategoriaInLinguaPred.Text = fnNotNull(rs("DescrizioneCategoriaInLinguaPred").Value)
    Me.txtDescrizioneArticoloInLinguaCliente.Text = fnNotNull(rs("DescrizioneArticoloInLinguaCliente").Value)
    Me.txtCalibroInLingua.Text = fnNotNull(rs("DescrizioneCalibroInLinguaCliente").Value)
    Me.txtCategoriaInLingua.Text = fnNotNull(rs("DescrizioneCategoriaInLinguaCliente").Value)
    Me.txtIDProcessoIVGamma.Value = fnNotNullN(rs!IDRV_POProcessoIVGamma)
    Me.txtRaggrOrd.Text = fnNotNull(rs!NotaRigaOrdRaggr)
    
    GET_DATI_CONFERIMENTO Link_CaricoMerceRighe
    
    LINK_PEDANA_LAVORAZIONE = fnNotNullN(rs!IDRV_POPedana)
    CODICE_PEDANA_LAVORAZIONE = fnNotNull(rs!CodicePedana)
    
    Me.chkTracciaImballoGest.Value = Abs(fnNotNullN(rs("TracciaImballo").Value))
    Me.chkConfermaDaUtente.Value = Abs(fnNotNullN(rs("ConfermaDaUtente").Value))
    
    Me.chkTracciaImballoGestPrim.Value = Abs(fnNotNullN(rs("TracciaImballoPrim").Value))
    Me.chkConfermaDaUtentePrim.Value = Abs(fnNotNullN(rs("ConfermaDaUtentePrim").Value))
    
    QUANTITA_PER_COLLO = fnNotNullN(rs!QuantitaPerCollo)
    PESO_LORDO_ARTICOLO = fnNotNullN(rs!PesoPerCollo)
    Moltiplicatore = fnNotNullN(rs!MoltiplicatorePerCollo)
    
    Me.txtQuantitaPerCollo.Value = QUANTITA_PER_COLLO
    Me.txtPesoPerCollo.Value = PESO_LORDO_ARTICOLO
    
    Link_OggettoOrdinePrec = fnNotNullN(rs!IDOggettoOrdinePrec)
    Link_ProcessoLavorazione = fnNotNullN(rs!IDRV_POProcessoLavorazione)
    Link_ProcessoLavorazioneRighe = fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe)
    Link_LineaProduzione = fnNotNullN(rs!IDRV_POLineaProduzione)
    Link_CaricoMerceRighePrelievi = fnNotNullN(rs!IDRV_POCaricoMerceRighePrelievi)
    Link_TipoUtilizzoLinea = fnNotNullN(rs!IDRV_POTipoUtilizzoLinea)
    Link_LottoCampagnaSuLotto = fnNotNullN(rs!IDRV_PO01_LottoCampagna)
    IsPreConferimento = fnNotNullN(rs!PreConferimento)
    
    CREA_RECORDSET_LOTTI_IMBALLI fnGetTipoOggetto("RV_POAssegnazioneMerce"), IDAssegnazionePrimaria
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_DATI_ASSEGNAZIONE:
    MsgBox Err.Description, vbCritical, "GET_DATI_ASSEGNAZIONE"
End Sub
Private Function GET_OGGETTO_ORDINE() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
    sSQL = sSQL & "WHERE ValoriOggettoPerTipo000F.Link_nom_anagrafica=" & Me.cdCliente.KeyFieldID
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_numero=" & Me.txtNumeroOrdine.Value
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_data=" & fnNormDate(Me.txtDataOrdine.Text)
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_OGGETTO_ORDINE = 0
    Else
        GET_OGGETTO_ORDINE = fnNotNullN(rs!IDOggetto)
    End If
rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    SBLOCCA_PER_SEMAFORO Link_CaricoMerceRighe
End Sub

Private Sub FraPedana_DblClick()
    If Me.FraPedana.Height = 855 Then
        Me.FraPedana.Height = 3615
    Else
        Me.FraPedana.Height = 855
    End If
End Sub

Private Sub txtCodicePedanaCliente_Change()
    Me.txtCodificaCodicePedanaCliente.Text = GET_CODICE_PEDANA_CLIENTE(Me.txtCodicePedanaCliente.Value)
End Sub
Private Function GET_CODICE_PEDANA_CLIENTE(NumeroPedanaCliente As Long) As String
Dim I As Integer
GET_CODICE_PEDANA_CLIENTE = ""
For I = Len(CStr(NumeroPedanaCliente)) To MAX_CARATTERI_PEDANA_CLIENTE
    GET_CODICE_PEDANA_CLIENTE = GET_CODICE_PEDANA_CLIENTE & "0"
Next

GET_CODICE_PEDANA_CLIENTE = GET_CODICE_PEDANA_CLIENTE & CStr(NumeroPedanaCliente)

End Function

Private Sub txtColli_Change()
    If Link_UnitaDiMisura_Coop = 1 Then
        Me.txtQta_UM.Value = Me.txtColli.Value
    End If
End Sub
Private Sub txtColli_LostFocus()
On Error Resume Next
    
    Me.txtTara.Value = (Me.txtColli.Value * Me.txtTaraUnitaria.Value) + GET_CALCOLA_TARA_CONFEZIONE
    If PESO_LORDO_ARTICOLO > 0 Then
        If TIPO_PESO_ARTICOLO <= 1 Then
            Me.txtPesoLordo.Value = PESO_LORDO_ARTICOLO * Me.txtColli.Value
        Else
            Me.txtPesoNetto.Value = PESO_LORDO_ARTICOLO * Me.txtColli.Value
            Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
        End If
    Else
        Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
    End If
    If QUANTITA_PER_COLLO > 1 Then
        Me.txtPezzi.Value = Me.txtColli.Value * QUANTITA_PER_COLLO
    End If
    
    CalcoloPesoNetto
End Sub
Private Sub CalcoloPesoNetto()
On Error Resume Next
Dim ArrayPesoNetto() As String
Dim PesoNetto As Double
Dim Decimal_PesoNetto As Double
Dim Tara As Double
If bloading_Form = False Then Exit Sub

'Me.txtTara.Value = Me.txtTaraUnitaria.Value * Me.txtColli.Value

Me.txtTara.Value = (Me.txtColli.Value * Me.txtTaraUnitaria.Value) + GET_CALCOLA_TARA_CONFEZIONE


If TIPO_PESO_ARTICOLO <= 1 Then
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
Else
    Me.txtPesoLordo.Value = Me.txtTara.Value + Me.txtPesoNetto.Value
End If

Select Case Link_Arrontondamento
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

    
    
'Me.txtTaraUnitaria.Value = Me.txtTara.Value / Me.txtColli.Value



End Sub

Private Sub txtIDOrdineCliente_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If Me.txtIDOrdineCliente.Value = 0 Then
    Me.lblStatoOrdine.Caption = "ORDINE NON ASSEGNATO"
    Me.lblStatoOrdine.BackColor = vbRed
    Exit Sub
End If
    
sSQL = "SELECT Link_Vet_Vettore, Link_nom_ult_sito, Doc_ordine_chiuso "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDOrdineCliente.Value

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.lblStatoOrdine.Caption = "ORDINE NON TROVATO"
    LINK_SITO_PER_ANAGRAFICA_ORDINE = 0
Else
    If fnNotNullN(rs!Doc_ordine_chiuso) = False Then
        Me.lblStatoOrdine.Caption = "ORDINE APERTO"
        Me.lblStatoOrdine.BackColor = vbGreen
    Else
        Me.lblStatoOrdine.Caption = "ORDINE CHIUSO"
        Me.lblStatoOrdine.BackColor = vbRed
    End If
    
    LINK_SITO_PER_ANAGRAFICA_ORDINE = fnNotNullN(rs!Link_nom_ult_sito)
    
End If

rs.CloseResultset
Set rs = Nothing


LINK_LISTINO_CLIENTE_ORDINE = GET_LINK_LISTINO_CLIENTE(Me.cdCliente.KeyFieldID, LINK_SITO_PER_ANAGRAFICA_ORDINE)


GET_INTESTAZIONE_DOCUMENTO Me.cdCliente.KeyFieldID, LINK_LISTINO_CLIENTE_ORDINE, LINK_LISTINO_AZIENDA


End Sub

Private Sub txtPesoLordo_Change()
    If Link_UnitaDiMisura_Coop = 2 Then
        Me.txtQta_UM.Value = Me.txtPesoLordo.Value
    End If
End Sub

Private Sub txtPesoLordo_LostFocus()
    CalcoloPesoNetto
End Sub

Private Sub txtPesoNetto_Change()
    If Link_UnitaDiMisura_Coop = 3 Then
        Me.txtQta_UM.Value = Me.txtPesoNetto.Value
    End If
    
End Sub

Private Sub txtPesoNetto_LostFocus()
    CalcoloPesoNetto
End Sub

Private Sub txtPezzi_Change()
    If Link_UnitaDiMisura_Coop = 5 Then
        Me.txtQta_UM.Value = Me.txtPezzi.Value
    End If
    
End Sub

Private Sub txtTara_Change()
    If Link_UnitaDiMisura_Coop = 4 Then
        Me.txtQta_UM.Value = Me.txtTara.Value
    End If
    
End Sub
Private Function GetNumeroPedana(Anno As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim CodicePedana As String
Dim NumeroPedana As Long

sSQL = "SELECT MAX(CodiceID) AS NumeroPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE Anno=" & Anno
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF Then
        CodicePedana = Anno & GET_CODICE_PEDANA(1)
        NumeroPedana = 1
    Else
        NumeroPedana = (fnNotNullN(rs!NumeroPedana) + 1)
        CodicePedana = Anno & GET_CODICE_PEDANA((fnNotNullN(rs!NumeroPedana) + 1))
        
    End If
rs.CloseResultset
Set rs = Nothing

Link_Pedana = fnGetNewKey("RV_POPedana", "IDRV_POPedana")

sSQL = "INSERT INTO RV_POPedana ("
sSQL = sSQL & "IDRV_POPedana, CodiceID, IDAzienda, IDFiliale, Anno, Mese, Giorno, IDRV_POTipoPedana, Codice) "
sSQL = sSQL & "VALUES ("
sSQL = sSQL & Link_Pedana & ", "
sSQL = sSQL & NumeroPedana & ","
sSQL = sSQL & TheApp.IDFirm & ", "
sSQL = sSQL & TheApp.Branch & ", "
sSQL = sSQL & DatePart("yyyy", Date) & ", "
sSQL = sSQL & Month(Date) & ", "
sSQL = sSQL & Day(Date) & ", "
sSQL = sSQL & GET_TIPO_PEDANA_DEFAULT & ", "
sSQL = sSQL & fnNormString(CodicePedana) & " "
sSQL = sSQL & ")"

CnDMT.Execute sSQL


GetNumeroPedana = CodicePedana

Exit Function
ERR_GetNumeroPedana:
    MsgBox Err.Description, vbCritical, "Nuova pedana"
    Link_Pedana = 0
    Me.txtIDPedana.Value = 0
    GetNumeroPedana = ""
End Function
Private Function GET_CODICE_PEDANA(NumeroPedana As String) As String
Dim I As Integer
Const MAX_CAR As Integer = 7
GET_CODICE_PEDANA = ""
For I = Len(NumeroPedana) + 1 To MAX_CAR
GET_CODICE_PEDANA = GET_CODICE_PEDANA & "0"
    
Next
GET_CODICE_PEDANA = GET_CODICE_PEDANA & NumeroPedana
End Function
Private Function GET_TIPO_PEDANA_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoPedanaDefault "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PEDANA_DEFAULT = 0
Else
    GET_TIPO_PEDANA_DEFAULT = fnNotNullN(rs!IDTipoPedanaDefault)
    
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function GET_ESISTENZA_ORDINE() As Long
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ValoriOggettoPerTipo000F.IDOggetto "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND "
sSQL = sSQL & "ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo000F.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica "
sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    If Me.cdCliente.KeyFieldID > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Link_nom_anagrafica=" & Me.cdCliente.KeyFieldID
    End If
    
    If Me.txtNumeroOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_numero=" & Me.txtNumeroOrdine.Value
    End If
    
    If Me.txtDataOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_data=" & fnNormDate(Me.txtDataOrdine.Value)
    End If
    
    sSQL = sSQL & sSQL_WHERE & " ORDER BY Doc_data DESC, Doc_numero DESC"

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_ESISTENZA_ORDINE = 0
Else
    GET_ESISTENZA_ORDINE = fnNotNullN(rs!IDOggetto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function GET_ESISTENZA_PEDANA() As Long
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE Codice =" & fnNormString(Me.txtCodicePedana.Text)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_ESISTENZA_PEDANA = 0
Else
    GET_ESISTENZA_PEDANA = fnNotNullN(rs!IDRV_POPedana)
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
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    
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
Private Sub GET_DATI_CONFERIMENTO(IDRigaConferimento As Long)
On Error GoTo ERR_GET_DATI_CONFERIMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POCaricoMerceTesta.IDAnagrafica, RV_POCaricoMerceTesta.Anagrafica, RV_POCaricoMerceTesta.Nome, "
sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento , RV_POCaricoMerceTesta.DataDocumento, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDMagazzinoConferimento, RV_POCaricoMerceTesta.IDMagazzinoVendita, "
sSQL = sSQL & "RV_POCaricoMerceRighe.LottoDiConferimento, RV_POCaricoMerceRighe.CodiceLotto, RV_POCaricoMerceRighe.IDUnitaDiMisuraDiamante, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceRighe.IDUnitaDiMisura "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_ANAGRAFICA_SOCIO = 0
    CODICE_SOCIO = ""
    ANAGRAFICA_SOCIO = ""
    NOME_SOCIO = ""
    DATA_CONFERIMENTO = ""
    NUMERO_CONFERIMENTO = ""
    CODICE_LOTTO_CAMPAGNA = ""
    CODICE_LOTTO_ENTRATA = ""
    LINK_ARTICOLO_CONFERITO = 0
    ARTICOLO_CONFERITO = ""
    CODICE_ARTICOLO_CONFERITO = ""
    LINK_UNITA_DI_MISURA_ARTICOLO_CONFERITO = 0
    Me.cboMagazzinoConf.WriteOn 0
    Me.CboMagazzinoVend.WriteOn 0
    Link_UnitaDiMisura_Coop_Conferimento = 0
Else
    LINK_ANAGRAFICA_SOCIO = fnNotNullN(rs!IDAnagrafica)
    CODICE_SOCIO = GET_CODICE_SOCIO(rs!IDAnagrafica)
    ANAGRAFICA_SOCIO = fnNotNull(rs!Anagrafica)
    NOME_SOCIO = fnNotNull(rs!Nome)
    DATA_CONFERIMENTO = fnNotNull(rs!DataDocumento)
    NUMERO_CONFERIMENTO = fnNotNull(rs!NumeroDocumento)
    CODICE_LOTTO_CAMPAGNA = fnNotNull(rs!LottoDiConferimento)
    CODICE_LOTTO_ENTRATA = fnNotNull(rs!CodiceLotto)
    LINK_ARTICOLO_CONFERITO = fnNotNullN(rs!IDArticolo)
    ARTICOLO_CONFERITO = fnNotNull(rs!Articolo)
    CODICE_ARTICOLO_CONFERITO = fnNotNull(rs!CodiceArticolo)
    LINK_UNITA_DI_MISURA_ARTICOLO_CONFERITO = fnNotNullN(rs!IDUnitaDiMisuraDiamante)
    Me.cboMagazzinoConf.WriteOn fnNotNullN(rs!IDMagazzinoConferimento)
    Me.CboMagazzinoVend.WriteOn fnNotNullN(rs!IDMagazzinoVendita)
    Link_UnitaDiMisura_Coop_Conferimento = fnNotNullN(rs!IDUnitaDiMisura)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_DATI_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "GET_DATI_CONFERIMENTO"
End Sub
Private Function GET_CODICE_SOCIO(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT Codice FROM Fornitore "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_SOCIO = ""
Else
    GET_CODICE_SOCIO = fnNotNull(rs!Codice)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TIPO_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POPedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PEDANA = 0
Else
    GET_TIPO_PEDANA = fnNotNullN(rs!IDRV_POTipoPedana)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TIPO_PRODOTTO_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM Articolo WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PRODOTTO_ARTICOLO = 0
Else
    GET_TIPO_PRODOTTO_ARTICOLO = fnNotNullN(rs!IDTipoProdotto)
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
Private Sub GET_INFO_CLIENTE(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoInclusoImballo,CodiceGSI, MaxCaratteriPedana, CodiceAssociato "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtCodiceGSICliente.Text = ""
    Me.chkPrezzoInclusoImballo.Value = vbUnchecked
    MAX_CARATTERI_PEDANA_CLIENTE = 0
    Me.txtCodiceAssociatoPressoCliente.Text = ""
Else
    Me.txtCodiceGSICliente.Text = fnNotNull(rs!CodiceGSI)
    Me.chkPrezzoInclusoImballo.Value = Abs(fnNotNullN(rs!PrezzoInclusoImballo))
    MAX_CARATTERI_PEDANA_CLIENTE = fnNotNullN(rs!MaxCaratteriPedana)
    Me.txtCodiceAssociatoPressoCliente.Text = fnNotNull(rs!CodiceAssociato)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_ARTICOLO_CODICE_A_BARRE_CLIENTE(IDAnagrafica As Long, IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceABarre, DescrizioneCodiceABarre "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteEAN13 "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtCodiceABarreArticoloCliente.Text = ""
    Me.txtCodificaArticoloCodiceABarre.Text = ""
Else
    Me.txtCodiceABarreArticoloCliente.Text = fnNotNull(rs!DescrizioneCodiceABarre)
    Me.txtCodificaArticoloCodiceABarre.Text = fnNotNull(rs!CodiceABarre)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_IMBALLO_CODICE_A_BARRE_CLIENTE(IDAnagrafica As Long, IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceABarre, DescrizioneCodiceABarre "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteEAN13 "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtCodiceABarreImballoCliente.Text = ""
    Me.txtCodificaImballoCodiceABarre.Text = ""
Else
    Me.txtCodiceABarreImballoCliente.Text = fnNotNullN(rs!DescrizioneCodiceABarre)
    Me.txtCodificaImballoCodiceABarre.Text = fnNotNullN(rs!CodiceABarre)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_DESCRIZIONE_IN_LINGUA(IDLingua As Long, IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ArticoloPerLinguaDescrizione "
sSQL = sSQL & "FROM ArticoloPerLinguaDescrizione "
sSQL = sSQL & "WHERE IDLinguaDescrizioneArticolo=" & IDLingua
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_IN_LINGUA = fnNotNull(rs!ArticoloPerLinguaDescrizione)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_CALIBRO_IN_LINGUA(IDLingua As Long, IDCalibro As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CalibroLingua "
sSQL = sSQL & "FROM RV_POCalibroLingua "
sSQL = sSQL & "WHERE IDLingua=" & IDLingua
sSQL = sSQL & " AND IDRV_POCalibro=" & IDCalibro

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_CALIBRO_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_CALIBRO_IN_LINGUA = fnNotNull(rs!CalibroLingua)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_CATEGORIA_IN_LINGUA(IDLingua As Long, IDCategoria As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CategoriaLingua "
sSQL = sSQL & "FROM RV_POCategoriaLingua "
sSQL = sSQL & "WHERE IDLingua=" & IDLingua
sSQL = sSQL & " AND IDRV_POTipoCategoria=" & IDCategoria

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_CATEGORIA_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_CATEGORIA_IN_LINGUA = fnNotNull(rs!CategoriaLingua)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICEABARRE(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceBarre FROM ArticoloPerTipoCodiceBarre "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDTipoCodiceBarre=13"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICEABARRE = ""
Else
    GET_CODICEABARRE = ean13$((fnNotNull(Trim(rs!CodiceBarre))))
    
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_DESCRIZIONECODICEABARRE(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceBarre FROM ArticoloPerTipoCodiceBarre "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDTipoCodiceBarre=13"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONECODICEABARRE = ""
Else
    GET_DESCRIZIONECODICEABARRE = fnNotNull(Trim(rs!CodiceBarre))
    
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_LINK_ORDINAMENTO(tabella As String, NomeCampoWhere As String, valoreCampoWhere As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(Link_Ordinamento) AS Link_Ordinamento "
sSQL = sSQL & "FROM " & tabella
sSQL = sSQL & " WHERE " & NomeCampoWhere & "=" & valoreCampoWhere

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ORDINAMENTO = 1
Else
    GET_LINK_ORDINAMENTO = fnNotNullN(rs!LINK_ORDINAMENTO) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub MOVIMENTAZIONE_RIGA_LAVORAZIONE(IDRigaConferimento As Long, IDAssegnazioneMerce As Long, CreaMovimento As Boolean)
On Error GoTo ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE
Dim OLD_Cursor As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsMov As DmtOleDbLib.adoResultset
Dim Movimentato As Long

OLD_Cursor = CnDMT.CursorLocation
CnDMT.CursorLocation = adUseClient
Movimentato = 1
Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection

'''''''''''''''''''''''ELIMINAZIONE MOVIMENTI DELLA RIGA DI LAVORAZIONE
Me.List1.AddItem "Eliminazione movimenti"
Me.List1.ListIndex = Me.List1.ListCount - 1
DoEvents
sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
'sSQL = sSQL & " AND IDOggetto=" & IDRigaConferimento
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    CnDMT.BeginTrans
        mov.Delete fnNotNullN(rs!IDMovimento)
    CnDMT.CommitTrans
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If CreaMovimento = False Then Exit Sub

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    
    Select Case Link_UnitaDiMisura_Coop_Conferimento
        Case 1
            Me.List1.AddItem "Generazione movimenti di carico merce nel magazzino " & Me.CboMagazzinoVend.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiCarico(IDAssegnazioneMerce, Link_CaricoMerceRighe, Link_ProcessoIVGamma, fnNotNull(rs!CodiceLottoVendita), CODICE_LOTTO_CAMPAGNA, Me.CDArticolo.KeyFieldID, Me.cboUM.CurrentID, _
                Me.TxtArticolo.Text, fnNotNullN(rs!Qta_UM), Me.txtDataLavorazione.Text, fnNotNullN(rs!Colli), _
                fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
                fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "IDRV_POTipoLavorazione"), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "PrezzoMedio"), _
                Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, Me.txtCodicePedana.Text, GET_PESO_PEDANA(Me.txtIDPedana.Value), Me.CDImballoPrimario.KeyFieldID, Me.CDImballoPrimario.Code, Me.CDImballoPrimario.Description, _
                Me.txtNumeroConfImballo.Value, Me.txtTaraConfImballo.Value, 0) = False Then
                
                If (Movimentato = 1) Then Movimentato = 0
            End If

            Me.List1.AddItem "Generazione movimenti di scarico merce conferita nel magazzino " & Me.cboMagazzinoConf.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiScarico(IDAssegnazioneMerce, fnNotNullN(rs!Colli), Me.txtDataLavorazione.Text, IDRigaConferimento) = False Then
                If (Movimentato = 1) Then Movimentato = 0
            End If
        Case 2
            Me.List1.AddItem "Generazione movimenti di carico merce nel magazzino " & Me.CboMagazzinoVend.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiCarico(IDAssegnazioneMerce, Link_CaricoMerceRighe, Link_ProcessoIVGamma, fnNotNull(rs!CodiceLottoVendita), CODICE_LOTTO_CAMPAGNA, Me.CDArticolo.KeyFieldID, Me.cboUM.CurrentID, _
                                        Me.TxtArticolo.Text, fnNotNullN(rs!Qta_UM), Me.txtDataLavorazione.Text, fnNotNullN(rs!PesoLordo), _
                                        fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
                                        fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "IDRV_POTipoLavorazione"), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "PrezzoMedio"), _
                                        Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, Me.txtCodicePedana.Text, GET_PESO_PEDANA(Me.txtIDPedana.Value), Me.CDImballoPrimario.KeyFieldID, Me.CDImballoPrimario.Code, Me.CDImballoPrimario.Description, _
                Me.txtNumeroConfImballo.Value, Me.txtTaraConfImballo.Value, 0) = False Then
            
                If (Movimentato = 1) Then Movimentato = 0
                
            End If

            Me.List1.AddItem "Generazione movimenti di scarico merce conferita nel magazzino " & Me.cboMagazzinoConf.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiScarico(IDAssegnazioneMerce, fnNotNullN(rs!PesoLordo), Me.txtDataLavorazione.Text, IDRigaConferimento) = False Then
                If (Movimentato = 1) Then Movimentato = 0
            End If
        Case 3
            Me.List1.AddItem "Generazione movimenti di carico merce nel magazzino " & Me.CboMagazzinoVend.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiCarico(IDAssegnazioneMerce, Link_CaricoMerceRighe, Link_ProcessoIVGamma, fnNotNull(rs!CodiceLottoVendita), CODICE_LOTTO_CAMPAGNA, Me.CDArticolo.KeyFieldID, Me.cboUM.CurrentID, _
            Me.TxtArticolo.Text, fnNotNullN(rs!Qta_UM), Me.txtDataLavorazione.Text, fnNotNullN(rs!PesoNetto), _
            fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
            fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "IDRV_POTipoLavorazione"), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "PrezzoMedio"), _
            Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, Me.txtCodicePedana.Text, GET_PESO_PEDANA(Me.txtIDPedana.Value), Me.CDImballoPrimario.KeyFieldID, Me.CDImballoPrimario.Code, Me.CDImballoPrimario.Description, _
            Me.txtNumeroConfImballo.Value, Me.txtTaraConfImballo.Value, 0) = False Then Movimentato = 0
            
            Me.List1.AddItem "Generazione movimenti di scarico merce conferita nel magazzino " & Me.cboMagazzinoConf.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            
            If GeneraMovimentoDiScarico(IDAssegnazioneMerce, fnNotNullN(rs!PesoNetto), Me.txtDataLavorazione.Text, IDRigaConferimento) = False Then
                If (Movimentato = 1) Then Movimentato = 0
            End If
        Case 4
            Me.List1.AddItem "Generazione movimenti di carico merce nel magazzino " & Me.CboMagazzinoVend.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiCarico(IDAssegnazioneMerce, Link_CaricoMerceRighe, Link_ProcessoIVGamma, fnNotNull(rs!CodiceLottoVendita), CODICE_LOTTO_CAMPAGNA, Me.CDArticolo.KeyFieldID, Me.cboUM.CurrentID, _
            Me.TxtArticolo.Text, fnNotNullN(rs!Qta_UM), Me.txtDataLavorazione.Text, fnNotNullN(rs!Tara), _
            fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
            fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "IDRV_POTipoLavorazione"), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "PrezzoMedio"), _
            Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, Me.txtCodicePedana.Text, GET_PESO_PEDANA(Me.txtIDPedana.Value), Me.CDImballoPrimario.KeyFieldID, Me.CDImballoPrimario.Code, Me.CDImballoPrimario.Description, _
                Me.txtNumeroConfImballo.Value, Me.txtTaraConfImballo.Value, 0) = False Then Movimentato = 0
            
            
            Me.List1.AddItem "Generazione movimenti di scarico merce conferita nel magazzino " & Me.cboMagazzinoConf.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            If GeneraMovimentoDiScarico(IDAssegnazioneMerce, fnNotNullN(rs!Tara), Me.txtDataLavorazione.Text, IDRigaConferimento) = False Then
                If (Movimentato = 1) Then Movimentato = 0
            End If
        Case 5

            Me.List1.AddItem "Generazione movimenti di carico merce nel magazzino " & Me.CboMagazzinoVend.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents

            If GeneraMovimentoDiCarico(IDAssegnazioneMerce, Link_CaricoMerceRighe, Link_ProcessoIVGamma, fnNotNull(rs!CodiceLottoVendita), CODICE_LOTTO_CAMPAGNA, Me.CDArticolo.KeyFieldID, Me.cboUM.CurrentID, _
            Me.TxtArticolo.Text, fnNotNullN(rs!Qta_UM), Me.txtDataLavorazione.Text, fnNotNullN(rs!Pezzi), _
            fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
            fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "IDRV_POTipoLavorazione"), GET_DATI_RIGA_CONFERIMENTO(Link_CaricoMerceRighe, "PrezzoMedio"), _
            Me.txtIDPedana.Value, Me.cboTipoPedana.CurrentID, Me.txtCodicePedana.Text, GET_PESO_PEDANA(Me.txtIDPedana.Value), Me.CDImballoPrimario.KeyFieldID, Me.CDImballoPrimario.Code, Me.CDImballoPrimario.Description, _
            Me.txtNumeroConfImballo.Value, Me.txtTaraConfImballo.Value, 0) = False Then Movimentato = 0

            Me.List1.AddItem "Generazione movimenti di scarico merce conferita nel magazzino " & Me.cboMagazzinoConf.Text
            Me.List1.ListIndex = Me.List1.ListCount - 1
            DoEvents
            
            If GeneraMovimentoDiScarico(IDAssegnazioneMerce, fnNotNullN(rs!Pezzi), Me.txtDataLavorazione.Text, IDRigaConferimento) = False Then
                If (Movimentato = 1) Then Movimentato = 0
            End If
    End Select
    


    
    If Me.CDCodiceImballo.KeyFieldID > 0 Then
        Me.List1.AddItem "Generazione movimenti di carico imballo della merce nel magazzino " & Me.CboMagazzinoVend.Text
        Me.List1.ListIndex = Me.List1.ListCount - 1
        DoEvents
        If GeneraMovimentoCaricoImballo(Link_CaricoMerceRighe, IDAssegnazioneMerce, Link_ProcessoIVGamma, CODICE_LOTTO_ENTRATA, CODICE_LOTTO_CAMPAGNA, Me.CDCodiceImballo.KeyFieldID, Me.cboUMImballo.CurrentID, Me.txtImballo.Text, fnNotNullN(rs!Colli)) = False Then
            If (Movimentato = 1) Then Movimentato = 0
        End If
        
        Me.List1.AddItem "Generazione movimenti di scarico imballo della merce nel magazzino " & Me.cboMagazzinoConf.Text
        Me.List1.ListIndex = Me.List1.ListCount - 1
        DoEvents
        
        If GeneraMovimentoScaricoImballo(Link_CaricoMerceRighe, IDAssegnazioneMerce, Link_ProcessoIVGamma, CODICE_LOTTO_ENTRATA, CODICE_LOTTO_CAMPAGNA, Me.CDCodiceImballo.KeyFieldID, Me.cboUMImballo.CurrentID, Me.txtImballo.Text, fnNotNullN(rs!Colli), 0, MAX_QUANTITA_COLLI) = False Then
           If (Movimentato = 1) Then Movimentato = 0
        End If
    End If
    If ((Me.CDImballoPrimario.KeyFieldID > 0) And (Me.txtNumeroConfImballo.Value > 0)) Then
        Me.List1.AddItem "Generazione movimenti di scarico imballo primario della merce nel magazzino " & Me.cboMagazzinoConf.Text
        Me.List1.ListIndex = Me.List1.ListCount - 1
        DoEvents
        
        If GeneraMovimentoScaricoImballo(Link_CaricoMerceRighe, IDAssegnazioneMerce, Link_ProcessoIVGamma, CODICE_LOTTO_ENTRATA, CODICE_LOTTO_CAMPAGNA, Me.CDImballoPrimario.KeyFieldID, GET_LINK_UM_ARTICOLO(Me.CDImballoPrimario.KeyFieldID), Me.CDImballoPrimario.Description, fnNotNullN(rs!Colli) * Me.txtNumeroConfImballo.Value, 0, MAX_QUANTITA_CONFEZ) = False Then
           If (Movimentato = 1) Then Movimentato = 0
        End If
    End If

    GeneraMovimentoScaricoKit Link_CaricoMerceRighe, IDAssegnazioneMerce, Link_ProcessoIVGamma, CODICE_LOTTO_ENTRATA, CODICE_LOTTO_CAMPAGNA
    

End If

rs.CloseResultset
Set rs = Nothing

CnDMT.CursorLocation = OLD_Cursor
Set mov = Nothing

sSQL = "UPDATE RV_POAssegnazioneMerce SET "
sSQL = sSQL & " Movimentato=" & Movimentato
sSQL = sSQL & " WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce
CnDMT.Execute sSQL

Exit Sub
ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
    
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
Private Function GET_FUNZIONE_MAGAZZINO(IDTipoDocumentoCoop As Long, IDTipoProcesso As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POProcessiDocumentoCoop.IDFunzione "
sSQL = sSQL & "FROM RV_POProcessiDocumentoCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE RV_POSchemaCoop.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDDocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDTipoProcessoCoop=" & IDTipoProcesso

Set rs = CnDMT.OpenResultset(sSQL)


If rs.EOF Then
    Select Case IDTipoProcesso
        Case 1 'Carico
            GET_FUNZIONE_MAGAZZINO = fnGetParametriMagazzino("IDCausale_Carico_Mag_Vendita")
        Case 2 'Scarico
            GET_FUNZIONE_MAGAZZINO = fnGetParametriMagazzino("IDCausale_Scarico_Mag_vendita")
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
Private Function GeneraMovimentoDiCarico(IDAssegnazione As Long, IDRigaConferimento As Long, IDProcessoIVGamma As Long, CodiceLottoVendita As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, DataLavorazione As String, QuantitaMovimentata As Double, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double, IDImballoPrim As Long, CodiceImballoPrim As String, DescrizioneImballoPrim As String, _
NumeroConfezioni As Long, TaraUnitariaConfezione As Double, CostoImballoConfezione As Double) As Boolean

On Error GoTo ERR_GeneraMovimentoDiCarico

mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = fncEsercizio(DataLavorazione)
mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
mov.IDOggetto = IDRigaConferimento
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = Me.CboMagazzinoVend.CurrentID
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
If IDProcessoIVGamma = 0 Then
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
    mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 1)
Else
    mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcessoIVGamma) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
    mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 1)
End If

mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 1
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", IDProcessoIVGamma
mov.Field "RV_POIDAnagraficaSocio", LINK_ANAGRAFICA_SOCIO
mov.Field "RV_PODataConferimento", DATA_CONFERIMENTO
mov.Field "RV_PONumeroConferimento", NUMERO_CONFERIMENTO
mov.Field "RV_POCodiceLotto", CODICE_LOTTO_ENTRATA
mov.Field "RV_POCodiceLottoCampagna", CODICE_LOTTO_CAMPAGNA
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", QuantitaMovimentata
mov.Field "RV_PONumeroColli", Colli
mov.Field "RV_POPesoLordo", PesoLordo
mov.Field "RV_POPesoNetto", PesoNetto
mov.Field "RV_POTara", Tara
mov.Field "RV_POQuantitaPezzi", Pezzi

mov.Field "RV_PODataLavorazione", DataLavorazione
mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
mov.Field "RV_POIDCalibro", IDCalibro
mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf

mov.Field "RV_POIDPedana", IDPedana
mov.Field "RV_POIDTipoPedana", IDTipoPedana
mov.Field "RV_POCodicePedana", CodicePedana
mov.Field "RV_POPesoPedana", PesoPedana

mov.Field "RV_POIDImballoPrim", IDImballoPrim
mov.Field "RV_POCodiceImballoPrim", CodiceImballoPrim
mov.Field "RV_PODescrizioneImballoPrim", DescrizioneImballoPrim
mov.Field "RV_PONumeroConfezioniPerImballo", NumeroConfezioni
mov.Field "RV_POTaraConfezioneImballo", TaraUnitariaConfezione
mov.Field "RV_POQuantitaTotaleConfImballo", Colli * NumeroConfezioni
mov.Field "RV_POCostoConfezioneImballo", CostoImballoConfezione

mov.Field "TipoRiga", trcNessuno

CnDMT.BeginTrans
GeneraMovimentoDiCarico = mov.Insert

CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoDiCarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function
Private Function GeneraMovimentoDiScarico(IDAssegnazione As Long, Qta_UM As Double, DataLavorazione As String, IDRigaConferimento As Long) As Boolean
On Error GoTo ERR_GeneraMovimentoDiScarico
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String

QuantitaRimasta = Qta_UM



If QuantitaRimasta > 0 Then

    mov.DataMovimento = DataLavorazione
    mov.FattoreDiConversione = Null
    
    mov.GestioneMatricole = False
    mov.IDEsercizio = fncEsercizio(DataLavorazione)
    mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    mov.IDOggetto = IDRigaConferimento
    mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoEntrata = Me.cboMagazzinoConf.CurrentID
    mov.IDMagazzinoUscita = Me.cboMagazzinoConf.CurrentID
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", LINK_ARTICOLO_CONFERITO
    mov.Field "IDUnitaDiMisura", LINK_UNITA_DI_MISURA_ARTICOLO_CONFERITO
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", ARTICOLO_CONFERITO
    mov.Field "QuantitaTotale", Qta_UM
    mov.Field "Importo", 0
    mov.Field "DataDocumento", DataLavorazione
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
    mov.Field "IDTipoMovimento", 1
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    mov.Field "RV_POTipoRiga", 0
    mov.Field "RV_POIDCaricoMerceRighe", 0
    mov.Field "RV_POIDAssegnazioneMerce", 0
    mov.Field "RV_POIDProcessoIVGamma", 0
    mov.Field "RV_POIDAnagraficaSocio", 0
    mov.Field "RV_PODataConferimento", ""
    mov.Field "RV_PONumeroConferimento", ""
    mov.Field "RV_POCodiceLotto", ""
    mov.Field "RV_POCodiceLottoCampagna", ""
    mov.Field "RV_POCodiceLottoVendita", ""
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    mov.Field "RV_POQuantitaMovimentata", 0
    mov.Field "RV_PONumeroColli", 0
    mov.Field "RV_POPesoLordo", 0
    mov.Field "RV_POPesoNetto", 0
    mov.Field "RV_POTara", 0
    mov.Field "RV_POQuantitaPezzi", 0
    
    
    mov.Field "RV_POIDImballoPrim", 0
    mov.Field "RV_POCodiceImballoPrim", ""
    mov.Field "RV_PODescrizioneImballoPrim", ""
    mov.Field "RV_PONumeroConfezioniPerImballo", 0
    mov.Field "RV_POTaraConfezioneImballo", 0
    mov.Field "RV_POQuantitaTotaleConfImballo", 0
    mov.Field "RV_POCostoConfezioneImballo", 0
    
    mov.Field "TipoRiga", trcNessuno
    CnDMT.BeginTrans
    GeneraMovimentoDiScarico = mov.Insert
    CnDMT.CommitTrans
End If
Exit Function
ERR_GeneraMovimentoDiScarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans

End Function

Public Function fnGetParametriMagazzino(NomeCampo As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE ((IDUtente=" & TheApp.IDUser & ") "
    sSQL = sSQL & "AND (IDFiliale=" & TheApp.Branch & "))"
    
    Set rsEse = CnDMT.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
            fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
        Else
            sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
            sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
            sSQL = sSQL & "AND (IDUtente=0))"
        
            Set rsEse = CnDMT.OpenResultset(sSQL)
        
            If rsEse.EOF = False Then
                If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                    fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
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
        
        Set rsEse = CnDMT.OpenResultset(sSQL)
        
        If rsEse.EOF = False Then
            If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
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
Public Function GeneraMovimentoCaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcesso As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double) As Boolean
On Error GoTo ERR_GeneraMovimentoCaricoImballo
mov.DataMovimento = Me.txtDataLavorazione.Text
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = fncEsercizio(Me.txtDataLavorazione.Text)
mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
mov.IDOggetto = IDRigaConferimento

mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = CboMagazzinoVend.CurrentID
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", Me.txtDataLavorazione.Text
If IDProcesso = 0 Then
    mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text
    mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 1)
Else
    mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
    mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 1)
End If

mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 2
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", IDProcesso
mov.Field "RV_POIDAnagraficaSocio", LINK_ANAGRAFICA_SOCIO
mov.Field "RV_PODataConferimento", DATA_CONFERIMENTO
mov.Field "RV_PONumeroConferimento", NUMERO_CONFERIMENTO
mov.Field "RV_POCodiceLotto", CODICE_LOTTO_ENTRATA
mov.Field "RV_POCodiceLottoCampagna", CODICE_LOTTO_CAMPAGNA
mov.Field "RV_POCodiceLottoVendita", Me.txtLottoVendita.Text
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0


mov.Field "TipoRiga", trcNessuno

CnDMT.BeginTrans
    GeneraMovimentoCaricoImballo = mov.Insert
CnDMT.CommitTrans

Exit Function
ERR_GeneraMovimentoCaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function

Public Function GeneraMovimentoScaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcesso As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, TracciaImballo As Long, ColliMax As Double) As Boolean
On Error GoTo ERR_GeneraMovimentoScaricoImballo
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String
    
rsLottoImballo.Filter = "IDArticoloImballo=" & IDArticolo

If Not ((rsLottoImballo.EOF) And (rsLottoImballo.BOF)) Then

    rsLottoImballo.MoveFirst
    
    While Not rsLottoImballo.EOF
            
        mov.DataMovimento = Me.txtDataLavorazione.Text
        mov.FattoreDiConversione = Null
        
        mov.GestioneMatricole = False
        mov.IDEsercizio = fncEsercizio(Me.txtDataLavorazione)
        mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
        mov.IDOggetto = IDRigaConferimento
        
        mov.IDUtente = TheApp.IDUser
        mov.IDMagazzinoUscita = cboMagazzinoConf.CurrentID
        mov.Cessione = 0
        mov.Field "IDAzienda", TheApp.IDFirm
        mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
        mov.Field "IDTipoAnagrafica", 3
        mov.Field "IDArticolo", IDArticolo
        mov.Field "IDUnitaDiMisura", IDUMDiamante
        mov.Field "IDcambio", Null
        mov.Field "DescrizioneArticolo", Articolo
        mov.Field "QuantitaTotale", (Qta_UM / ColliMax) * fnNotNullN(rsLottoImballo!QuantitaMovimentata)
        mov.Field "Importo", 0
        mov.Field "DataDocumento", Me.txtDataLavorazione.Text
        If IDProcesso = 0 Then
            mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text
            mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
        Else
            mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
            mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 2)
        End If
        mov.Field "IDTipoMovimento", 1
        
        'DATI DI CONFERIMENTO
        mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
        mov.Field "RV_POTipoRiga", 2
        mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
        mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
        mov.Field "RV_POIDProcessoIVGamma", IDProcesso
        mov.Field "RV_POIDAnagraficaSocio", LINK_ANAGRAFICA_SOCIO
        mov.Field "RV_PODataConferimento", DATA_CONFERIMENTO
        mov.Field "RV_PONumeroConferimento", NUMERO_CONFERIMENTO
        mov.Field "RV_POCodiceLotto", CODICE_LOTTO_ENTRATA
        mov.Field "RV_POCodiceLottoCampagna", CODICE_LOTTO_CAMPAGNA
        mov.Field "RV_POCodiceLottoVendita", Me.txtLottoVendita.Text
        mov.Field "RV_POQuantitaLiquidazione", 0
        mov.Field "RV_POImportoInclusoImballo", 0
        mov.Field "RV_POImportoLiquidazione", 0
        mov.Field "RV_POQuantitaMovimentata", 0
        mov.Field "RV_PONumeroColli", 0
        mov.Field "RV_POPesoLordo", 0
        mov.Field "RV_POPesoNetto", 0
        mov.Field "RV_POTara", 0
        mov.Field "RV_POQuantitaPezzi", 0

        mov.Field "RV_PODataLavorazione", Null
        mov.Field "RV_POIDTipoLavorazione", 0
        mov.Field "RV_POIDCalibro", 0
        mov.Field "RV_POIDTipoCategoria", 0
        mov.Field "RV_POIDTipoLavorazioneConf", 0
        mov.Field "RV_POPrezzoMedioConf", 0
        
        mov.Field "RV_POIDPedana", 0
        mov.Field "RV_POIDTipoPedana", 0
        mov.Field "RV_POCodicePedana", ""
        mov.Field "RV_POPesoPedana", ""
        
        mov.Field "RV_POIDImballoPrim", 0
        mov.Field "RV_POCodiceImballoPrim", ""
        mov.Field "RV_PODescrizioneImballoPrim", ""
        mov.Field "RV_PONumeroConfezioniPerImballo", 0
        mov.Field "RV_POTaraConfezioneImballo", 0
        mov.Field "RV_POQuantitaTotaleConfImballo", 0
        mov.Field "RV_POCostoConfezioneImballo", 0
                                
        mov.Field "RV_POIDLottoImballo", fnNotNullN(rsLottoImballo!IDLottoImballo)
        mov.Field "LottoImballo", fnNotNull(rsLottoImballo!CodiceLottoImballo)
        
        
        mov.Field "TipoRiga", trcNessuno
        CnDMT.BeginTrans
        GeneraMovimentoScaricoImballo = mov.Insert
        CnDMT.CommitTrans
            
    rsLottoImballo.MoveNext
    Wend
    
Else

    mov.DataMovimento = Me.txtDataLavorazione.Text
    mov.FattoreDiConversione = Null
    
    mov.GestioneMatricole = False
    mov.IDEsercizio = fncEsercizio(Me.txtDataLavorazione)
    mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    mov.IDOggetto = IDRigaConferimento
    
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = cboMagazzinoConf.CurrentID
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", IDArticolo
    mov.Field "IDUnitaDiMisura", IDUMDiamante
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", Articolo
    mov.Field "QuantitaTotale", Qta_UM
    mov.Field "Importo", 0
    mov.Field "DataDocumento", Me.txtDataLavorazione.Text
    If IDProcesso = 0 Then
        mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
    Else
        mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 2)
    End If
    mov.Field "IDTipoMovimento", 1
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    mov.Field "RV_POTipoRiga", 2
    mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
    mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
    mov.Field "RV_POIDProcessoIVGamma", IDProcesso
    mov.Field "RV_POIDAnagraficaSocio", LINK_ANAGRAFICA_SOCIO
    mov.Field "RV_PODataConferimento", DATA_CONFERIMENTO
    mov.Field "RV_PONumeroConferimento", NUMERO_CONFERIMENTO
    mov.Field "RV_POCodiceLotto", CODICE_LOTTO_ENTRATA
    mov.Field "RV_POCodiceLottoCampagna", CODICE_LOTTO_CAMPAGNA
    mov.Field "RV_POCodiceLottoVendita", Me.txtLottoVendita.Text
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    mov.Field "RV_POQuantitaMovimentata", 0
    mov.Field "RV_PONumeroColli", 0
    mov.Field "RV_POPesoLordo", 0
    mov.Field "RV_POPesoNetto", 0
    mov.Field "RV_POTara", 0
    mov.Field "RV_POQuantitaPezzi", 0
    
    
    mov.Field "RV_PODataLavorazione", Null
    mov.Field "RV_POIDTipoLavorazione", 0
    mov.Field "RV_POIDCalibro", 0
    mov.Field "RV_POIDTipoCategoria", 0
    mov.Field "RV_POIDTipoLavorazioneConf", 0
    mov.Field "RV_POPrezzoMedioConf", 0
    
    mov.Field "RV_POIDPedana", 0
    mov.Field "RV_POIDTipoPedana", 0
    mov.Field "RV_POCodicePedana", ""
    mov.Field "RV_POPesoPedana", ""
    
    mov.Field "RV_POIDImballoPrim", 0
    mov.Field "RV_POCodiceImballoPrim", ""
    mov.Field "RV_PODescrizioneImballoPrim", ""
    mov.Field "RV_PONumeroConfezioniPerImballo", 0
    mov.Field "RV_POTaraConfezioneImballo", 0
    mov.Field "RV_POQuantitaTotaleConfImballo", 0
    mov.Field "RV_POCostoConfezioneImballo", 0

    mov.Field "RV_POIDLottoImballo", 0
    mov.Field "LottoImballo", ""
    
    
    mov.Field "TipoRiga", trcNessuno
    CnDMT.BeginTrans
    GeneraMovimentoScaricoImballo = mov.Insert
    CnDMT.CommitTrans
    
End If
Exit Function
ERR_GeneraMovimentoScaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
End Function
Private Function GET_LINK_UM_ARTICOLO(IDArticolo) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraVendita FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_ARTICOLO = 0
Else
    GET_LINK_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisuraVendita)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function ean13$(chaine$)

  Dim I%, checksum%, first%, CodeBarre$, tableA As Boolean
  ean13$ = ""
  If Len(chaine$) = 12 Then
    For I% = 1 To 12
      If Asc(Mid$(chaine$, I%, 1)) < 48 Or Asc(Mid$(chaine$, I%, 1)) > 57 Then
        I% = 0
        Exit For
      End If
    Next
    If I% = 13 Then
      For I% = 12 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, I%, 1))
      Next
      checksum% = checksum% * 3
      For I% = 11 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, I%, 1))
      Next
      chaine$ = chaine$ & (10 - checksum% Mod 10) Mod 10
      CodeBarre$ = Left$(chaine$, 1) & Chr$(65 + Val(Mid$(chaine$, 2, 1)))
      first% = Val(Left$(chaine$, 1))
      For I% = 3 To 7
        tableA = False
         Select Case I%
         Case 3
           Select Case first%
           Case 0 To 3
             tableA = True
           End Select
         Case 4
           Select Case first%
           Case 0, 4, 7, 8
             tableA = True
           End Select
         Case 5
           Select Case first%
           Case 0, 1, 4, 5, 9
             tableA = True
           End Select
         Case 6
           Select Case first%
           Case 0, 2, 5, 6, 7
             tableA = True
           End Select
         Case 7
           Select Case first%
           Case 0, 3, 6, 8, 9
             tableA = True
           End Select
         End Select
       If tableA Then
         CodeBarre$ = CodeBarre$ & Chr$(65 + Val(Mid$(chaine$, I%, 1)))
       Else
         CodeBarre$ = CodeBarre$ & Chr$(75 + Val(Mid$(chaine$, I%, 1)))
       End If
     Next
      CodeBarre$ = CodeBarre$ & "*"   'Ajout séparateur central / Add middle separator
      For I% = 8 To 13
        CodeBarre$ = CodeBarre$ & Chr$(97 + Val(Mid$(chaine$, I%, 1)))
      Next
      CodeBarre$ = CodeBarre$ & "+"   'Ajout de la marque de fin / Add end mark
      ean13$ = CodeBarre$
    End If
  End If
End Function

Private Sub BLOCCA_PER_SEMAFORO(IDRigaConferimento As Long)
On Error GoTo ERR_BLOCCA_PER_SEMAFORO
Dim sSQL As String
Dim rs As ADODB.Recordset

Dim LINK_TIPO_OGGETTO As Long
Dim LINK_FUNZIONE As Long

LINK_TIPO_OGGETTO = fnGetTipoOggetto("RV_POAssegnazioneMerce")
LINK_FUNZIONE = GET_FUNZIONE(LINK_TIPO_OGGETTO)

sSQL = "SELECT * FROM Semaforo"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDUtente = TheApp.IDUser
    rs!IDAzione = 1
    'rs!IDAzienda = TheApp.IDFirm
    rs!IDFiliale = TheApp.Branch
    rs!IDFunzione = LINK_FUNZIONE
    rs!IDTipoOggetto = LINK_TIPO_OGGETTO
    rs!IDOggetto = IDRigaConferimento
    rs!DataAttivazione = Date
    rs!Variato = 0
    rs!IDSemaforo = fnGetNewKey("Semaforo", "IDSemaforo")

CnDMT.BeginTrans
rs.Update
CnDMT.CommitTrans

rs.Close
Set rs = Nothing

Exit Sub
ERR_BLOCCA_PER_SEMAFORO:
    MsgBox Err.Description, vbCritical, "Blocco semaforo"
    CnDMT.RollbackTrans
    Unload Me

End Sub
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
Private Sub SBLOCCA_PER_SEMAFORO(IDRigaConferimento As Long)
Dim sSQL As String
Dim LINK_TIPO_OGGETTO As Long
Dim LINK_FUNZIONE As Long

LINK_TIPO_OGGETTO = fnGetTipoOggetto("RV_POAssegnazioneMerce")
LINK_FUNZIONE = GET_FUNZIONE(LINK_TIPO_OGGETTO)


sSQL = "DELETE FROM Semaforo "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOggetto=" & IDRigaConferimento
sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE
sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO

CnDMT.Execute sSQL
End Sub
Private Sub ParametroPesoArticolo()
On Error GoTo ERR_ParametroPesoArticolo
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPesoArticolo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    TIPO_PESO_ARTICOLO = fnNotNullN(rs!IDRV_POTipoPesoArticolo)
Else
    TIPO_PESO_ARTICOLO = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ParametroPesoArticolo:
    MsgBox Err.Description, vbCritical, "ParametroPesoArticolo"
End Sub
Public Function ParametroTipoArrotondamento() As Long
On Error GoTo ERR_ParametroTipoArrotondamento
Dim rsEse As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"
       
Set rsEse = CnDMT.OpenResultset(sSQL)

If rsEse.EOF = False Then
    Link_Arrontondamento = fnNotNullN(rsEse("IDTipoArrotondamento"))
Else
    Link_Arrontondamento = 0
End If

rsEse.CloseResultset
Set rsEse = Nothing
Exit Function
ERR_ParametroTipoArrotondamento:
    MsgBox Err.Description, vbCritical, "ParametroTipoArrotondamento"
End Function

Private Sub txtTara_LostFocus()
    CalcoloPesoNetto
End Sub
Private Sub SCRIVI_CODA(IDOggetto As Long, IDTipoOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

CnDMT.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rs = New ADODB.Recordset

rs.Open "RV_POTMP", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDSessione = fnGetNewKey("RV_POTMP", "IDSessione")
    rs!IDUtente = TheApp.IDUser
    rs!IDTipoOggetto = IDTipoOggetto
    rs!IDOggetto = IDOggetto
    rs!Utente = TheApp.User
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Function GET_PREZZO_IMBALLO_INCLUSO(IDArticoloImballo As Long, IDCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCli As DmtOleDbLib.adoResultset


sSQL = "SELECT PrezzoInclusoImballo "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteImb "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "SELECT PrezzoInclusoImballo "
    sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    'sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    Set rsCli = CnDMT.OpenResultset(sSQL)
    
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
End Function
Private Function GET_PREZZO_IMBALLO_2(IDArticolo As Long, IDListinoCliente As Long, IDListinoAzienda As Long, IDListinoParAzienda As Long, IDOggettoOrdine As Long, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDDestinazione As Long, IDCliente As Long, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long) As Double
On Error GoTo ERR_GET_PREZZO_IMBALLO_2
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Listino_Dest As Long
Dim Link_Riga_Ordine As Long

ImportoUnitario = 0

If PrezziDaOrdine = 1 Then
    Link_Riga_Ordine = 0
    IDArticoloPadre = GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
    IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
        sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
        If (TROVA_PREZZI_ORD_CAT = 1) Then
            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
        End If
        If (TROVA_PREZZI_ORD_CAL = 1) Then
            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
        End If
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            '''''''''''''''''''TROVO IL LINK_RIGA DELL'ORDINE'''''''''''''''''''''''''''
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'            sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
            If (TROVA_PREZZI_ORD_CAT = 1) Then
                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
            End If
            If (TROVA_PREZZI_ORD_CAL = 1) Then
                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
            End If
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                Link_Riga_Ordine = fnNotNullN(rs!RV_POLinkRiga)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Link_Riga_Ordine > 0 Then
                sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
                sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
                sSQL = sSQL & " AND RV_POTipoRiga=2 "
                sSQL = sSQL & " AND RV_POLinkRiga=" & Link_Riga_Ordine
                sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo
                
                Set rs = CnDMT.OpenResultset(sSQL)
                
                If Not rs.EOF Then
                    ImportoUnitario = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                End If
                
                rs.CloseResultset
                Set rs = Nothing
            End If
            
        End If
    End If
End If

'''''IMPORTO UNITARIO DAL LISTINO CLIENTE'''''''''''''''''''''''''''''''''''''''''''''''''''
If ImportoUnitario = 0 Then
    sSQL = "SELECT PrezzoNettoIva "
    sSQL = sSQL & "FROM ListinoPerArticolo "
    sSQL = sSQL & " WHERE IDListino=" & IDListinoCliente
    sSQL = sSQL & " AND IDArticolo=" & IDImballo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        ImportoUnitario = 0
    Else
        ImportoUnitario = fnNotNullN(rs!PrezzoNettoIva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''IMPORTO UNITARIO DAL LISTINO IMBALLI INDICATO NEI PARAMETRI FILIALE'''''''''''''''''''''''
If ImportoUnitario = 0 Then
    sSQL = "SELECT PrezzoNettoIva "
    sSQL = sSQL & "FROM ListinoPerArticolo "
    sSQL = sSQL & " WHERE IDListino=" & IDListinoParAzienda
    sSQL = sSQL & " AND IDArticolo=" & IDImballo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        ImportoUnitario = 0
    Else
        ImportoUnitario = fnNotNullN(rs!PrezzoNettoIva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''IMPORTO UNITARIO DAL LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''''''
If ImportoUnitario = 0 Then
    sSQL = "SELECT PrezzoNettoIva "
    sSQL = sSQL & "FROM ListinoPerArticolo "
    sSQL = sSQL & " WHERE IDListino=" & IDListinoAzienda
    sSQL = sSQL & " AND IDArticolo=" & IDImballo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        ImportoUnitario = 0
    Else
        ImportoUnitario = fnNotNullN(rs!PrezzoNettoIva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GET_PREZZO_IMBALLO_2 = ImportoUnitario
Exit Function
ERR_GET_PREZZO_IMBALLO_2:
    GET_PREZZO_IMBALLO_2 = 0
End Function
Private Function GET_PREZZO_IMBALLO_INCLUSO_2(IDArticolo As Long, IDOggettoOrdine As Long, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDCliente As Long, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long) As Long
On Error GoTo ERR_GET_PREZZO_IMBALLO_INCLUSO_2
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCli As DmtOleDbLib.adoResultset
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Listino_Dest As Long

GET_PREZZO_IMBALLO_INCLUSO_2 = 0

'If PrezziDaOrdine = 1 Then
'    IDArticoloPadre = GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
'    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
'
'    If IDArticoloPadre > 0 Then
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
'        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'        sSQL = sSQL & " AND RV_POTipoRiga=1 "
'        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'        If GESTIONE_ORDINE_VIVAIO = 0 Then
'            If TROVA_PREZZI_NO_IMB = 0 Then
'                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'            End If
'        End If
'
'        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'        If (TROVA_PREZZI_ORD_CAT = 1) Then
'            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'        End If
'        If (TROVA_PREZZI_ORD_CAL = 1) Then
'            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'        End If
'        Set rs = CnDMT.OpenResultset(sSQL)
'
'        If rs.EOF Then
'            NumeroCombinazioni = 0
'        Else
'            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
'        End If
'
'        rs.CloseResultset
'        Set rs = Nothing
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If NumeroCombinazioni = 1 Then
'            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
'            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'            sSQL = sSQL & " AND RV_POTipoRiga=1 "
'            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'            If GESTIONE_ORDINE_VIVAIO = 0 Then
'                If TROVA_PREZZI_NO_IMB = 0 Then
'                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'                End If
'            End If
'
'            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'            If (TROVA_PREZZI_ORD_CAT = 1) Then
'                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'            End If
'            If (TROVA_PREZZI_ORD_CAL = 1) Then
'                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'            End If
'
'            Set rs = CnDMT.OpenResultset(sSQL)
'
'            If Not rs.EOF Then
'                GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rs!RV_POImportoImballoInArticolo)
'            End If
'
'            rs.CloseResultset
'            Set rs = Nothing
'            Exit Function
'        End If
'    End If
'End If


If GET_PREZZO_IMBALLO_INCLUSO_2 = 0 Then
    sSQL = "SELECT PrezzoInclusoImballo "
    sSQL = sSQL & "FROM RV_POConfigurazioneClienteImb "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDArticoloImballo=" & IDImballo
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        sSQL = "SELECT PrezzoInclusoImballo "
        sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
        sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
        sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
        'sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
        
        Set rsCli = CnDMT.OpenResultset(sSQL)
        
        If rsCli.EOF Then
            GET_PREZZO_IMBALLO_INCLUSO_2 = 0
        Else
            GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rsCli!PrezzoInclusoImballo)
        End If
        
        rsCli.CloseResultset
        Set rsCli = Nothing
        
    Else
        GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rs!PrezzoInclusoImballo)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

Exit Function
ERR_GET_PREZZO_IMBALLO_INCLUSO_2:
    GET_PREZZO_IMBALLO_INCLUSO_2 = 0
End Function
Private Function GET_PREZZO_UNITARIO_ARTICOLO(IDArticolo As Long, IDListinoCliente As Long, IDListinoAzienda As Long, IDOggettoOrdine As Long, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, rstmp As ADODB.Recordset, IDDestinazione As Long, IDCliente As Long) As Double
On Error GoTo ERR_GET_PREZZO_UNITARIO_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Listino_Dest As Long

ImportoUnitario = 0

If PrezziDaOrdine = 1 Then
    IDArticoloPadre = IDArticolo
    IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
        sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                ImportoUnitario = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                rstmp!Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
                rstmp!Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
        End If
    End If
End If


'''''IMPORTO UNITARIO DAL LISTINO CLIENTE'''''''''''''''''''''''''''''''''''''''''''''''''''
If ImportoUnitario = 0 Then
    sSQL = "SELECT PrezzoNettoIva "
    sSQL = sSQL & "FROM ListinoPerArticolo "
    sSQL = sSQL & " WHERE IDListino=" & IDListinoCliente
    sSQL = sSQL & " AND IDArticolo=" & IDArticolo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        ImportoUnitario = 0
    Else
        ImportoUnitario = fnNotNullN(rs!PrezzoNettoIva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If ImportoUnitario = 0 Then
    '''''IMPORTO UNITARIO DAL LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT PrezzoNettoIva "
    sSQL = sSQL & "FROM ListinoPerArticolo "
    sSQL = sSQL & " WHERE IDListino=" & IDListinoAzienda
    sSQL = sSQL & " AND IDArticolo=" & IDArticolo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        ImportoUnitario = 0
    Else
        ImportoUnitario = fnNotNullN(rs!PrezzoNettoIva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

GET_PREZZO_UNITARIO_ARTICOLO = ImportoUnitario
Exit Function
ERR_GET_PREZZO_UNITARIO_ARTICOLO:
    GET_PREZZO_UNITARIO_ARTICOLO = 0
    
End Function
Private Function GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticoloDerivato As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POArticoloFiglioOrdine "
sSQL = sSQL & "WHERE IDArticoloFiglio=" & IDArticoloDerivato

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ARTICOLO_PADRE_ORDINATO = 0
Else
    GET_LINK_ARTICOLO_PADRE_ORDINATO = fnNotNullN(rs!IDArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_LISTINO_CLIENTE(IDAnagrafica As Long, IDDestinazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_LISTINO_CLIENTE = 0

'''''LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListino "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteListino "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDDestinazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LISTINO_CLIENTE = 0
Else
    GET_LINK_LISTINO_CLIENTE = fnNotNullN(rs!IDListino)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_LISTINO_CLIENTE > 0 Then Exit Function

'''''LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LISTINO_CLIENTE = 0
Else
    GET_LINK_LISTINO_CLIENTE = fnNotNullN(rs!IDListinoDefault)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Sub GET_CONFIGURAZIONE_IMPORTI_ARTICOLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long

ImportoUnitario = 0

'If PrezziDaOrdine = 1 Then
'    IDArticoloPadre = IDArticolo
'    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
'
'    If IDArticoloPadre > 0 Then
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
'        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'        sSQL = sSQL & " AND RV_POTipoRiga=1 "
'        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'        If GESTIONE_ORDINE_VIVAIO = 0 Then
'                If TROVA_PREZZI_NO_IMB = 0 Then
'                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'                End If
'            End If
'
'        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'        If (TROVA_PREZZI_ORD_CAT = 1) Then
'            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'        End If
'        If (TROVA_PREZZI_ORD_CAL = 1) Then
'            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'        End If
'        Set rs = CnDMT.OpenResultset(sSQL)
'
'        If rs.EOF Then
'            NumeroCombinazioni = 0
'        Else
'            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
'        End If
'
'        rs.CloseResultset
'        Set rs = Nothing
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If NumeroCombinazioni = 1 Then
'            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
'            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'            sSQL = sSQL & " AND RV_POTipoRiga=1 "
'            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'            If GESTIONE_ORDINE_VIVAIO = 0 Then
'                If TROVA_PREZZI_NO_IMB = 0 Then
'                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'                End If
'            End If
'
'            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'            If (TROVA_PREZZI_ORD_CAT = 1) Then
'                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'            End If
'            If (TROVA_PREZZI_ORD_CAL = 1) Then
'                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'            End If
'            Set rs = CnDMT.OpenResultset(sSQL)
'
'            If Not rs.EOF Then
'                Me.txtImportoUnitarioArticolo.Value = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'
'                Me.txtSconto1.Value = fnNotNullN(rs!Art_sco_in_percentuale_1)
'                Me.txtSconto2.Value = fnNotNullN(rs!Art_sco_in_percentuale_2)
'                ImportoUnitario = Me.txtImportoUnitarioArticolo.Value
'            End If
'
'            rs.CloseResultset
'            Set rs = Nothing
'        End If
'    End If
'End If
'
'If ImportoUnitario > 0 Then Exit Sub

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal
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

Me.txtImportoUnitarioArticolo.Value = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))

Me.txtSconto1.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglioLocal))
Me.txtSconto2.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglioLocal))


End Sub
Private Sub GET_CONFIGURAZIONE_IMPORTI_IMBALLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Riga_Ordine As Long

ImportoUnitario = 0

'If PrezziDaOrdine = 1 Then
'    Link_Riga_Ordine = 0
'    IDArticoloPadre = IDArticolo
'    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
'
'    If IDArticoloPadre > 0 Then
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
'        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'        sSQL = sSQL & " AND RV_POTipoRiga=1 "
'        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'        If GESTIONE_ORDINE_VIVAIO = 0 Then
'            If TROVA_PREZZI_NO_IMB = 0 Then
'                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'            End If
'        End If
'
'        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'        If (TROVA_PREZZI_ORD_CAT = 1) Then
'            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'        End If
'        If (TROVA_PREZZI_ORD_CAL = 1) Then
'            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'        End If
'
'        Set rs = CnDMT.OpenResultset(sSQL)
'
'        If rs.EOF Then
'            NumeroCombinazioni = 0
'        Else
'            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
'        End If
'
'        rs.CloseResultset
'        Set rs = Nothing
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If NumeroCombinazioni = 1 Then
'            '''''''''''''''''''TROVO IL LINK_RIGA DELL'ORDINE'''''''''''''''''''''''''''
'            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
'            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'            sSQL = sSQL & " AND RV_POTipoRiga=1 "
'            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'            If GESTIONE_ORDINE_VIVAIO = 0 Then
'                If TROVA_PREZZI_NO_IMB = 0 Then
'                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'                End If
'            End If
'            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'            If (TROVA_PREZZI_ORD_CAT = 1) Then
'                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'            End If
'            If (TROVA_PREZZI_ORD_CAL = 1) Then
'                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'            End If
'
'            Set rs = CnDMT.OpenResultset(sSQL)
'
'            If Not rs.EOF Then
'                Link_Riga_Ordine = fnNotNullN(rs!RV_POLinkRiga)
'            End If
'
'            rs.CloseResultset
'            Set rs = Nothing
'            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'            If Link_Riga_Ordine > 0 Then
'                sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
'                sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'                sSQL = sSQL & " AND RV_POTipoRiga=2 "
'                sSQL = sSQL & " AND RV_POLinkRiga=" & Link_Riga_Ordine
'                sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo
'
'                Set rs = CnDMT.OpenResultset(sSQL)
'
'                If Not rs.EOF Then
'                    Me.txtImportoUnitarioImballo.Value = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'                    ImportoUnitario = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'                End If
'
'                rs.CloseResultset
'                Set rs = Nothing
'            End If
'
'        End If
'    End If
'End If
'
'If ImportoUnitario > 0 Then Exit Sub

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal
ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.ReadDataFromArticle IDImballo, sTabellaDettaglioLocal
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

End Sub



Private Function GET_LINK_ATTIVITA_AZIENDA(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "WHERE (Azienda.IDAzienda =" & IDAzienda & ") And (Filiale.IDFiliale = " & IDFiliale & ")"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function MOVIMENTAZIONE_RIGA_QUADRATURA(IDRigaQuadratura As Long) As String
Dim OLD_Cursor As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim IDTipoProdotto As Long


    
Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection


MOVIMENTAZIONE_RIGA_QUADRATURA = ""

IDTipoProdotto = Link_TipoProdotto_Q

GeneraMovimentoDiCarico_Q IDRigaQuadratura, Link_TipoProdotto_Q
GeneraMovimentoDiScarico_Q IDRigaQuadratura, Link_TipoProdotto_Q


           

    
'Cn.CursorLocation = OLD_CURSOR
Set mov = Nothing

End Function

Private Function GeneraMovimentoDiCarico_Q(IDRiga As Long, IDTipoProdotto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIELavorazioneQuadratura "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & IDRiga

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then

    mov.DataMovimento = rs!DataDocumento
    mov.FattoreDiConversione = Null
    mov.GestioneMatricole = False
    mov.IDEsercizio = fncEsercizio(rs!DataDocumento)
    mov.IDTipoOggetto = fnGetTipoOggetto("RV_POLavorazioneL")
    mov.IDOggetto = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Select Case IDTipoProdotto
        Case Link_TipoCaloPeso
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleCaloPesoCarico")
            mov.Field "Oggetto", "Calo peso"
        Case Link_TipoAumentoPeso
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleAumentoPesoCarico")
            mov.Field "Oggetto", "Aumento di peso"
        Case Link_TipoScarto
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleScartoCarico")
            mov.Field "Oggetto", "Scarto"
    End Select
    
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = CboMagazzinoVend.CurrentID
    mov.IDMagazzinoEntrata = CboMagazzinoVend.CurrentID
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", fnNotNull(rs!IDAnagraficaSocio)
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", fnNotNullN(rs!IDArticolo)
    mov.Field "IDUnitaDiMisura", fnNotNullN(rs!IDUnitaDiMisura)
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", fnNotNull(rs!Articolo)
    mov.Field "QuantitaTotale", fnNotNullN(rs!Qta_UM)
    mov.Field "Importo", 0
    mov.Field "DataDocumento", Date
    
    mov.Field "IDTipoMovimento", 1
    mov.Field "TipoRiga", trcNessuno
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDRiga
    mov.Field "RV_POTipoRiga", 1
    mov.Field "RV_POIDCaricoMerceRighe", fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    mov.Field "RV_POIDAssegnazioneMerce", 0
    mov.Field "RV_POIDProcessoIVGamma", 0
    mov.Field "RV_POIDAnagraficaSocio", fnNotNullN(rs!IDAnagraficaSocio)
    mov.Field "RV_PODataConferimento", rs!DataDocumentoConferito
    mov.Field "RV_PONumeroConferimento", rs!NumeroDocumentoConferito
    mov.Field "RV_POCodiceLotto", fnNotNull(rs!CodiceLotto)
    mov.Field "RV_POCodiceLottoCampagna", fnNotNull(rs!LottoDiConferimento)
    mov.Field "RV_POCodiceLottoVendita", ""
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    
    Select Case fnNotNullN(rs!IDUnitaDiMisuraCoopConferito)
        Case 1
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!Colli)
        Case 2
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!PesoLordo)
        Case 3
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!PesoNetto)
        Case 4
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!Tara)
        Case 5
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!Pezzi)
        Case Else
            mov.Field "RV_POQuantitaMovimentata", fnNotNullN(rs!PesoNetto)
    End Select
End If
    
    rs.CloseResultset
    Set rs = Nothing

    CnDMT.BeginTrans
    GeneraMovimentoDiCarico_Q = mov.Insert
    
    If GeneraMovimentoDiCarico_Q = False Then
        CnDMT.RollbackTrans
    Else
        CnDMT.CommitTrans
    End If

End Function
Private Function GeneraMovimentoDiScarico_Q(IDRiga As Long, IDTipoProdotto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIELavorazioneQuadratura "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & IDRiga

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    mov.DataMovimento = rs!DataDocumento
    mov.FattoreDiConversione = Null
    mov.GestioneMatricole = False
    mov.IDEsercizio = fncEsercizio(rs!DataDocumento)
    mov.IDTipoOggetto = fnGetTipoOggetto("RV_POLavorazioneL")
    mov.IDOggetto = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Select Case IDTipoProdotto
        Case Link_TipoCaloPeso
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleCaloPeso")
            mov.Field "Oggetto", "Calo peso"
        Case Link_TipoAumentoPeso
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleAumentoPeso")
            mov.Field "Oggetto", "Aumento di peso"
        Case Link_TipoScarto
            mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleScarto")
            mov.Field "Oggetto", "Scarto"
    End Select
    
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = cboMagazzinoConf.CurrentID
    mov.IDMagazzinoEntrata = cboMagazzinoConf.CurrentID
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", fnNotNullN(rs!IDAnagraficaSocio)
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", fnNotNullN(rs!IDArticoloConferito)
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", fnNotNull(rs!ArticoloConferito)
    Select Case fnNotNullN(rs!IDUnitaDiMisuraCoopConferito)
        Case 1
            mov.Field "QuantitaTotale", fnNotNullN(rs!Colli)
        Case 2
            mov.Field "QuantitaTotale", fnNotNullN(rs!PesoLordo)
        Case 3
            mov.Field "QuantitaTotale", fnNotNullN(rs!PesoNetto)
        Case 4
            mov.Field "QuantitaTotale", fnNotNullN(rs!Tara)
        Case 5
            mov.Field "QuantitaTotale", fnNotNullN(rs!Pezzi)
        Case Else
            mov.Field "QuantitaTotale", fnNotNullN(rs!PesoNetto)
    End Select
    
    mov.Field "Importo", 0
    mov.Field "IDUnitaDiMisura", fnNotNullN(rs!IDUnitaDiMisuraDiamanteConferito)
    mov.Field "DataDocumento", Date
    
    mov.Field "IDTipoMovimento", 1
    mov.Field "TipoRiga", trcNessuno
    
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDRiga
    mov.Field "RV_POTipoRiga", 1
    mov.Field "RV_POIDCaricoMerceRighe", 0
    mov.Field "RV_POIDAssegnazioneMerce", 0
    mov.Field "RV_POIDProcessoIVGamma", 0
    mov.Field "RV_POIDAnagraficaSocio", ""
    mov.Field "RV_PODataConferimento", ""
    mov.Field "RV_PONumeroConferimento", 0
    mov.Field "RV_POCodiceLotto", ""
    mov.Field "RV_POCodiceLottoCampagna", ""
    mov.Field "RV_POCodiceLottoVendita", ""
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    mov.Field "RV_POQuantitaMovimentata", 0
    mov.Field "RV_PONumeroColli", 0
    mov.Field "RV_POPesoLordo", 0
    mov.Field "RV_POPesoNetto", 0
    mov.Field "RV_POTara", 0
    mov.Field "RV_POQuantitaPezzi", 0
End If

rs.CloseResultset
Set rs = Nothing

CnDMT.BeginTrans
GeneraMovimentoDiScarico_Q = mov.Insert
If GeneraMovimentoDiScarico_Q = False Then
    CnDMT.RollbackTrans
Else
    CnDMT.CommitTrans
End If

End Function
Private Function GET_CAUSALE_QUADRATURA(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_CAUSALE_QUADRATURA = fnNotNullN(rs.adoColumns(NomeCampo).Value)
Else
    GET_CAUSALE_QUADRATURA = 0
End If

rs.CloseResultset
Set rs = Nothing
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
Private Function GET_DATI_LAVORAZIONE(IDLavorazione As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DATI_LAVORAZIONE = ""
Else
    GET_DATI_LAVORAZIONE = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PESO_PEDANA(IDPedana As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PORepPedana "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_PEDANA = 0
Else
    If IDTipoPedana <> fnNotNullN(rs!IDRV_POTipoPedana) Then
        GET_PESO_PEDANA = 0
    Else
        GET_PESO_PEDANA = fnNotNullN(rs!PesoPedana)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_CONFIGURAZIONE_DOCUMENTO()


If Not (ObjDoc Is Nothing) Then
    Set ObjDoc = Nothing
End If
Set ObjDoc = New DmtDocs.cDocument
Set ObjDoc.Connection = TheApp.Database.Connection
ObjDoc.SetTipoOggetto 2
ObjDoc.IDFunzione = 105
ObjDoc.TablesNames ObjDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
ObjDoc.IDAzienda = TheApp.IDFirm
ObjDoc.IDFiliale = TheApp.Branch
ObjDoc.IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.IDFirm, TheApp.Branch)
ObjDoc.IDTipoAnagrafica = 2
ObjDoc.IDUtente = TheApp.IDUser
ObjDoc.DataEmissione = Date


End Sub

Private Sub GET_INTESTAZIONE_DOCUMENTO(IDAnagrafica As Long, IDListino As Long, IDListinoAzienda As Long)
ObjDoc.ClearValues

 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica

ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal



End Sub
Private Function AGGIORNA_TIPO_PEDANA(IDPedana As Long, IDTipoPedana As Long, PesoPedana As Double)
On Error GoTo ERR_AGGIORNA_TIPO_PEDANA
Dim sSQL As String

sSQL = "UPDATE RV_POPedana SET "
sSQL = sSQL & "IDRV_POTipoPedana=" & IDTipoPedana & ", "
sSQL = sSQL & "PesoPedana=" & fnNormNumber(PesoPedana)
sSQL = sSQL & " WHERE IDRV_POPedana=" & IDPedana


CnDMT.Execute sSQL

Exit Function
ERR_AGGIORNA_TIPO_PEDANA:
    MsgBox Err.Description, vbCritical, "AGGIORNA_TIPO_PEDANA"
End Function

Private Sub txtIDRigaOrdine_Change()
On Error GoTo ERR_txtIDRigaOrdine_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If Me.txtIDRigaOrdine.Value = 0 Then Exit Sub

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & Me.txtIDRigaOrdine.Value
sSQL = sSQL & "  AND RV_POTipoRiga=1"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtImportoUnitarioArticolo.Value = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
    Me.txtImportoUnitarioListino.Value = fnNotNullN(rs!RV_POImportoUnitarioListino)
    Me.txtSconto1.Value = fnNotNullN(rs!Art_sco_in_percentuale_1)
    Me.txtSconto2.Value = fnNotNullN(rs!Art_sco_in_percentuale_2)
    Me.chkPrezzoInclusoImballo.Value = fnNotNullN(rs!RV_POImportoImballoInArticolo)
    If (Me.CDCodiceImballo.KeyFieldID = fnNotNullN(rs!RV_POIDImballo)) Then
        Me.txtImportoUnitarioImballo.Value = GET_IMPORTO_IMBALLO_DA_RIGA_ORDINE(Me.txtIDOrdinePadre.Value, fnNotNullN(rs!RV_POLinkRiga))
    End If
    Me.txtRaggrOrd.Text = fnNotNull(rs!RV_PONotaRigaOrdRaggr)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_txtIDRigaOrdine_Change:
    MsgBox Err.Description, vbCritical, "txtIDRigaOrdine_Change"
End Sub

Private Function GET_IMPORTO_IMBALLO_DA_RIGA_ORDINE(IDOggetto As Long, linkRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IMPORTO_IMBALLO_DA_RIGA_ORDINE = 0


sSQL = "SELECT IDValoriOggettoDettaglio, Art_prezzo_unitario_netto_IVA"
sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
sSQL = sSQL & " AND RV_POTipoRiga=2"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_IMPORTO_IMBALLO_DA_RIGA_ORDINE = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(IDArticolo As Long, IDImballo As Long, IDOggettoOrdine As Long) As Boolean
On Error GoTo ERR_GET_CONFIGURAZIONE_PREZZO_DA_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDArticoloPadre As Long
Dim NumeroCombinazioni As Long
Dim Result As Boolean
Dim AttivaCursore As Boolean

    Result = False
    RETURN_SEL_PREZZO_IMB_DA_ORD = 0
    AttivaCursore = False
    
    IDArticoloPadre = IDArticolo
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        If TROVA_PREZZI_NO_IMB = 0 Then
            sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
        End If
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            If TROVA_PREZZI_NO_IMB = 0 Then
                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            End If
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                Me.txtImportoUnitarioArticolo.Value = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                Me.txtImportoUnitarioListino.Value = fnNotNullN(rs!RV_POImportoUnitarioListino)
                Me.txtSconto1.Value = fnNotNullN(rs!Art_sco_in_percentuale_1)
                Me.txtSconto2.Value = fnNotNullN(rs!Art_sco_in_percentuale_2)
                Me.txtRaggrOrd.Text = fnNotNull(rs!RV_PONotaRigaOrdRaggr)
                
                RECUPERO_PREZZI_DA_ORD = True
                Me.txtIDRigaOrdine.Value = fnNotNullN(rs!IDValoriOggettoDettaglio)
                RECUPERO_PREZZI_DA_ORD = False
                
                If (IDImballo = fnNotNullN(rs!RV_POIDImballo)) Then
                    If fnNotNullN(rs!RV_POLinkRiga) > 0 Then
                        RETURN_SEL_PREZZO_IMB_DA_ORD = 1
                        GET_PREZZO_IMBALLO_DA_ORDINE IDImballo, fnNotNullN(rs!RV_POLinkRiga), IDOggettoOrdine
                    End If
                Else
                    RETURN_SEL_PREZZO_IMB_DA_ORD = 0
                End If
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            
            Result = True
            
        End If
        If NumeroCombinazioni > 1 Then
            If Screen.MousePointer = 11 Then
                Screen.MousePointer = 0
                AttivaCursore = True
            End If
            
            LINK_ARTICOLO_ORDINE = Me.CDArticolo.KeyFieldID
            LINK_ORDINE_PER_PREZZO = IDOggettoOrdine
            MODALITA_RECUPERO_RIGA_ORD = 1
            
            frmCorpoOrdine.Show vbModal
            
            If CONFERMA_SEL_PREZZO_DA_ORD = 1 Then
                Result = True
            End If
            If AttivaCursore = True Then
                Screen.MousePointer = 11
                DoEvents
            End If
        End If
        If NumeroCombinazioni = 0 Then
            If VIS_ELECO_RIGHE_ORD = 1 Then
                If Screen.MousePointer = 11 Then
                    Screen.MousePointer = 0
                    AttivaCursore = True
                End If
                
                LINK_ARTICOLO_ORDINE = 0
                LINK_ORDINE_PER_PREZZO = IDOggettoOrdine
                MODALITA_RECUPERO_RIGA_ORD = 1
                
                frmCorpoOrdine.Show vbModal
                
                If CONFERMA_SEL_PREZZO_DA_ORD = 1 Then
                    Result = True
                End If
                If AttivaCursore = True Then
                    Screen.MousePointer = 11
                    DoEvents
                End If
            End If
        End If
    End If

    GET_CONFIGURAZIONE_PREZZO_DA_ORDINE = Result

Exit Function
ERR_GET_CONFIGURAZIONE_PREZZO_DA_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_CONFIGURAZIONE_PREZZO_DA_ORDINE"
End Function

Private Sub GET_PREZZO_IMBALLO_DA_ORDINE(IDImballo As Long, linkRiga As Long, IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=2 "
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtImportoUnitarioImballo.Value = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
    Me.chkPrezzoInclusoImballo.Value = Abs(fnNotNullN(rs!RV_POImportoImballoInArticolo))
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_CONTROLLO_MERCE_IN_ORDINE(IDOggettoOrdine As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

GET_CONTROLLO_MERCE_IN_ORDINE = False

sSQL = "SELECT COUNT(IDValoriOggettoDettaglio) AS NumeroRecordOrdine "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecordOrdine)
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord > 0 Then
    GET_CONTROLLO_MERCE_IN_ORDINE = True
End If

End Function
Private Function GET_CALCOLA_TARA_CONFEZIONE()

GET_CALCOLA_TARA_CONFEZIONE = txtColli.Value * (Me.txtNumeroConfImballo.Value * txtTaraConfImballo.Value)

End Function
Private Sub CREA_RECORDSET_LOTTI_IMBALLI(IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long)
On Error GoTo ERR_CREA_RECORDSET_LOTTI_IMBALLI
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsVista As ADODB.Recordset
Dim I As Long
Dim QuantitaLottoProcesso As Double
    
    If Not (rsLottoImballo Is Nothing) Then
        If rsLottoImballo.State > 0 Then
            rsLottoImballo.Close
        End If
        Set rsLottoImballo = Nothing
    End If
    
    Set rsLottoImballo = New ADODB.Recordset
    rsLottoImballo.CursorLocation = adUseClient

    rsLottoImballo.Fields.Append "IDLottoImballo", adInteger, , adFldIsNullable
    rsLottoImballo.Fields.Append "CodiceLottoImballo", adVarChar, 250, adFldIsNullable
    rsLottoImballo.Fields.Append "QuantitaMovimentata", adDouble, , adFldIsNullable
    rsLottoImballo.Fields.Append "IDArticoloImballo", adInteger, , adFldIsNullable
    
    rsLottoImballo.Open , , adOpenKeyset, adLockBatchOptimistic

    sSQL = "SELECT  IDMovimento,RV_POIDLottoImballo, LottoImballo , QuantitaTotale, IDArticolo "
    sSQL = sSQL & " FROM Movimento "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
    'sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
    sSQL = sSQL & " AND RV_POIDLottoImballo>0"
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection
    
    While Not rs.EOF
        rsLottoImballo.AddNew
            rsLottoImballo!IDLottoImballo = fnNotNullN(rs!RV_POIDLottoImballo)
            rsLottoImballo!CodiceLottoImballo = fnNotNull(rs!LottoImballo)
            rsLottoImballo!QuantitaMovimentata = fnNotNullN(rs!QuantitaTotale)
            rsLottoImballo!IDArticoloImballo = fnNotNullN(rs!IDArticolo)
        rsLottoImballo.Update
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Exit Sub
ERR_CREA_RECORDSET_LOTTI_IMBALLI:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_LOTTI_IMBALLI"
End Sub

Private Sub CREA_RECORDSET_LOTTI_IMBALLI_PRIM(IDArticoloImballo As Long, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long)
On Error GoTo ERR_CREA_RECORDSET_LOTTI_IMBALLI
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsVista As ADODB.Recordset
Dim I As Long
Dim QuantitaLottoProcesso As Double

    sSQL = "SELECT * FROM RV_POLottoImballo "
    sSQL = sSQL & "WHERE IDRV_POLottoImballo=0"
    
    Set rsVista = New ADODB.Recordset
    rsVista.Open sSQL, CnDMT.InternalConnection
    
    If Not (rsLottoImballoPrim Is Nothing) Then
        If rsLottoImballoPrim.State > 0 Then
            rsLottoImballoPrim.Close
        End If
        Set rsLottoImballoPrim = Nothing
    End If
    
    Set rsLottoImballoPrim = New ADODB.Recordset
    rsLottoImballoPrim.CursorLocation = adUseClient
    
    With rsVista
        For I = 0 To rsVista.Fields.Count - 1
            Select Case rsVista.Fields(I).Type
                Case adChar, adVarChar, adVarWChar, adWChar, 201
                    rsLottoImballoPrim.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adInteger, adLongVarBinary, adSingle
                    rsLottoImballoPrim.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                Case adDate, adDBTimeStamp, adDBDate
                    rsLottoImballoPrim.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                Case adSmallInt, adBoolean
                    rsLottoImballoPrim.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                Case Else
                    rsLottoImballoPrim.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            End Select
        Next
        rsLottoImballoPrim.Fields.Append "QuantitaSelezionata", adDouble, , adFldIsNullable
        rsLottoImballoPrim.Fields.Append "Rimanenza", adDouble, , adFldIsNullable
        rsLottoImballoPrim.Fields.Append "Registra", adSmallInt, , adFldIsNullable
        rsLottoImballoPrim.Fields.Append "RegistraDaUtente", adSmallInt, , adFldIsNullable
        rsLottoImballoPrim.Fields.Append "QuantitaMovimentata", adDouble, , adFldIsNullable
    End With
    
    
    rsVista.Close
    Set rsVista = Nothing
    
    rsLottoImballoPrim.Open , , adOpenKeyset, adLockBatchOptimistic

    sSQL = "SELECT * FROM RV_POLottoImballo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    'sSQL = sSQL & " AND Giacenza>0"
    
    sSQL = sSQL & "ORDER BY DataCreazione, NumeroProgressivo"
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection
    
    While Not rs.EOF
        rsLottoImballoPrim.AddNew
            For I = 0 To rs.Fields.Count - 1
                rsLottoImballoPrim.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
            
            QuantitaLottoProcesso = GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(fnNotNullN(rs!IDRV_POLottoImballo), IDTipoOggetto, IDOggetto, IDValoriOggettoDettaglio)
            rsLottoImballoPrim!Giacenza = rsLottoImballoPrim!Giacenza + QuantitaLottoProcesso
            rsLottoImballoPrim!QuantitaSelezionata = QuantitaLottoProcesso
            rsLottoImballoPrim!Rimanenza = rsLottoImballoPrim!Giacenza - rsLottoImballoPrim!QuantitaSelezionata
            rsLottoImballoPrim!QuantitaMovimentata = 0
            
            If QuantitaLottoProcesso = 0 Then
                rsLottoImballoPrim!Registra = 0
                rsLottoImballoPrim!RegistraDaUtente = 0
            Else
                rsLottoImballoPrim!Registra = 1
                rsLottoImballoPrim!RegistraDaUtente = 1
            End If
            
            
        rsLottoImballoPrim.Update
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Exit Sub
ERR_CREA_RECORDSET_LOTTI_IMBALLI:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_LOTTI_IMBALLI"
End Sub


Private Function GET_NUMERO_PROCESSO(IDProcesso As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POProcessoIVGamma "
sSQL = sSQL & "WHERE IDRV_POProcessoIVGamma=" & IDProcesso


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_PROCESSO = ""
Else
    GET_NUMERO_PROCESSO = fnNotNullN(rs!AnnoProcesso) & "-" & fnNotNullN(rs!NumeroProcesso)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(IDLottoImballo As Long, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long) As Double
On Error GoTo ERR_GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT  SUM(QuantitaTotale) AS QuantitaTotale "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
'sSQL = sSQL & " AND IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
sSQL = sSQL & " AND RV_POIDLottoImballo>0"
If IDLottoImballo > 0 Then
    sSQL = sSQL & " AND RV_POIDLottoImballo=" & IDLottoImballo
End If
Set rs = CnDMT.OpenResultset(sSQL)

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
Private Sub CREA_RECORDSET_KIT(IDArticoloMerce, IDArticoloImballo As Long, IDArticoloConfezione, IDLavorazione As Long, IDLavorazioneKIT As Long)
On Error GoTo ERR_CREA_RECORDSET_LOTTI_IMBALLI
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsVista As ADODB.Recordset
Dim I As Long
Dim QuantitaLottoProcesso As Double
Dim EsistenzaKit As Long
Dim rsLav As DmtOleDbLib.adoResultset
Dim Colli As Double
Dim PesoLordo As Double
Dim PesoNetto As Double
Dim Tara As Double
Dim Pezzi As Double

'VALORI DALLA LAVORAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Colli = 0
    PesoLordo = 0
    PesoNetto = 0
    Tara = 0
    Pezzi = 0
    
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione
    
    Set rsLav = CnDMT.OpenResultset(sSQL)
    
    If Not rsLav.EOF Then
        Colli = fnNotNullN(rsLav!Colli)
        PesoLordo = fnNotNullN(rsLav!PesoLordo)
        PesoNetto = fnNotNullN(rsLav!PesoNetto)
        Tara = fnNotNullN(rsLav!Tara)
        Pezzi = fnNotNullN(rsLav!Pezzi)
    End If
    
    rsLav.CloseResultset
    Set rsLav = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'EsistenzaKit = GET_ESISTENZA_KIT_LAVORAZIONE(IDLavorazione)

    sSQL = "SELECT * FROM RV_POIEDistintaBaseRighe "
    sSQL = sSQL & "WHERE IDRV_PODistintaBaseRighe=0"
    
    Set rsVista = New ADODB.Recordset
    rsVista.Open sSQL, CnDMT.InternalConnection
    
    If Not (rsKIT Is Nothing) Then
        If rsKIT.State > 0 Then
            rsKIT.Close
        End If
        Set rsKIT = Nothing
    End If
    
    Set rsKIT = New ADODB.Recordset
    rsKIT.CursorLocation = adUseClient
    
    With rsVista
        For I = 0 To rsVista.Fields.Count - 1
            Select Case rsVista.Fields(I).Type
                Case adChar, adVarChar, adVarWChar, adWChar, 201
                    rsKIT.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adInteger, adLongVarBinary, adSingle
                    rsKIT.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                Case adDate, adDBTimeStamp, adDBDate
                    rsKIT.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                Case adSmallInt, adBoolean
                    rsKIT.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                Case Else
                    rsKIT.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            End Select
        Next
        rsKIT.Fields.Append "Selezionato", adBoolean, , adFldIsNullable
        rsKIT.Fields.Append "QuantitaTotale", adDouble, , adFldIsNullable
        rsKIT.Fields.Append "CostoTotale", adDouble, , adFldIsNullable
        rsKIT.Fields.Append "Annotazioni", adVarChar, 250, adFldIsNullable
    End With
    
    
    rsVista.Close
    Set rsVista = Nothing
    
    rsKIT.Open , , adOpenKeyset, adLockBatchOptimistic
    
    
    '''''''''''''''''''''ARTICOLO DA SCARICARE SEMPRE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM RV_POIEDistintaBaseRighe "
    sSQL = sSQL & "WHERE IDArticoloMerce=" & IDArticoloMerce
    sSQL = sSQL & " AND IDArticoloImballo=0"
    sSQL = sSQL & " AND IDArticoloConfezione=0"
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection
    
    While Not rs.EOF
        rsKIT.AddNew
            For I = 0 To rs.Fields.Count - 1
                rsKIT.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
                    
            'If ((EsistenzaKit = 0) Or (IDLavorazione <= 0)) Then
            '    rsKIT!Selezionato = 1
            'Else
                rsKIT!Selezionato = GET_KIT_SELEZIONATO(IDLavorazioneKIT, fnNotNullN(rs!IDRV_PODistintaBaseRighe))
            'End If
            
            Select Case rs!IDUnitaDiMisuraCoop
                Case 1 'Colli
                    rsKIT!QuantitaTotale = Colli * fnNotNullN(rs!Quantita)
                Case 2 'PesoLordo
                    rsKIT!QuantitaTotale = PesoLordo * fnNotNullN(rs!Quantita)
                Case 3 'PesoNetto
                    rsKIT!QuantitaTotale = PesoNetto * fnNotNullN(rs!Quantita)
                Case 4 'Tara
                    rsKIT!QuantitaTotale = Tara * fnNotNullN(rs!Quantita)
                Case 5 'Pezzi
                    rsKIT!QuantitaTotale = Pezzi * fnNotNullN(rs!Quantita)
            End Select
            
            rsKIT!CostoTotale = rsKIT!QuantitaTotale * fnNotNullN(rs!Costo)
            rsKIT!Annotazioni = "Componente da scaricare sempre"
            
        rsKIT.Update
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''ARTICOLO DA SCARICARE PER IL KIT'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM RV_POIEDistintaBaseRighe "
    sSQL = sSQL & "WHERE IDArticoloMerce=" & IDArticoloMerce
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    sSQL = sSQL & " AND IDArticoloConfezione=" & IDArticoloConfezione
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection
    
    While Not rs.EOF
        rsKIT.AddNew
            
            For I = 0 To rs.Fields.Count - 1
                rsKIT.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
            
            'If ((EsistenzaKit = 0) Or (IDLavorazione <= 0)) Then
                rsKIT!Selezionato = 1
            'Else
            '    rsKIT!Selezionato = GET_KIT_SELEZIONATO(IDLavorazione, fnNotNullN(rs!IDRV_PODistintaBaseRighe))
            'End If
            
            Select Case rs!IDUnitaDiMisuraCoop
                Case 1 'Colli
                    rsKIT!QuantitaTotale = Colli * fnNotNullN(rs!Quantita)
                Case 2 'PesoLordo
                    rsKIT!QuantitaTotale = PesoLordo * fnNotNullN(rs!Quantita)
                Case 3 'PesoNetto
                    rsKIT!QuantitaTotale = PesoNetto * fnNotNullN(rs!Quantita)
                Case 4 'Tara
                    rsKIT!QuantitaTotale = Tara * fnNotNullN(rs!Quantita)
                Case 5 'Pezzi
                    rsKIT!QuantitaTotale = Pezzi * fnNotNullN(rs!Quantita)
            End Select
            
            rsKIT!CostoTotale = rsKIT!QuantitaTotale * fnNotNullN(rs!Costo)
            rsKIT!Annotazioni = "Componente del KIT"
            
        rsKIT.Update
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Exit Sub
ERR_CREA_RECORDSET_LOTTI_IMBALLI:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_KIT"
End Sub
Private Sub SALVA_KIT(IDLavorazione As Long)
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "DELETE FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione
CnDMT.Execute sSQL

rsKIT.Filter = "Selezionato=" & fnNormBoolean(1)

If ((rsKIT.EOF) And (rsKIT.BOF)) Then Exit Sub

rsKIT.MoveFirst

sSQL = "SELECT * FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsKIT.EOF
    rsNew.AddNew
        rsNew!IDRV_POAssegnazioneMerce = IDLavorazione
        rsNew!IDArticolo = fnNotNullN(rsKIT!IDArticolo)
        rsNew!Quantita = fnNotNullN(rsKIT!QuantitaTotale)
        rsNew!CostoUnitario = fnNotNullN(rsKIT!Costo)
        rsNew!IDRV_PODistintaBaseRighe = fnNotNullN(rsKIT!IDRV_PODistintaBaseRighe)
        rsNew!IDRV_PODistintaBaseRigheConf = fnNotNullN(rsKIT!IDRV_PODistintaBaseRigheConf)
        rsNew!CostoTotaleRiga = fnNotNullN(rsKIT!CostoTotale)
        rsNew!TracciaImballo = 0
    rsNew.Update
rsKIT.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rsKIT.Close
Set rsKIT = Nothing
End Sub


Public Function GeneraMovimentoScaricoKit(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcesso As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String) As Boolean
On Error GoTo ERR_GeneraMovimentoScaricoImballo
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazione

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    QuantitaRimasta = fnNotNullN(rs!Quantita)
            
    mov.DataMovimento = Me.txtDataLavorazione.Text
    mov.FattoreDiConversione = Null
    
    mov.GestioneMatricole = False
    mov.IDEsercizio = fncEsercizio(Me.txtDataLavorazione)
    mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    mov.IDOggetto = IDRigaConferimento
    
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = cboMagazzinoConf.CurrentID
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", LINK_ANAGRAFICA_SOCIO
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", fnNotNullN(rs!IDArticolo) 'IDArticolo
    mov.Field "IDUnitaDiMisura", fnNotNullN(rs!IDUnitaDiMisura)
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", fnNotNull(rs!Articolo)
    mov.Field "QuantitaTotale", QuantitaRimasta 'Me.txtColli.Value
    mov.Field "Importo", 0
    mov.Field "DataDocumento", Me.txtDataLavorazione.Text
    mov.Field "IDTipoMovimento", 1
    If IDProcesso = 0 Then
        mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
    Else
        mov.Field "Oggetto", "Lavorazione merce del " & Me.txtDataLavorazione.Text & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 2)
    End If
    mov.Field "IDTipoMovimento", 1
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    mov.Field "RV_POTipoRiga", 2
    mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
    mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
    mov.Field "RV_POIDProcessoIVGamma", IDProcesso
    mov.Field "RV_POIDAnagraficaSocio", LINK_ANAGRAFICA_SOCIO
    mov.Field "RV_PODataConferimento", DATA_CONFERIMENTO
    mov.Field "RV_PONumeroConferimento", NUMERO_CONFERIMENTO
    mov.Field "RV_POCodiceLotto", CODICE_LOTTO_ENTRATA
    mov.Field "RV_POCodiceLottoCampagna", CODICE_LOTTO_CAMPAGNA
    mov.Field "RV_POCodiceLottoVendita", Me.txtLottoVendita.Text
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    mov.Field "RV_POQuantitaMovimentata", 0
    mov.Field "RV_PONumeroColli", 0
    mov.Field "RV_POPesoLordo", 0
    mov.Field "RV_POPesoNetto", 0
    mov.Field "RV_POTara", 0
    mov.Field "RV_POQuantitaPezzi", 0

    mov.Field "RV_PODataLavorazione", Null
    mov.Field "RV_POIDTipoLavorazione", 0
    mov.Field "RV_POIDCalibro", 0
    mov.Field "RV_POIDTipoCategoria", 0
    mov.Field "RV_POIDTipoLavorazioneConf", 0
    mov.Field "RV_POPrezzoMedioConf", 0
    
    mov.Field "RV_POIDPedana", 0
    mov.Field "RV_POIDTipoPedana", 0
    mov.Field "RV_POCodicePedana", ""
    mov.Field "RV_POPesoPedana", ""
    
    mov.Field "RV_POIDImballoPrim", 0
    mov.Field "RV_POCodiceImballoPrim", ""
    mov.Field "RV_PODescrizioneImballoPrim", ""
    mov.Field "RV_PONumeroConfezioniPerImballo", 0
    mov.Field "RV_POTaraConfezioneImballo", 0
    mov.Field "RV_POQuantitaTotaleConfImballo", 0
    mov.Field "RV_POCostoConfezioneImballo", 0
                            
    mov.Field "RV_POIDLottoImballo", 0
    mov.Field "LottoImballo", ""
    
    
    mov.Field "TipoRiga", trcNessuno
    CnDMT.BeginTrans
    GeneraMovimentoScaricoKit = mov.Insert
    CnDMT.CommitTrans
            
rs.MoveNext
Wend


Exit Function
ERR_GeneraMovimentoScaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName & " (GeneraMovimentoScaricoKit)"
    CnDMT.RollbackTrans
End Function
Private Function GET_KIT_SELEZIONATO(IDLavorazione As Long, IDDistintaBaseRighe As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_KIT_SELEZIONATO = True

sSQL = "SELECT ID FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione
sSQL = sSQL & " AND IDRV_PODistintaBaseRighe=" & IDDistintaBaseRighe

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
   GET_KIT_SELEZIONATO = False
End If

rs.CloseResultset
Set rs = Nothing
End Function

