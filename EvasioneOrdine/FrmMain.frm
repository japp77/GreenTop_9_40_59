VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Passaggio ordini in fatturazione - Imputazione prezzi - Smistamento merce "
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20205
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
   ScaleHeight     =   11055
   ScaleWidth      =   20205
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic1 
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
      Height          =   10935
      Left            =   0
      ScaleHeight     =   10905
      ScaleWidth      =   20025
      TabIndex        =   0
      Top             =   0
      Width           =   20055
      Begin VB.PictureBox Picture2 
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
         Height          =   9975
         Left            =   120
         ScaleHeight     =   9945
         ScaleWidth      =   19785
         TabIndex        =   1
         Top             =   840
         Width           =   19815
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   16560
            Picture         =   "FrmMain.frx":1084A
            Style           =   1  'Graphical
            TabIndex        =   152
            ToolTipText     =   "Spezzatura - Ripesatura pedana"
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton cmdRipesatura 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   15240
            Picture         =   "FrmMain.frx":10DD4
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Ripesatura pedana"
            Top             =   4680
            Width           =   855
         End
         Begin VB.Frame FraOrdinePrep 
            Height          =   4335
            Left            =   4080
            TabIndex        =   95
            Top             =   0
            Width           =   15855
            Begin VB.CommandButton cmdSalvaStatoGrigliaOrdine 
               Height          =   255
               Left            =   15240
               Picture         =   "FrmMain.frx":1135E
               Style           =   1  'Graphical
               TabIndex        =   131
               ToolTipText     =   "Salva lo stato delle colonne"
               Top             =   120
               Width           =   375
            End
            Begin VB.Frame FraTotaliOrdine 
               Caption         =   "Totali ordine da preparare"
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
               Height          =   3855
               Left            =   12600
               TabIndex        =   96
               Top             =   360
               Visible         =   0   'False
               Width           =   3015
               Begin VB.TextBox txtImponibileOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   136
                  Top             =   3480
                  Width           =   1695
               End
               Begin VB.TextBox txtTotaleTaraOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   270
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   98
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePesoOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   127
                  Top             =   2600
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePesoNettoOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   102
                  Top             =   2160
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePedaneOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   101
                  Top             =   380
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePesoLordoOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   270
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   99
                  Top             =   1240
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePezziOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   97
                  Top             =   3060
                  Width           =   1695
               End
               Begin VB.TextBox txtTotaleColliOrdPrep 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   270
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   100
                  Top             =   800
                  Width           =   1695
               End
               Begin VB.Label Label8 
                  Caption         =   "Imponibile"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   137
                  Top             =   3480
                  Width           =   975
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   2
                  X1              =   120
                  X2              =   2880
                  Y1              =   3360
                  Y2              =   3360
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   120
                  X2              =   2880
                  Y1              =   2900
                  Y2              =   2900
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   120
                  X2              =   2880
                  Y1              =   2440
                  Y2              =   2440
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   2
                  X1              =   120
                  X2              =   2880
                  Y1              =   1540
                  Y2              =   1540
               End
               Begin VB.Label Label8 
                  Caption         =   "Totale peso"
                  Height          =   255
                  Index           =   13
                  Left            =   120
                  TabIndex        =   126
                  Top             =   2600
                  Width           =   1095
               End
               Begin VB.Label Label8 
                  Caption         =   "Peso Netto"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   108
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   120
                  X2              =   2880
                  Y1              =   680
                  Y2              =   680
               End
               Begin VB.Label Label8 
                  Caption         =   "N° pedane"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   107
                  Top             =   380
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Colli"
                  Height          =   270
                  Index           =   1
                  Left            =   120
                  TabIndex        =   106
                  Top             =   800
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Peso lordo"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   105
                  Top             =   1240
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Tara"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   104
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   120
                  X2              =   2880
                  Y1              =   1100
                  Y2              =   1100
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   120
                  X2              =   2880
                  Y1              =   2000
                  Y2              =   2000
               End
               Begin VB.Label Label8 
                  Caption         =   "Pezzi"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   103
                  Top             =   3060
                  Width           =   975
               End
            End
            Begin DMTLblLinkCtl.LabelLink LabelLink1 
               Height          =   225
               Left            =   120
               TabIndex        =   109
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   397
               Caption         =   "Modifica lavorazione"
               Name            =   "LabelLink"
            End
            Begin DmtGridCtl.DmtGrid GrigliaAssegnazione 
               Height          =   3855
               Left            =   120
               TabIndex        =   110
               Top             =   360
               Width           =   15525
               _ExtentX        =   27384
               _ExtentY        =   6800
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
            Begin DMTLblLinkCtl.LabelLink LabelLink4 
               Height          =   135
               Left            =   0
               TabIndex        =   134
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   238
               Caption         =   "IVGamma"
               Name            =   "LabelLink"
            End
            Begin VB.Label lblInfoOrdine 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   120
               Width           =   15105
            End
         End
         Begin VB.CommandButton cmdAvanti 
            Caption         =   "EMISSIONE DOCUMENTI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2040
            TabIndex        =   93
            Top             =   3840
            Width           =   1935
         End
         Begin VB.TextBox txtAssegnazioneVeloce 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   7440
            TabIndex        =   53
            Top             =   4540
            Width           =   4455
         End
         Begin VB.CommandButton cmdEliminaConferma 
            Caption         =   "ELIMINA CONFERMA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   0
            TabIndex        =   51
            Top             =   3840
            Width           =   1935
         End
         Begin MSComctlLib.ListView LVordiniDaElaborare 
            Height          =   3495
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   6165
            View            =   3
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
         Begin VB.CommandButton cmdSuTutto 
            Caption         =   "Tutto su"
            Enabled         =   0   'False
            Height          =   495
            Left            =   17160
            TabIndex        =   46
            Top             =   3360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdSuSingolo 
            Caption         =   "Singola"
            Height          =   405
            Left            =   13920
            TabIndex        =   45
            ToolTipText     =   "Sposta su"
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton cmdGiuSingolo 
            Caption         =   "Singola"
            Height          =   405
            Left            =   4560
            TabIndex        =   44
            ToolTipText     =   "Sposta giù"
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton cmdGiuTutto 
            Caption         =   "Tutto giù"
            Enabled         =   0   'False
            Height          =   495
            Left            =   16080
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   3360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdSu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12600
            Picture         =   "FrmMain.frx":118E8
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Sposta su con modifiche"
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton cmdGiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5880
            Picture         =   "FrmMain.frx":11E72
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Sposta giù con modifiche"
            Top             =   4680
            Width           =   855
         End
         Begin VB.Frame FraOrdineSmist 
            Caption         =   "MERCE DA SMISTARE"
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
            Height          =   4575
            Left            =   0
            TabIndex        =   2
            Top             =   5280
            Width           =   19695
            Begin VB.TextBox txtRaggrOrdine 
               Height          =   315
               Left            =   5040
               TabIndex        =   153
               Top             =   1080
               Width           =   4815
            End
            Begin VB.CommandButton cmdEliminaLavorazioni 
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
               Left            =   4560
               Picture         =   "FrmMain.frx":123FC
               Style           =   1  'Graphical
               TabIndex        =   135
               TabStop         =   0   'False
               ToolTipText     =   "Lista delle lavorazioni con collii zero"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdSalvaStatoGrigliaSmist 
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
               Left            =   19200
               Picture         =   "FrmMain.frx":12986
               Style           =   1  'Graphical
               TabIndex        =   132
               TabStop         =   0   'False
               ToolTipText     =   "Salva lo stato delle colonne della griglia"
               Top             =   0
               Width           =   375
            End
            Begin VB.TextBox txtLottoVendita 
               Height          =   315
               Left            =   120
               TabIndex        =   124
               Top             =   1080
               Width           =   4815
            End
            Begin VB.CommandButton cmdTotaliOrdSmist 
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
               Left            =   3960
               Picture         =   "FrmMain.frx":12F10
               Style           =   1  'Graphical
               TabIndex        =   94
               TabStop         =   0   'False
               ToolTipText     =   "Visualizza totali ricerca"
               Top             =   0
               Width           =   375
            End
            Begin VB.Frame FraTotaliOrdineSmist 
               Caption         =   "Totali merce da smistare"
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
               Height          =   3495
               Left            =   16560
               TabIndex        =   79
               Top             =   960
               Visible         =   0   'False
               Width           =   3015
               Begin VB.TextBox txtTotalePesoNettoOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   86
                  Top             =   2400
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePedaneOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   85
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.TextBox txtTotaleColliOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   84
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePesoLordoOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   83
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.TextBox txtTotaleTaraOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   82
                  Top             =   1920
                  Width           =   1695
               End
               Begin VB.TextBox txtTotalePezziOrdSmist 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   81
                  Top             =   2880
                  Width           =   1695
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
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
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   80
                  Top             =   2880
                  Width           =   1695
               End
               Begin VB.Label Label8 
                  Caption         =   "Peso Netto"
                  Height          =   255
                  Index           =   11
                  Left            =   120
                  TabIndex        =   92
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   7
                  X1              =   120
                  X2              =   2880
                  Y1              =   840
                  Y2              =   840
               End
               Begin VB.Label Label8 
                  Caption         =   "N° pedane"
                  Height          =   255
                  Index           =   10
                  Left            =   120
                  TabIndex        =   91
                  Top             =   480
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Colli"
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   90
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Peso lordo"
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   89
                  Top             =   1440
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Caption         =   "Tara"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1920
                  Width           =   975
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   6
                  X1              =   120
                  X2              =   2880
                  Y1              =   1320
                  Y2              =   1320
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   5
                  X1              =   120
                  X2              =   2880
                  Y1              =   1800
                  Y2              =   1800
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  Index           =   4
                  X1              =   120
                  X2              =   2880
                  Y1              =   2280
                  Y2              =   2280
               End
               Begin VB.Label Label8 
                  Caption         =   "Pezzi"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   87
                  Top             =   2880
                  Width           =   975
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   120
                  X2              =   2880
                  Y1              =   2760
                  Y2              =   2760
               End
            End
            Begin VB.CommandButton cmdTrovaOrdineSmistamento 
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
               Left            =   2760
               Picture         =   "FrmMain.frx":1349A
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Trova ordine da smistare"
               Top             =   0
               Width           =   375
            End
            Begin DMTEDITNUMLib.dmtNumber txtIDOrdineSmistamento 
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin VB.CommandButton cmdStampaOrdineSmist 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3360
               Picture         =   "FrmMain.frx":13A24
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Stampa ordine da smistare"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdReimpostaOrdDft 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2160
               Picture         =   "FrmMain.frx":13FAE
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Merce in giacenza"
               Top             =   0
               Width           =   375
            End
            Begin VB.TextBox txtCodicePedana 
               Height          =   315
               Left            =   13080
               TabIndex        =   4
               Top             =   480
               Width           =   1905
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
               Height          =   315
               Left            =   15000
               Picture         =   "FrmMain.frx":14538
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Trova pedana"
               Top             =   480
               Width           =   375
            End
            Begin DmtGridCtl.DmtGrid GrigliaSmistamento 
               Height          =   3015
               Left            =   120
               TabIndex        =   5
               Top             =   1440
               Width           =   19455
               _ExtentX        =   34316
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
               UpdatePosition  =   0   'False
               UseUserSettings =   0   'False
               ColumnsHeaderHeight=   20
            End
            Begin DMTEDITNUMLib.dmtNumber txtNumeroSmistamento 
               Height          =   315
               Left            =   6360
               TabIndex        =   6
               Top             =   480
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   253
               ForeColor       =   -2147483630
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DmtCodDescCtl.DmtCodDesc CDAltroCliente 
               Height          =   615
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   1085
               PropCodice      =   $"FrmMain.frx":14AC2
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"FrmMain.frx":14B10
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"FrmMain.frx":14B62
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
            Begin DMTDATETIMELib.dmtDate txtDataSmistamento 
               Height          =   315
               Left            =   5040
               TabIndex        =   8
               Top             =   480
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   253
               ForeColor       =   -2147483630
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
               Height          =   615
               Left            =   8640
               TabIndex        =   9
               Top             =   240
               Width           =   4455
               _ExtentX        =   7858
               _ExtentY        =   1085
               PropCodice      =   $"FrmMain.frx":14BBC
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"FrmMain.frx":14C14
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"FrmMain.frx":14C74
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
            Begin DmtCodDescCtl.DmtCodDesc CDSocio 
               Height          =   615
               Left            =   15480
               TabIndex        =   49
               Top             =   240
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1085
               PropCodice      =   $"FrmMain.frx":14CCE
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"FrmMain.frx":14D1C
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"FrmMain.frx":14D76
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
            Begin DMTLblLinkCtl.LabelLink LabelLink2 
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   1680
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               Caption         =   "Modifica lavorazione"
               Name            =   "LabelLink"
            End
            Begin DMTEDITNUMLib.dmtNumber txtNListaSmistamento 
               Height          =   315
               Left            =   7800
               TabIndex        =   150
               Top             =   480
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   556
               _StockProps     =   253
               ForeColor       =   -2147483630
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin VB.Label Label2 
               Caption         =   "Sub lotto"
               Height          =   255
               Index           =   3
               Left            =   5040
               TabIndex        =   154
               Top             =   840
               Width           =   4815
            End
            Begin VB.Label Label5 
               Caption         =   "N° lista"
               Height          =   255
               Index           =   1
               Left            =   7800
               TabIndex        =   151
               Top             =   285
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Lotto di vendita"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   125
               Top             =   840
               Width           =   4815
            End
            Begin VB.Label Label2 
               Caption         =   "Pedana"
               Height          =   255
               Index           =   1
               Left            =   13080
               TabIndex        =   12
               Top             =   285
               Width           =   1815
            End
            Begin VB.Label Label3 
               Caption         =   "Numero ordine"
               Height          =   255
               Left            =   6360
               TabIndex        =   11
               Top             =   285
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Data ordine"
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   10
               Top             =   280
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkAttitaMaschera 
            Caption         =   "Proponi sempre le maschere"
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
            Left            =   7440
            TabIndex        =   129
            Top             =   5040
            Width           =   4455
         End
         Begin DMTLblLinkCtl.LabelLink LabelLink3 
            Height          =   135
            Left            =   360
            TabIndex        =   133
            Top             =   4800
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   238
            Caption         =   "IVGamma"
            Name            =   "LabelLink"
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   480
            Index           =   0
            Left            =   1920
            Top             =   3840
            Width           =   135
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Assegnazione veloce"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   7440
            TabIndex        =   52
            Top             =   4300
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ORDINI CONFERMATI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Top             =   90
            Width           =   3975
         End
      End
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
         Left            =   12480
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   6255
         Begin VB.OptionButton optSceltaStampa 
            Caption         =   "Ordine da smistare"
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   70
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton optSceltaStampa 
            Caption         =   "Ordine da preparare"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton cmdStampa 
            Height          =   375
            Left            =   5760
            Picture         =   "FrmMain.frx":14DD0
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "STAMPA"
            Top             =   160
            Width           =   375
         End
         Begin DmtActBox.DmtActBoxCtl ActivityBox 
            Height          =   9255
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   16325
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
      Begin VB.Frame Frame3 
         Caption         =   "Indicare l'ordine da preparare"
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   19815
         Begin VB.TextBox txtNLetteraIntento 
            Height          =   315
            Left            =   9480
            Locked          =   -1  'True
            TabIndex        =   155
            Top             =   3240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdAggiornaOrdinePadre 
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
            Left            =   3720
            Picture         =   "FrmMain.frx":1535A
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Aggiorna ordine padre"
            Top             =   0
            Width           =   375
         End
         Begin DmtCodDescCtl.DmtCodDesc cdCliente 
            Height          =   615
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":158E4
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":15932
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":15984
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
         Begin VB.TextBox txtNumeroOrdineOri 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   141
            Top             =   6075
            Width           =   1455
         End
         Begin VB.CommandButton cmdNListaPrelievo 
            Caption         =   "NUOVA LISTA PRELIEVO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11880
            TabIndex        =   138
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdRipPedOrd 
            Caption         =   "RIPESATURA PEDANA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   16680
            TabIndex        =   130
            Top             =   240
            Width           =   1455
         End
         Begin DMTLblLinkCtl.LabelLink lblLinkOrdine 
            Height          =   135
            Left            =   8280
            TabIndex        =   123
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   238
            Caption         =   "Nuovo ordine"
            Name            =   "LabelLink"
         End
         Begin VB.CommandButton cmdEliminaRifLetInt 
            Height          =   315
            Left            =   8760
            Picture         =   "FrmMain.frx":159DE
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Elimina riferimento lettera intento"
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdLetteraIntento 
            Height          =   315
            Left            =   9120
            Picture         =   "FrmMain.frx":15F68
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Lettere di intento del cliente"
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdTotaliOrdine 
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
            Left            =   5160
            Picture         =   "FrmMain.frx":164F2
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Visualizza totali ordine"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdRiepilogoOridine 
            Caption         =   "RIEPILOGO ORDINE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   18240
            TabIndex        =   77
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDescrizioneRigaDoc 
            Height          =   765
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   4440
            Width           =   8535
         End
         Begin VB.CommandButton cmdStampaReportAut 
            Height          =   315
            Left            =   4200
            Picture         =   "FrmMain.frx":16A7C
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Stampa ordine"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdApriStampa 
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
            Left            =   4680
            Picture         =   "FrmMain.frx":17006
            Style           =   1  'Graphical
            TabIndex        =   71
            TabStop         =   0   'False
            ToolTipText     =   "Apri scelta report"
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtAnnotazioniInterna 
            Height          =   765
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   3360
            Width           =   8535
         End
         Begin VB.TextBox txtNumeroOrdineCliente 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtAnnotazioniOrdine 
            Height          =   765
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   2280
            Width           =   8535
         End
         Begin VB.CommandButton cmdAggiornaOrdineCliente 
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
            Left            =   3240
            Picture         =   "FrmMain.frx":17590
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Aggiorna ordine cliente"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdApriFrameOrdine 
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
            Left            =   2760
            Picture         =   "FrmMain.frx":17B1A
            Style           =   1  'Graphical
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Ricalcola dati cliente"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdImputazionePrezziVeloce 
            Caption         =   "IMPUTAZIONE PREZZI VELOCE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   15120
            TabIndex        =   54
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdNuovoOrdine 
            Caption         =   "NUOVO ORDINE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13560
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdRicerca 
            Caption         =   "RICERCA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8760
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdTrasporta 
            Caption         =   "CONFERMA ORDINE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10320
            TabIndex        =   18
            Top             =   240
            Width           =   1455
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
            Height          =   315
            Left            =   6480
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   480
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDataOrdine 
            Height          =   315
            Left            =   5160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTEDITNUMLib.dmtNumber txtIDOrdine 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboDestinazione 
            Height          =   315
            Left            =   2880
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTDataCmb.DMTCombo cboVettore 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
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
         Begin DMTDATETIMELib.dmtDate txtDataPartenza 
            Height          =   315
            Left            =   4200
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDataOrdineCliente 
            Height          =   315
            Left            =   1080
            TabIndex        =   29
            Top             =   1080
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   5520
            Width           =   1935
            _ExtentX        =   3413
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
         Begin DMTDataCmb.DMTCombo cboVettoreSuccessivo 
            Height          =   315
            Left            =   4560
            TabIndex        =   38
            Top             =   5520
            Width           =   4095
            _ExtentX        =   7223
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
         Begin DMTDATETIMELib.dmtDate txtDataArrivoMerceL 
            Height          =   315
            Left            =   2160
            TabIndex        =   36
            Top             =   5520
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDataArrivoMerce 
            Height          =   315
            Left            =   6360
            TabIndex        =   25
            Top             =   1680
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtTime txtOraArrivoMerce 
            Height          =   315
            Left            =   7680
            TabIndex        =   26
            Top             =   1680
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboTipoTrasporto 
            Height          =   315
            Left            =   5760
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1080
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
         Begin DMTDATETIMELib.dmtTime txtOraArrivoMerceL 
            Height          =   315
            Left            =   3480
            TabIndex        =   37
            Top             =   5520
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
            Height          =   255
            Left            =   8760
            TabIndex        =   119
            Top             =   3000
            Visible         =   0   'False
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   450
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Appearance      =   1
            AllowEmpty      =   0   'False
         End
         Begin DMTDATETIMELib.dmtDate txtDataLetteraIntento 
            Height          =   315
            Left            =   9960
            TabIndex        =   120
            Top             =   3240
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboIvaCliente 
            Height          =   315
            Left            =   11280
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   3195
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTEDITNUMLib.dmtNumber txtNListaPrelievo 
            Height          =   315
            Left            =   7920
            TabIndex        =   139
            Top             =   480
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTEDITNUMLib.dmtNumber txtIDOrdinePadre 
            Height          =   315
            Left            =   120
            TabIndex        =   142
            Top             =   6075
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            ForeColor       =   -2147483630
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
         Begin DMTDATETIMELib.dmtDate txtDataOrdineOri 
            Height          =   315
            Left            =   5760
            TabIndex        =   143
            Top             =   6075
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   -2147483630
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboSezionaleOrdOri 
            Height          =   315
            Left            =   2040
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   6075
            Width           =   3615
            _ExtentX        =   6376
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
            Caption         =   "Sezionale"
            Height          =   255
            Index           =   15
            Left            =   2040
            TabIndex        =   148
            Top             =   5880
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Ordine padre"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   147
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Data ord. ori."
            Height          =   255
            Index           =   6
            Left            =   5760
            TabIndex        =   146
            Top             =   5880
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Num. ord. ori."
            Height          =   255
            Index           =   7
            Left            =   7200
            TabIndex        =   145
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "N° lista"
            Height          =   255
            Index           =   6
            Left            =   7920
            TabIndex        =   140
            Top             =   280
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   18180
            X2              =   18180
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   16620
            X2              =   16620
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Label lblLetteraIntento 
            Caption         =   "Lettera d'intento"
            Height          =   255
            Left            =   9480
            TabIndex        =   121
            Top             =   3000
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Data arrivo"
            Height          =   255
            Index           =   12
            Left            =   6360
            TabIndex        =   116
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Ora arrivo"
            Height          =   255
            Index           =   20
            Left            =   7680
            TabIndex        =   115
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo trasporto"
            Height          =   255
            Index           =   13
            Left            =   5760
            TabIndex        =   114
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Data arrivo"
            Height          =   255
            Index           =   14
            Left            =   2160
            TabIndex        =   113
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Ora arrivo"
            Height          =   255
            Index           =   21
            Left            =   3480
            TabIndex        =   112
            Top             =   5280
            Width           =   975
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   8680
            X2              =   8680
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   15060
            X2              =   15060
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   11820
            X2              =   11820
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Label Label4 
            Caption         =   "Luogo di presa merce"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   76
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label Label4 
            Caption         =   "Vettore successivo"
            Height          =   255
            Index           =   11
            Left            =   4560
            TabIndex        =   75
            Top             =   5280
            Width           =   2535
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazioni finali del corpo del documento di evasione"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   74
            Top             =   4200
            Width           =   6975
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazioni interne"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   65
            Top             =   3120
            Width           =   7095
         End
         Begin VB.Label Label4 
            Caption         =   "Data ordine cli."
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   63
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Num. ordine cli."
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   62
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Data partenza"
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   60
            Top             =   840
            Width           =   1335
         End
         Begin VB.Line Line2 
            X1              =   8680
            X2              =   19680
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label4 
            Caption         =   "Destinazione diversa"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   59
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Vettore"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   13500
            X2              =   13500
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   10260
            X2              =   10260
            Y1              =   120
            Y2              =   840
         End
         Begin VB.Label Label5 
            Caption         =   "Numero"
            Height          =   255
            Index           =   0
            Left            =   6480
            TabIndex        =   40
            Top             =   280
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Data"
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   39
            Top             =   280
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazioni di fatturazione"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   2040
            Width           =   6975
         End
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   14520
      TabIndex        =   48
      Top             =   4320
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   14520
      TabIndex        =   47
      Top             =   3600
      Width           =   255
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
Public rsGrigliaAss As ADODB.Recordset
Public rsGrigliaOrd As DmtOleDbLib.adoResultset
Public rsGrigliaSmistamento As DmtOleDbLib.adoResultset
Private gPaintNotify As PaintNotify
Public oReport As dmtReportLib.dmtReport
Public lngTimerID As Long

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
'***Supporto tecndmtdmtico                                                       -
Private oSupportActivity As DmtActBoxLib.SupportActivity                  '-
'***Nome dell'attività predefinita del riquadro attività                   -
Private m_DefaultActivity As String                                       '-
'---------------------------------------------------------------------------

'VARIABILI PER LA CONFIGURAZIONE DELL'OGGETTO DOCUMENTO PER PRELEVARE I PREZZI
Private ObjDoc As DmtDocs.cDocument
Private sTabellaTestataLocal As String
Private sTabellaDettaglioLocal As String
Private sTabellaIVALocal As String
Private sTabellaScadenzeLocal As String
Private Link_Pedana As Long

'Variabili recordset per le visualizzazioni delle griglie

Public Sub ConnessioneADO()
On Error GoTo ERR_ConnessioneADO
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

    SettaggioIniziale
  
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"

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
    
    Me.optSceltaStampa(0).Value = True

    'Inizializzazione della LabelLink
    '------------------------------------------------------------------------------------------------
    
    'Lavorazione da ordine
    Set LabelLink1.Application = TheApp    'L’oggetto Application
    LabelLink1.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POAssegnazioneMerce"))
    LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False
    LabelLink1.DisableDoEvents = True
    
    'Lavorazione da smistamento
    Set LabelLink2.Application = TheApp    'L’oggetto Application
    LabelLink2.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    LabelLink2.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POAssegnazioneMerce"))
    LabelLink2.PopMenuItems("Mnu_SearchObject").Enabled = False
    LabelLink2.DisableDoEvents = True
    
    'Processo di IV° Gamma
    Set LabelLink3.Application = TheApp    'L’oggetto Application
    LabelLink3.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    LabelLink3.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POIVGamma"))
    LabelLink3.PopMenuItems("Mnu_SearchObject").Enabled = False
    LabelLink3.DisableDoEvents = True
    
    'Processo di IV° Gamma
    Set LabelLink4.Application = TheApp    'L’oggetto Application
    LabelLink4.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    LabelLink4.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POIVGamma"))
    LabelLink4.PopMenuItems("Mnu_SearchObject").Enabled = False
    LabelLink4.DisableDoEvents = True
    
    Set lblLinkOrdine.Application = TheApp    'L’oggetto Application
    lblLinkOrdine.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    lblLinkOrdine.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POOrdineL"))
    lblLinkOrdine.PopMenuItems("Mnu_SearchObject").Enabled = False
    lblLinkOrdine.DisableDoEvents = True
    
    With Me.cboVettoreSuccessivo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .Sql = "SELECT * FROM Vettore ORDER BY Vettore"
        .Fill
    End With

    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica  "
        .Sql = .Sql & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
    
   
    Me.lblInfoOrdine.BackStyle = 0
    
    GET_CONFIGURAZIONE_DOCUMENTO
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
Exit Sub
ERR_ConnessioneADO:
    MsgBox Err.Description, vbCritical, "ConnessioneADO"
End Sub
Public Sub SettaggioIniziale()
Dim sSQL As String
    
    sSQL = "DELETE FROM RV_POTMPEvasioneOrdini "
    sSQL = sSQL & "WHERE InFatturazione=" & fnNormBoolean(1)
    CnDMT.Execute sSQL
    
    GET_PARAMETRI_EVASIONE_ORDINI
    
    GET_CLIENTE_ORDINE_PRED
    GET_CLIENTE_ORDINE_PREP
    
    ParametroGestioneOrdineVivaio
    
    fnGrigliaAssegnazione
    fnGrigliaOrdiniDaElaborare
    fncGrigliaSmistamento
    GET_PARAMETRI_FILIALI_PER_TOTALI
    GET_PREZZI_DA_ORDINE
    
    Me.chkAttitaMaschera.Value = fnGetParametriMagazzino("ProponiMaschereAvvioVeloce")
End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property

Private Sub CDAltroCliente_ChangeElement()
If Me.CDAltroCliente.KeyFieldID <> LINK_CLIENTE_ORD_PRED Then
    
End If

fncGrigliaSmistamento
End Sub


Private Sub CDArticolo_ChangeElement()
    fncGrigliaSmistamento
End Sub

Private Sub cdCliente_ChangeElement()
If Me.cdCliente.KeyFieldID = 0 Then
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
Else
    
End If
End Sub


Private Sub CDSocio_ChangeElement()
    fncGrigliaSmistamento
End Sub

Private Sub cmdAggiornaOrdineCliente_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsConf As DmtOleDbLib.adoResultset

If Me.txtIDOrdine.Value > 0 Then
    Me.lblLinkOrdine.IDReturn = Me.txtIDOrdine.Value
    Me.lblLinkOrdine.RunApplication
    
    txtIDOrdine_Change
End If


If Me.txtIDOrdine.Value = 0 Then Exit Sub

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE STANDARD''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & 15
sSQL = sSQL & "  AND IDFunzione=" & 128

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    MsgBox "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente), vbCritical, "Impossibile aggiornare ordine"

    rsConf.CloseResultset
    Set rsConf = Nothing
    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE GREEN TOP'''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POOrdineL")
sSQL = sSQL & "  AND IDFunzione=" & GET_FUNZIONE(fnGetTipoOggetto("RV_POOrdineL"))

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    MsgBox "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente), vbCritical, "Impossibile aggiornare ordine"
    rsConf.CloseResultset
    Set rsConf = Nothing
    
    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




If CHECK_ABILITAZIONE_DIAMANTE = False Then Exit Sub
    
    
    
AGGIORNA_ORDINE_CLIENTE Me.cboDestinazione.CurrentID, Me.cboVettore.CurrentID, Me.cboTipoTrasporto.CurrentID
End Sub



Public Sub InitControlli()
     With Me.cdCliente
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIETipoAnagraficaCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Clienti"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Anagrafica"
        .CodeIsNumeric = False
    End With

    With Me.CDAltroCliente
       Set .Application = TheApp
       Set .Database = TheApp.Database
       .HwndContainer = Me.hwnd
       .CodeField = "Codice"
       .DescriptionField = "Anagrafica"
       .KeyField = "IDAnagrafica"
       .TableName = "RV_POIETipoAnagraficaCliente"
       .Filter = "IDAzienda = " & TheApp.IDFirm
       .PropCodice.Caption = "Codice"
       .PropDescrizione.Caption = "Clienti"
       .CodeCaption4Find = "Codice"
       .DescriptionCaption4Find = "Anagrafica"
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
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Socio\Fornitore"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Socio\Fornitore"
        .CodeIsNumeric = False
    End With
    
    
     With Me.CDArticolo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .CodeIsNumeric = False
    End With

    With Me.cboVettore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .Sql = "SELECT * FROM Vettore ORDER BY Vettore"
        .Fill
    End With

    With Me.cboTipoTrasporto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDTipoSpedizione"
        .DisplayField = "TipoSpedizione"
        .Sql = "SELECT * FROM TipoSpedizione ORDER BY TipoSpedizione"
        .Fill
    End With

    With Me.cboSezionaleOrdOri
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .Sql = "SELECT  Sezionale.IDSezionale, Sezionale.Sezionale, RegistroIvaPerTipoOggetto.IDFiliale "
        .Sql = .Sql & "FROM Sezionale INNER JOIN "
        .Sql = .Sql & "RegistroIvaPerTipoOggetto ON Sezionale.IDRegistroIva = RegistroIvaPerTipoOggetto.IDRegistroIva AND "
        .Sql = .Sql & "Sezionale.IDFiliale = RegistroIvaPerTipoOggetto.IDFiliale "
        .Sql = .Sql & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & 15
        .Sql = .Sql & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
    End With
    
    'Inizializza la ListView contenente la ricerca
    With Me.LVordiniDaElaborare
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "N°", 20
        .ColumnHeaders.Add , , "N° ord.", 1000
        .ColumnHeaders.Add , , "Data ord.", 1000
        .ColumnHeaders.Add , , "N°", 25
        .ColumnHeaders.Add , , "Cliente", 1500
        .ColumnHeaders.Add , , "Destinazione", 1500
        .ColumnHeaders.Add , , "Vettore", 1500
    End With

    With Me.cboIvaCliente
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .Sql = "SELECT IDIva, Iva FROM Iva"
        .Sql = .Sql & " ORDER BY Codice"
    End With
End Sub

Private Sub cmdAggiornaOrdinePadre_Click()

If Me.txtIDOrdinePadre.Value > 0 Then
    
    NUOVO_ORDINE_LISTA_PRELIEVO = False
    
    Me.lblLinkOrdine.DisableDoEvents = True
    Me.lblLinkOrdine.IDReturn = Me.txtIDOrdinePadre.Value
    Me.lblLinkOrdine.RunApplication
   
End If


End Sub

Private Sub cmdApriFrameOrdine_Click()
    'If Me.txtIDOrdine.Value > 0 Then
    If Me.Frame3.Height = 855 Then
        Me.Frame3.ZOrder 0
        Me.Frame3.Height = 6615

    Else
        Me.Frame3.Height = 855
        Me.Frame3.ZOrder 1
    End If
    'End If
End Sub

Private Sub cmdApriStampa_Click()
If Me.FraStampa.Visible = False Then
    Me.FraStampa.Visible = True
    Me.FraStampa.ZOrder 0
    Me.FraStampa.Height = 9975
Else
    Me.FraStampa.Visible = False
End If
End Sub

Private Sub cmdAvanti_Click()

    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If

    Unload Me
End Sub

Private Sub cmdEliminaConferma_Click()
Dim IDOggettoOrdineDelete As Long
Dim sSQL As String
If Me.LVordiniDaElaborare.ListItems.Count = 0 Then Exit Sub
IDOggettoOrdineDelete = Me.LVordiniDaElaborare.ListItems(Me.LVordiniDaElaborare.SelectedItem.Index)

If IDOggettoOrdineDelete > 0 Then
    sSQL = "DELETE FROM RV_POTMPEvasioneOrdini "
    sSQL = sSQL & "WHERE NumeroRiga=" & IDOggettoOrdineDelete
    
    CnDMT.Execute sSQL
End If

fnGrigliaOrdiniDaElaborare

End Sub

Private Sub cmdEliminaLavorazioni_Click()
On Error GoTo ERR_cmdEliminaLavorazioni_Click
Dim Testo As String
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Integer


If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
    
    Exit Sub
End If

If Not (rsLavorazioni Is Nothing) Then
    If rsLavorazioni.State > 0 Then
        rsLavorazioni.Close
    End If
    Set rsLavorazioni = Nothing
End If

Set rsLavorazioni = New ADODB.Recordset
rsLavorazioni.CursorLocation = adUseClient

'RECUPERO COLONNE
'sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
'sSQL = sSQL & " WHERE IDRV_POAssegnazioneMerce=0"
'
'Set rs = New ADODB.Recordset
'
'rs.Open sSQL, CnDMT.InternalConnection

'For I = 0 To rs.Fields.Count - 1
'    Select Case rs.Fields(I).Type
'        Case adChar, adVarChar, adVarWChar, adWChar, 201
'            rsLavorazioni.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
'
'        Case adInteger
'            rsLavorazioni.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
'
'        Case adDate, adDBDate, adDBTime, adDBTimeStamp
'            rsLavorazioni.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
'
'        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
'            rsLavorazioni.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
'
'        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
'            rsLavorazioni.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
'    End Select
'
'Next
rsLavorazioni.Fields.Append "Elimina", adSmallInt, , adFldIsNullable
rsLavorazioni.Fields.Append "IDRV_POAssegnazioneMerce", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "Qta_UM", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "Colli", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "PesoLordo", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "Tara", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "PesoNetto", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "Pezzi", adDouble, , adFldIsNullable
rsLavorazioni.Fields.Append "CodicePedana", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "IDAnagraficaSocio", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "CodiceSocio", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "AnagraficaSocio", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "NomeSocio", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "DataConferimento", adDBDate, , adFldIsNullable
rsLavorazioni.Fields.Append "NumeroConferimento", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "CodiceLottoVendita", adVarChar, 250, adFldIsNullable
rsLavorazioni.Fields.Append "IDRV_POProcessoIVGamma", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "AnnoProcesso", adInteger, , adFldIsNullable
rsLavorazioni.Fields.Append "NumeroProcesso", adInteger, , adFldIsNullable

'rs.Close
'Set rs = Nothing

rsLavorazioni.Open , , adOpenKeyset, adLockBatchOptimistic

'RECUPERO DATI
sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND Colli<=" & fnNormNumber(0.99)


If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDCliente=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
End If

If Me.txtNListaSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroListaPrelievo=" & Me.txtNListaSmistamento.Value
End If
If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
End If

If Len(Me.txtCodicePedana.Text) > 0 Then
    sSQL = sSQL & " AND CodicePedana LIKE " & fnNormString(Me.txtCodicePedana.Text)
End If

If Me.CDSocio.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If

If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
    sSQL = sSQL & " AND CodiceLottoVendita LIKE " & fnNormString(Me.txtLottoVendita.Text)
End If

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection
N_EliminaLav = 0

While Not rs.EOF
    rsLavorazioni.AddNew
        
        For I = 0 To rsLavorazioni.Fields.Count - 1
            If rsLavorazioni.Fields(I).Name <> "Elimina" Then
                rsLavorazioni.Fields(I).Value = rs.Fields(rsLavorazioni.Fields(I).Name).Value
            End If
        Next
        
        rsLavorazioni!Elimina = 1
        N_EliminaLav = N_EliminaLav + 1
        
    rsLavorazioni.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'APERTURA FORM
frmEliminaLav.Show vbModal

fncGrigliaSmistamento

Exit Sub
ERR_cmdEliminaLavorazioni_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaLavorazioni_Click"
End Sub

Private Sub cmdEliminaRifLetInt_Click()
On Error GoTo ERR_cmdEliminaRifLetInt_Click
Dim Testo As String
Dim LINK_CLIENTE_IVA As Long

If Me.txtIDLetteraIntento.Value = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il riferimento alla lettera d'intento?" & vbCrLf
If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento lettera d'intento") = vbNo Then Exit Sub

Me.txtIDLetteraIntento.Value = 0

LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdCliente.KeyFieldID)

LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
    
Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA

Exit Sub
ERR_cmdEliminaRifLetInt_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifLetInt_Click"
End Sub

Private Sub cmdGiu_Click()
On Error GoTo ERR_cmdGiu_Click
If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

If rsGrigliaAss Is Nothing Then Exit Sub
If ((rsGrigliaAss.EOF) And (rsGrigliaAss.BOF)) Then Exit Sub

If LINK_CLIENTE_ORD_PRED = 0 Then
    If Me.CDAltroCliente.KeyFieldID = 0 Then
        MsgBox "Inserire il cliente dello smistamento", vbInformation, "Smistamento merce"
        Me.CDAltroCliente.SetFocus
        Exit Sub
    End If
    
    If Me.txtNumeroSmistamento.Value = 0 Then
        MsgBox "Inserire il numero dell'ordine", vbInformation, "Smistamento merce"
        Me.txtNumeroSmistamento.SetFocus
        Exit Sub
    End If
    
    If Me.txtDataSmistamento.Value = 0 Then
        MsgBox "Inserire la data dell'ordine", vbInformation, "Smistamento merce"
        Me.txtDataSmistamento.SetFocus
        Exit Sub
    End If
    
    If Me.txtNListaSmistamento.Value = 0 Then
        MsgBox "Inserire il numero della lista prelievo", vbInformation, "Smistamento merce"
        Me.txtNListaSmistamento.SetFocus
        Exit Sub
    End If

End If

Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value, Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value, 0)

If Len(Testo) > 0 Then
    MsgBox Testo, vbCritical, "Assegnazione merce"
    Exit Sub
End If

LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO = fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value)
LINK_ORDINE_MERCE_PER_SMISTAMENTO = GET_ESISTENZA_ORDINE(Me.CDAltroCliente.KeyFieldID, Me.txtDataSmistamento.Text, Me.txtNumeroSmistamento.Value, Me.txtNListaSmistamento.Value)

If LINK_ORDINE_MERCE_PER_SMISTAMENTO = 0 Then
    MsgBox "Selezionare un ordine", vbInformation, "Smistamento merce"
    Exit Sub
End If

COMANDO_SPEZZATUTA = 1
COMANDO_RIPESATURA = 0
SPACCATURA_MERCE_VERSO = 1
LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO = Me.CDAltroCliente.KeyFieldID
NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtNumeroSmistamento.Value
DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtDataSmistamento.Text
NUMERO_LISTA_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtNListaSmistamento.Value
LINK_ORDINE_PADRE_MERCE_PER_SMISTAMENTO = GET_LINK_ORDINE_PADRE(LINK_ORDINE_MERCE_PER_SMISTAMENTO)

If (LINK_ORDINE_PADRE_MERCE_PER_SMISTAMENTO) = 0 Then
    MsgBox "Impossibile recuperare il riferimento padre dell'ordine selezionato", vbInformation, "Smistamento merce"
    Exit Sub
End If

frmAssegnazioneMerce.Show vbModal
    
rsGrigliaAss.UpdateBatch

For I = 0 To 100000

Next

fnGrigliaAssegnazione
fncGrigliaSmistamento
'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
'GET_TOTALI_ORDINE_DA_SMISTARE
GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

If FLAG_ASSEGNAZIONE_VELOCE = True Then
    Me.txtAssegnazioneVeloce.SetFocus
End If
Exit Sub
ERR_cmdGiu_Click:
    MsgBox Err.Description, vbCritical, "cmdGiu_Click"
End Sub

Private Sub cmdGiuSingolo_Click()
On Error GoTo ERR_cmdGiuSingolo_Click
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As ADODB.Recordset
Dim LINK_ORDINE_SMIST As Long
Dim LINK_ORDINE_SMIST_PADRE As Long

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If
    
    
    If Me.CDAltroCliente.KeyFieldID = 0 Then
        MsgBox "Inserire il cliente dello smistamento", vbInformation, "Smistamento merce"
        Me.CDAltroCliente.SetFocus
        Exit Sub
    End If

    If Me.txtNumeroSmistamento.Value = 0 Then
        MsgBox "Inserire il numero dell'ordine", vbInformation, "Smistamento merce"
        Me.txtNumeroSmistamento.SetFocus
        Exit Sub
    End If
    
    If Me.txtDataSmistamento.Value = 0 Then
        MsgBox "Inserire la data dell'ordine", vbInformation, "Smistamento merce"
        Me.txtDataSmistamento.SetFocus
        Exit Sub
    End If
    
    If Me.txtNListaSmistamento.Value = 0 Then
        MsgBox "Inserire il numero lista prelievo dell'ordine", vbInformation, "Smistamento merce"
        Me.txtNListaSmistamento.SetFocus
        Exit Sub
    End If
    
    LINK_ORDINE_SMIST = GET_ESISTENZA_ORDINE(Me.CDAltroCliente.KeyFieldID, Me.txtDataSmistamento.Text, Me.txtNumeroSmistamento.Value, Me.txtNListaSmistamento.Value)
    LINK_ORDINE_SMIST_PADRE = GET_LINK_ORDINE_PADRE(LINK_ORDINE_SMIST)
    
    
    
    If LINK_ORDINE_SMIST = 0 Then
        MsgBox "Ordine di smistamento chiuso o non trovato!", vbCritical, "Smistamente merce"
        Exit Sub
    End If
    
    If LINK_ORDINE_SMIST_PADRE = 0 Then
        MsgBox "Impossibile recuperare il riferimento dell'ordine padre", vbCritical, "Smistamente merce"
        Exit Sub
    End If
    
    
    If LINK_ORDINE_SMIST <> LINK_ORDINE_PRED Then
        If MsgBox("L'ordine indicato non è quello della cooperativa." & vbCrLf & "Continuare?", vbQuestion + vbYesNo, "Smistamente merce") = vbNo Then Exit Sub
    End If
    
    Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value, Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value, 0)
    
    If Len(Testo) > 0 Then
        MsgBox Testo, vbCritical, "Assegnazione merce"
        Exit Sub
    End If

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value)

    Set rs = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenDynamic, adLockPessimistic
    
    
    While Not rs.EOF
        rs!IDOggettoOrdine = LINK_ORDINE_SMIST
        rs!IDCliente = Me.CDAltroCliente.KeyFieldID
        rs!NumeroOrdine = Me.txtNumeroSmistamento.Value
        rs!DataOrdine = Me.txtDataSmistamento.Text
        rs!NumeroListaPrelievo = Me.txtNListaSmistamento.Value
        rs!IDOggettoOrdinePadre = LINK_ORDINE_SMIST_PADRE
        rs!IDValoriOggettoDettaglioRigaOrd = 0
        rs!RV_POImportoUnitarioListino = 0
        rs!ImportoUnitarioArticolo = 0
        rs!ImportoUnitarioImballo = 0
        rs!MerceInclusoImballo = 0
        rs!Sconto1 = 0
        rs!Sconto2 = 0
        
        rs.Update
    rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing

    If GESTIONE_ORDINE_VIVAIO = 1 Then
        ELIMINA_COMMISSIONE_LAVORAZIONE fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value)
    End If

    fnGrigliaAssegnazione

    fncGrigliaSmistamento
    
    
    
    
Screen.MousePointer = 0

GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

If FLAG_ASSEGNAZIONE_VELOCE = True Then
    Me.txtAssegnazioneVeloce.SetFocus
End If

Exit Sub
ERR_cmdGiuSingolo_Click:
    MsgBox Err.Description, vbCritical, "Smistamento merce"
    Screen.MousePointer = 0
End Sub

Private Sub cmdImputazionePrezziVeloce_Click()
    If Me.txtIDOrdine.Value > 0 Then
    
        frmImputazionePrezzi.Show vbModal
        
        fnGrigliaAssegnazione
        
        GET_IMPONIBILE_ORDINE Me.txtIDOrdine.Value
    End If
End Sub



Private Sub cmdLetteraIntento_Click()
On Error GoTo ERR_cmdLetteraIntento_Click
Dim LINK_CLIENTE_IVA As Long

LINK_CLIENTE_IVA = 0
If Me.txtIDOrdine.Value = 0 Then Exit Sub

Set Control_Return = Me.txtIDLetteraIntento
Set Control_Return_Cliente = Me.cdCliente
Set Control_Return_Data_Ordine = Me.txtDataOrdine

frmLetteraIntento.Show vbModal

LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdCliente.KeyFieldID)

If Me.txtIDLetteraIntento.Value > 0 Then
    
    LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
        
    Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
    
End If
    
Exit Sub
ERR_cmdLetteraIntento_Click:
    MsgBox Err.Description, vbCritical, "cmdLetteraIntento_Click"
End Sub

Private Sub cmdNListaPrelievo_Click()
    
    LINK_ORDINE_PADRE_SEL = 0
    
    If (Me.txtIDOrdinePadre.Value > 0) Then
        LINK_ORDINE_PADRE_SEL = Me.txtIDOrdinePadre.Value
        
        NUOVO_ORDINE_LISTA_PRELIEVO = False

        BLOCCA_ORDINE LINK_ORDINE, 0
        
        Me.txtIDOrdine.Value = 0
        
    Else
        frmTrovaOrdinePadre.Show vbModal
        
    End If
    
    If LINK_ORDINE_PADRE_SEL = 0 Then Exit Sub
    
    NUOVO_ORDINE_LISTA_PRELIEVO = True
    
    Me.lblLinkOrdine.DisableDoEvents = True
    Me.lblLinkOrdine.IDReturn = LINK_ORDINE_PADRE_SEL
    Me.lblLinkOrdine.RunApplication
    
End Sub

Private Sub cmdNuovoOrdine_Click()
On Error Resume Next

    NUOVO_ORDINE_LISTA_PRELIEVO = False
    
    BLOCCA_ORDINE LINK_ORDINE, 0
    
    Me.txtIDOrdine.Value = 0
    Me.cdCliente.Load 0
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
    Me.cboDestinazione.WriteOn 0
    Me.cboVettore.WriteOn 0
    Me.txtDataPartenza.Value = 0
    
    
    Me.lblLinkOrdine.IDReturn = 0
    Me.lblLinkOrdine.RunApplication
    
    
    


End Sub

Private Sub cmdReimpostaOrdDft_Click()
    Me.CDAltroCliente.Load 0
    Me.CDSocio.Load 0
    Me.txtCodicePedana.Text = ""

    GET_CLIENTE_ORDINE_PRED
    
    
    fncGrigliaSmistamento
End Sub

Private Sub cmdRicerca_Click()
    
    AGGIORNA_NUMERAZIONE_ORDINE
    
    BLOCCA_ORDINE Me.txtIDOrdine.Value, 0
    
    Me.txtIDOrdine.Value = 0
    Me.cdCliente.Load 0
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
    Me.txtAnnotazioniOrdine.Text = ""
    Me.txtDataOrdineCliente.Value = 0
    Me.txtNumeroOrdineCliente.Text = ""
    Me.txtDataPartenza.Value = 0
    Me.txtAnnotazioniInterna.Text = ""
    Me.txtDescrizioneRigaDoc.Text = ""
    Me.cboLuogoPresaMerce.WriteOn 0
    Me.cboVettoreSuccessivo.WriteOn 0
    Me.txtIDLetteraIntento.Value = 0
    Me.cboIvaCliente.WriteOn 0

    fnGrigliaAssegnazione
    
    frmTrovaOrdine.Show vbModal
    
    Set gPaintNotify = New PaintNotify
    
    If Me.txtIDOrdine.Value > 0 Then
        BLOCCA_ORDINE Me.txtIDOrdine.Value, TheApp.IDUser
        
        Me.GrigliaAssegnazione.SaveUserSettings
        
        fnGrigliaAssegnazione
        
        GET_STATO_ORDINE
    End If
    
    GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value
    
End Sub
Private Sub cmdRiepilogoOridine_Click()
    frmRiepilogoOrdine.Show
End Sub

Private Sub cmdRipesatura_Click()
Dim Testo As String

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

RIPORTA_IN_ORDINE_DA_PESATURA = fnGetParametriMagazzino("RiportaInOrdineClienteDaPesPed")


If RIPORTA_IN_ORDINE_DA_PESATURA = 1 Then
    If Me.txtIDOrdine.Value = 0 Then
        MsgBox "Inserire l'ordine di smistamento", vbInformation, "Associazione merce"
        Exit Sub
    End If
End If

Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(Me.GrigliaSmistamento.AllColumns("IDRV_POCaricoMerceRighe").Value, Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value, 0)

If Len(Testo) > 0 Then
    MsgBox Testo, vbCritical, "Assegnazione merce"
    Exit Sub
End If


If AVVIO_VELOCE_RIPESATURA = 0 Then
    GET_PEDANA_PER_RIPESATURA False
Else
    GET_PEDANA_PER_RIPESATURA True
End If

frmPesaturaPedana.Show vbModal

fnGrigliaAssegnazione

fncGrigliaSmistamento

GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value

'GET_TOTALI_ORDINE_DA_SMISTARE

If FLAG_ASSEGNAZIONE_VELOCE = True Then
    Me.txtAssegnazioneVeloce.SetFocus
End If

End Sub

Private Sub cmdRipPedOrd_Click()
On Error GoTo ERR_cmdRipPedOrd_Click
If Me.txtIDOrdine.Value = 0 Then Exit Sub
    
GET_PEDANA_PER_RIPESATURA_ORDINE False

frmPesPedOrd.Show vbModal

fnGrigliaAssegnazione

fncGrigliaSmistamento

GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value

GET_IMPONIBILE_ORDINE Me.txtIDOrdine.Value

'GET_TOTALI_ORDINE_DA_SMISTARE

Exit Sub
ERR_cmdRipPedOrd_Click:
    MsgBox Err.Description, vbCritical, "cmdRipPedOrd_Click"
    
End Sub

Private Sub cmdSalvaStatoGrigliaOrdine_Click()
    Me.GrigliaAssegnazione.SaveUserSettings
    Me.GrigliaAssegnazione.Refresh
    
End Sub

Private Sub cmdSalvaStatoGrigliaSmist_Click()
    Me.GrigliaSmistamento.SaveUserSettings
    Me.GrigliaSmistamento.Refresh
    
End Sub

Private Sub cmdSelezionaPedana_Click()
    WHERE_TROVA_PEDANA = 1

    frmTrovaPedana.Show vbModal
    
    fncGrigliaSmistamento
End Sub

Private Sub cmdStampa_Click()
    If Me.optSceltaStampa(0).Value = True Then
        StampaDocumento 0
    Else
        StampaDocumento 1
    End If
End Sub
Private Sub cmdStampaOrdineSmist_Click()
    StampaDocumento 1
End Sub

Private Sub cmdStampaReportAut_Click()
    StampaDocumento 0
End Sub

Private Sub cmdSu_Click()
On Error GoTo ERR_cmdSu_Click
Dim Testo As String

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

If Me.txtIDOrdine.Value = 0 Then
    MsgBox "Inserire l'ordine di smistamento", vbInformation, "Associazione merce"
    Exit Sub
End If

If GET_CONTROLLO_RIGA_SMISTAMENTO(fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value)) = False Then
    MsgBox "La riga selezionata è stata associata ad un altro ordine", vbInformation, "Smistamento merce"
    fncGrigliaSmistamento
    Exit Sub
End If

Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(Me.GrigliaSmistamento.AllColumns("IDRV_POCaricoMerceRighe").Value, Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value, 0)

If Len(Testo) > 0 Then
    MsgBox Testo, vbCritical, "Assegnazione merce"
    Exit Sub
End If

    COMANDO_SPEZZATUTA = 0
    COMANDO_RIPESATURA = 0
    SPACCATURA_MERCE_VERSO = 0
    LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO = fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value)
    LINK_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtIDOrdine.Value
    LINK_ORDINE_PADRE_MERCE_PER_SMISTAMENTO = Me.txtIDOrdinePadre.Value
    LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO = Me.cdCliente.KeyFieldID
    NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtNumeroOrdine.Value
    DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtDataOrdine.Text
    NUMERO_LISTA_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtNListaPrelievo.Value
    
    frmAssegnazioneMerce.Show vbModal
    
    rsGrigliaAss.UpdateBatch
    
    For I = 0 To 100000
    
    Next
    
    fnGrigliaAssegnazione
    
    fncGrigliaSmistamento

    GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value
    
'    GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
'    GET_TOTALI_ORDINE_DA_SMISTARE
    
    If FLAG_ASSEGNAZIONE_VELOCE = True Then
        Me.txtAssegnazioneVeloce.SetFocus
    End If

Exit Sub
ERR_cmdSu_Click:
    MsgBox Err.Description, vbCritical, "cmdSu_Click"
End Sub

Private Sub cmdSuSingolo_Click()
On Error GoTo ERR_cmdSuSingolo_Click
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As ADODB.Recordset
Dim Testo As String

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

If Me.txtIDOrdine.Value = 0 Then
    MsgBox "Indicare l'ordine da preparare", vbInformation, "Smistamento merce"
    Me.cmdRicerca.SetFocus
    Exit Sub
End If

Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(Me.GrigliaSmistamento.AllColumns("IDRV_POCaricoMerceRighe").Value, Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value, 0)

If Len(Testo) > 0 Then
    MsgBox Testo, vbCritical, "Assegnazione merce"
    Exit Sub
End If


'If GET_NUMERO_LAVORAZIONI_PER_PEDANA(Me.GrigliaSmistamento.AllColumns("IDRV_POPedana").Value) > 1 Then
'    cmdSu_Click
'    Exit Sub
'End If

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value)

Set rs = New ADODB.Recordset

Screen.MousePointer = 11

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    If GET_ORDINE_CHIUSO(fnNotNullN(rs!IDOggettoOrdine)) = 0 Then
        If GET_CONTROLLO_RIGA_SMISTAMENTO(fnNotNullN(rs!IDRV_POAssegnazioneMerce)) = True Then
            
            rs!IDOggettoOrdine = Me.txtIDOrdine.Value
            rs!IDCliente = Me.cdCliente.KeyFieldID
            rs!NumeroOrdine = Me.txtNumeroOrdine.Value
            rs!DataOrdine = Me.txtDataOrdine.Text
            rs!NumeroListaPrelievo = Me.txtNListaPrelievo.Value
            rs!IDOggettoOrdinePadre = Me.txtIDOrdinePadre.Value
            If PAR_AssNewPedDaAssSingola = 1 Then
                rs!CodicePedana = GetNumeroPedana(DatePart("yyyy", Date))
                rs!IDRV_POPedana = Link_Pedana
            End If
            
            If PREZZI_ARTICOLI_DA_ORDINE = 0 Then
                GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
            Else
                If (GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDImballoVendita), Me.txtIDOrdinePadre.Value, rs) = False) Then
                    GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                    GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                    rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
                Else
                    If RETURN_SEL_PREZZO_IMB_DA_ORD = 0 Then
                        GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                        rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
                    End If
                End If
            End If
            
            rs.Update
            
        End If
    End If
rs.MoveNext
Wend
rs.Close
Set rs = Nothing


fnGrigliaAssegnazione

fncGrigliaSmistamento

'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
'GET_TOTALI_ORDINE_DA_SMISTARE
Screen.MousePointer = 0
GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

If FLAG_ASSEGNAZIONE_VELOCE = True Then
    Me.txtAssegnazioneVeloce.SetFocus
End If

Exit Sub
ERR_cmdSuSingolo_Click:
    MsgBox Err.Description, vbCritical, "Smistamento merce"
    Screen.MousePointer = 0
End Sub

Private Sub cmdSuTutto_Click()
On Error GoTo ERR_cmdSuTutto_Click
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As ADODB.Recordset

If LINK_ORDINE = 0 Then
    MsgBox "Indicare l'ordine da preparare", vbInformation, "Smistamento merce"
    Me.cmdRicerca.SetFocus
    Exit Sub
End If


sSQL_WHERE = ""

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "


If Me.CDAltroCliente.KeyFieldID > 0 Then
    If sSQL_WHERE = "" Then
        sSQL_WHERE = sSQL_WHERE & " WHERE RV_POAssegnazioneMerce.IDCliente=" & Me.CDAltroCliente.KeyFieldID
    Else
        sSQL_WHERE = sSQL_WHERE & " AND RV_POAssegnazioneMerce.IDCliente=" & Me.CDAltroCliente.KeyFieldID
    End If
End If
If Me.txtNumeroSmistamento.Value > 0 Then
    If sSQL_WHERE = "" Then
        sSQL_WHERE = sSQL_WHERE & " WHERE RV_POAssegnazioneMerce.NumeroOrdine=" & Me.txtNumeroSmistamento.Value
    Else
        sSQL_WHERE = sSQL_WHERE & " AND RV_POAssegnazioneMerce.NumeroOrdine=" & Me.txtNumeroSmistamento.Value
    End If
End If
If Me.txtDataSmistamento.Value > 0 Then
    If sSQL_WHERE = "" Then
        sSQL_WHERE = sSQL_WHERE & " WHERE RV_POAssegnazioneMerce.DataOrdine=" & fnNormDate(Me.txtDataSmistamento.Text)
    Else
        sSQL_WHERE = sSQL_WHERE & " AND RV_POAssegnazioneMerce.DataOrdine=" & fnNormDate(Me.txtDataSmistamento.Text)
    End If
End If
If Me.CDArticolo.KeyFieldID > 0 Then
    If sSQL_WHERE = "" Then
        sSQL_WHERE = sSQL_WHERE & " WHERE RV_POAssegnazioneMerce.IDArticolo=" & Me.CDArticolo.KeyFieldID
    Else
        sSQL_WHERE = sSQL_WHERE & " AND RV_POAssegnazioneMerce.IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
End If
If Len(Me.txtCodicePedana.Text) > 0 Then
    If sSQL_WHERE = "" Then
        sSQL_WHERE = sSQL_WHERE & " WHERE RV_POAssegnazioneMerce.CodicePedana=" & fnNormString(Me.txtCodicePedana.Text)
    Else
        sSQL_WHERE = sSQL_WHERE & " AND RV_POAssegnazioneMerce.CodicePedana=" & fnNormString(Me.txtCodicePedana.Text)
    End If
End If
sSQL = sSQL & sSQL_WHERE & " ORDER BY RV_POAssegnazioneMerce.CodiceArticolo"

Set rs = New ADODB.Recordset

Screen.MousePointer = 11

rs.Open sSQL, CnDMT.InternalConnection, adOpenDynamic, adLockPessimistic


While Not rs.EOF
    If GET_ORDINE_CHIUSO(fnNotNullN(rs!IDOggettoOrdine)) = 0 Then
        If GET_CONTROLLO_RIGA_SMISTAMENTO(fnNotNullN(rs!IDRV_POAssegnazioneMerce)) = True Then
            rs!IDOggettoOrdine = LINK_ORDINE
            rs!IDCliente = Me.cdCliente.KeyFieldID
            rs!NumeroOrdine = Me.txtNumeroOrdine.Value
            rs!DataOrdine = Me.txtDataOrdine.Text
            rs.Update
        End If
    End If
rs.MoveNext
Wend
rs.Close
Set rs = Nothing


    fnGrigliaAssegnazione

    fncGrigliaSmistamento

Screen.MousePointer = 0
Exit Sub
ERR_cmdSuTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdSuTutto_Click"

End Sub
Private Function GET_ORDINE_CHIUSO(IDOggetto) As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT Doc_Ordine_chiuso "
    sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_ORDINE_CHIUSO = 1
    Else
        GET_ORDINE_CHIUSO = fnNotNullN(rs!Doc_ordine_chiuso)
    End If
rs.CloseResultset
Set rs = Nothing
End Function

Private Sub cmdTotaliOrdine_Click()
    If Me.FraTotaliOrdine.Visible = True Then
        Me.FraTotaliOrdine.Visible = False
        Me.GrigliaAssegnazione.Width = Me.FraOrdinePrep.Width - 240
        Me.lblInfoOrdine.Width = Me.GrigliaAssegnazione.Width
    Else
        GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
        Me.FraTotaliOrdine.Visible = True
        Me.GrigliaAssegnazione.Width = Me.FraOrdinePrep.Width - Me.FraTotaliOrdine.Width - 240
        Me.lblInfoOrdine.Width = Me.GrigliaAssegnazione.Width
        
    End If
End Sub

Private Sub cmdTotaliOrdSmist_Click()
    If Me.FraTotaliOrdineSmist.Visible = True Then
        Me.FraTotaliOrdineSmist.Visible = False
        Me.GrigliaSmistamento.Width = Me.FraOrdineSmist.Width - 240
        
    Else
        GET_TOTALI_ORDINE_DA_SMISTARE
        Me.FraTotaliOrdineSmist.Visible = True
        Me.GrigliaSmistamento.Width = Me.FraOrdineSmist.Width - Me.FraTotaliOrdineSmist.Width - 240
        
    End If
End Sub

Private Sub cmdTrasporta_Click()
Dim sSQL As String
Dim rsConf As DmtOleDbLib.adoResultset
Dim Testo As String

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

If Me.txtIDOrdine.Value = 0 Then Exit Sub


'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE STANDARD''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & 15
sSQL = sSQL & "  AND IDFunzione=" & 128

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    MsgBox "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente), vbCritical, "Impossibile aggiornare ordine"
    rsConf.CloseResultset
    Set rsConf = Nothing
    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE GREEN TOP'''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POOrdineL")
sSQL = sSQL & "  AND IDFunzione=" & GET_FUNZIONE(fnGetTipoOggetto("RV_POOrdineL"))

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    MsgBox "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente), vbCritical, "Impossibile aggiornare ordine"
    rsConf.CloseResultset
    Set rsConf = Nothing
    
    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CONTROLLO BLOCCO CLIENTE
If GET_CONTROLLO_CLIENTE_BLOCCATO(Me.cdCliente.KeyFieldID) = True Then
    MsgBox "Il cliente risulta bloccato, pertanto è impossibile confermare l'ordine", vbInformation, "Controllo dati"
    Exit Sub
End If

If PAR_CalcImpAConfOrd = 1 Then
    PREZZATURA_ORDINE Me.txtIDOrdine.Value
End If

If PAR_NonVisMsgImpZeroConfOrd = 0 Then
    If GET_CONTROLLO_IMPORTI_A_ZERO(Me.txtIDOrdine.Value) = True Then
        Testo = "ATTENZIONE!!!!!" & vbCrLf
        Testo = Testo & "Nell'ordine confermato risultano importi articolo o importi imballo a zero " & vbCrLf
        Testo = Testo & "Vuoi Continuare?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Conferma ordine") = vbNo Then Exit Sub
    End If

    If GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE(Me.txtIDOrdine.Value) = False Then
        Testo = "ATTENZIONE!!!!!" & vbCrLf
        Testo = Testo & "Nell'ordine confermato è impostato un agente che ha una regola predefinita per il calcolo della provvigione " & vbCrLf
        Testo = Testo & "Vuoi Continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Conferma ordine") = vbNo Then Exit Sub
    End If
End If
If GET_CONTROLLO_QTA_LIQ_A_ZERO(Me.txtIDOrdine.Value) = True Then
    Testo = "ATTENZIONE!!!!!" & vbCrLf
    Testo = Testo & "Nell'ordine confermato sono presenti una o più lavorazioni che la quantità di liquidazione uguale a zero " & vbCrLf
    Testo = Testo & "Impossibile continuare"
    MsgBox Testo, vbCritical, "Conferma ordine"
    Exit Sub
End If





KillTimer Me.hwnd, 0

If GESTIONE_ORDINE_VIVAIO = 1 Then
    GESTIONE_COMMISSIONE_PER_CONFERIMENTO txtIDOrdine.Value, True
End If

rsGrigliaAss.UpdateBatch

If GET_ESISTENZA_ORDINE_DA_ELABORARE = False Then
    sSQL = "INSERT INTO RV_POTMPEvasioneOrdini ("
    sSQL = sSQL & "NumeroRiga, IDOggetto, IDAzienda, NumeroOrdine, DataOrdine, DaRegistrare, "
    sSQL = sSQL & "IDCliente, IDSitoPerAnagrafica, Cliente, SitoPerAnagrafica, "
    sSQL = sSQL & "IDVettore, Vettore, DescrizioneCorpoDocEv, IDLuogoPresaMerce, IDVettoreSuccessivo, InFatturazione, NumeroListaPrelievo) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_POTMPEvasioneOrdini", "NumeroRiga") & ", "
    sSQL = sSQL & Me.txtIDOrdine.Value & ", "
    sSQL = sSQL & TheApp.IDFirm & ","
    sSQL = sSQL & Me.txtNumeroOrdine.Value & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataOrdine.Text) & ", "
    sSQL = sSQL & fnNormBoolean(0) & ", "
    sSQL = sSQL & Me.cdCliente.KeyFieldID & ", "
    sSQL = sSQL & Me.cboDestinazione.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cdCliente.Description) & ", "
    sSQL = sSQL & fnNormString(Me.cboDestinazione.Text) & ", "
    sSQL = sSQL & Me.cboVettore.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cboVettore.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtDescrizioneRigaDoc.Text) & ", "
    sSQL = sSQL & Me.cboLuogoPresaMerce.CurrentID & ", "
    sSQL = sSQL & Me.cboVettoreSuccessivo.CurrentID & ", "
    sSQL = sSQL & fnNormBoolean(0) & ", "
    sSQL = sSQL & Me.txtNListaPrelievo.Value & ")"
    
    CnDMT.Execute sSQL
End If

BLOCCA_ORDINE Me.txtIDOrdine.Value, 0

LINK_ORDINE = 0
Me.txtIDOrdine.Value = 0
Me.cdCliente.Load 0
Me.txtNumeroOrdine.Value = 0
Me.txtDataOrdine.Value = 0

fnGrigliaOrdiniDaElaborare

fnGrigliaAssegnazione

lngTimerID = SetTimer(Me.hwnd, 0, 5000, AddressOf TimerProc)

End Sub
Private Function GET_SITO_ANAGRAFICA(IDOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SitoPerAnagrafica "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F LEFT OUTER JOIN "
sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo000F.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica "
sSQL = sSQL & "WHERE IDOggetto = " & IDOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_SITO_ANAGRAFICA = ""
Else
    GET_SITO_ANAGRAFICA = fnNotNull(rs!SitoPerAnagrafica)
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_SITO_ANAGRAFICA(IDOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Link_Nom_Ult_Sito "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto = " & IDOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SITO_ANAGRAFICA = 0
Else
    GET_LINK_SITO_ANAGRAFICA = fnNotNull(rs!Link_nom_ult_sito)
End If


rs.CloseResultset
Set rs = Nothing

End Function

Private Sub cmdTrovaOrdineSmistamento_Click()
    AGGIORNA_NUMERAZIONE_ORDINE
    frmTrovaOrdineSmistamento.Show vbModal
    fncGrigliaSmistamento
End Sub



Private Sub Command1_Click()
    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If
    
    
    
    If Me.txtIDOrdine.Value = 0 Then
        MsgBox "Indicare l'ordine da preparare", vbInformation, "Smistamento merce"
        Me.cmdRicerca.SetFocus
        Exit Sub
    End If
    
    frmSpezzaturaVeloce.Show vbModal
    
    fnGrigliaAssegnazione
    
    fncGrigliaSmistamento
    
    GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value
    
    'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
    
    'GET_TOTALI_ORDINE_DA_SMISTARE
    
    If FLAG_ASSEGNAZIONE_VELOCE = True Then
        Me.txtAssegnazioneVeloce.SetFocus
    End If
    
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
If FLAG_ASSEGNAZIONE_VELOCE = True Then
    Me.txtAssegnazioneVeloce.SetFocus
End If

If BLoading = False Then
    frmRiepilogoOrdine.Show
    Unload frmRiepilogoOrdine
End If

LINK_LISTINO_AZIENDA = GET_LINK_LISTINO_AZIENDA
LINK_LISTINO_IMBALLI_AZIENDA = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA


    
If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
End If

Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If Len(Trim(Me.txtAssegnazioneVeloce.Text)) > 0 Then
            AvviaAssegnazioneVeloce
        End If
    End If
    
    
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
Dim oActivity As IActivity
Dim o As Activity
Dim oFilter As Filter

    FLAG_ASSEGNAZIONE_VELOCE = False

    If BLoading = True Then
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
        
        Me.optSceltaStampa(0).Value = True
        
    
    
        'Inizializzazione della LabelLink
        '----------------------------
        Set LabelLink1.Application = TheApp    'L’oggetto Application
        LabelLink1.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
        LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POAssegnazioneMerce"))
        LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    
        Set LabelLink2.Application = TheApp    'L’oggetto Application
        LabelLink2.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
        LabelLink2.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POAssegnazioneMerce"))
        LabelLink2.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    
        Set lblLinkOrdine.Application = TheApp    'L’oggetto Application
        lblLinkOrdine.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
        lblLinkOrdine.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POOrdineL"))
        lblLinkOrdine.PopMenuItems("Mnu_SearchObject").Enabled = False
        lblLinkOrdine.DisableDoEvents = True
    
        With Me.cboVettoreSuccessivo
            Set .Database = TheApp.Database.Connection
            .AddFieldKey "IDVettore"
            .DisplayField = "Vettore"
            .Sql = "SELECT * FROM Vettore ORDER BY Vettore"
            .Fill
        End With
    
        With Me.cboLuogoPresaMerce
            Set .Database = TheApp.Database.Connection
            .AddFieldKey "IDSitoPerAnagrafica"
            .DisplayField = "SitoPerAnagrafica"
            .Sql = "SELECT * FROM SitoPerAnagrafica  "
            .Sql = .Sql & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
            .Sql = .Sql & " ORDER BY SitoPerAnagrafica "
            .Fill
        End With

        

    End If
    
    
    
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

    lngTimerID = SetTimer(Me.hwnd, 0, 5000, AddressOf TimerProc)
    BLoading = True
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo ERR_Form_Resize
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
Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"
End Sub

Private Sub GrigliaAssegnazione_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
On Error GoTo ERR_AggiornamentoRiga
    
    rsGrigliaAss.UpdateBatch

Exit Sub
ERR_AggiornamentoRiga:
    MsgBox Err.Description, vbCritical, "ERR_AggiornamentoRiga"
End Sub

Private Sub GrigliaAssegnazione_DblClick()
    If fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value) > 0 Then
        Me.LabelLink1.IDReturn = fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value)
        Me.LabelLink1.RunApplication
    Else
        If fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POProcessoIVGamma").Value) > 0 Then
            Me.LabelLink4.IDReturn = fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POProcessoIVGamma").Value)
            Me.LabelLink4.RunApplication
        End If
    End If

    'If fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value) = 0 Then Exit Sub
    'Me.LabelLink1.IDReturn = fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POCaricoMerceRighe").Value)
    'Me.LabelLink1.RunApplication
End Sub
Private Sub GrigliaAssegnazione_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaAssegnazione.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGrigliaAss.Fields("MerceInclusoImballo").Value)), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsGrigliaAss.EOF And Not rsGrigliaAss.BOF Then
                
        rsGrigliaAss.Fields("MerceInclusoImballo").Value = Abs(CLng(Selected))
        'sbCheckSelected
        

        rsGrigliaAss.UpdateBatch
        Me.GrigliaAssegnazione.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub

Private Function GET_PREZZO_IMBALLO(IDImballo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDListinoDefault As Long

IDListinoDefault = GET_LISTINO_DEFAULT(Me.cdCliente.KeyFieldID, Me.cboDestinazione.CurrentID)

sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(IDListino=" & IDListinoDefault & ") "
sSQL = sSQL & "AND (IDArticolo=" & IDImballo & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIva)
Else
    GET_PREZZO_IMBALLO = 0
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_PREZZO_IMBALLO_2(IDArticolo As Long, IDListinoCliente As Long, IDListinoAzienda As Long, IDListinoParAzienda As Long, IDOggettoOrdine As Long, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDDestinazione As Long, IDCliente As Long) As Double
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
    IDArticoloPadre = IDArticolo 'GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
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
            '''''''''''''''''''TROVO IL LINK_RIGA DELL'ORDINE'''''''''''''''''''''''''''
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            
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



Private Function GET_LISTINO_DEFAULT(IDAnagraficaCliente As Long, IDDestinazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset
Dim Link_Listino_Imballo As Long

GET_LISTINO_DEFAULT = 0

''''''LISTINO CLIENTE PER DESTINAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''
If IDDestinazione > 0 Then
    sSQL = "SELECT IDListino "
    sSQL = sSQL & "FROM RV_POConfigurazioneClienteListino "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDDestinazione
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_LISTINO_DEFAULT = 0
    Else
        GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListino)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LISTINO_DEFAULT > 0 Then Exit Function

''''''''''''''''''''LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDefault)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LISTINO_DEFAULT > 0 Then Exit Function


'''''LISTINO NEI PARAMETRI DI GREENTOP''''''''''''''''''''''''''''''''''''''''''''''''''
GET_LISTINO_DEFAULT = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GET_LISTINO_DEFAULT > 0 Then Exit Function

'''''''''''''''''''''''LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Sub GrigliaSmistamento_DblClick()
On Error GoTo ERR_GrigliaSmistamento_DblClick
    If fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POCaricoMerceRighe").Value) > 0 Then
        Me.LabelLink2.IDReturn = fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POCaricoMerceRighe").Value)
        Me.LabelLink2.RunApplication
    Else
        If fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POProcessoIVGamma").Value) > 0 Then
            Me.LabelLink3.IDReturn = fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POProcessoIVGamma").Value)
            Me.LabelLink3.RunApplication
        End If
    End If
Exit Sub
ERR_GrigliaSmistamento_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaSmistamento_DblClick"

End Sub


Private Sub Label8_Click(Index As Integer)
    GET_IMPONIBILE_ORDINE txtIDOrdine.Value
End Sub

Private Sub LabelLink1_AfterRunServerApplication(ByVal lIDResultKey As Long)
On Error Resume Next
Dim NumeroRecord As Long

    NumeroRecord = Me.GrigliaAssegnazione.ListIndex - 1
    
    fnGrigliaAssegnazione
    
    fncGrigliaSmistamento
    
    Me.GrigliaAssegnazione.Recordset.Move NumeroRecord
    
    Me.WindowState = 2
End Sub

Private Sub LabelLink1_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "IDLavorazione", fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POAssegnazioneMerce").Value)
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "Tipo", 1
End Sub

Private Sub LabelLink2_AfterRunServerApplication(ByVal lIDResultKey As Long)
On Error Resume Next
Dim NumeroRecord As Long

    NumeroRecord = Me.GrigliaSmistamento.ListIndex - 1
    fncGrigliaSmistamento
    fnGrigliaAssegnazione
    
    Me.GrigliaSmistamento.Recordset.Move NumeroRecord
    
    Me.WindowState = 2
    
End Sub

Private Sub LabelLink2_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "IDLavorazione", fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POAssegnazioneMerce").Value)
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "Tipo", 1
End Sub

Private Sub LabelLink3_AfterRunServerApplication(ByVal lIDResultKey As Long)
On Error Resume Next
Dim NumeroRecord As Long

    NumeroRecord = Me.GrigliaSmistamento.ListIndex - 1
    fncGrigliaSmistamento
    fnGrigliaAssegnazione
    
    Me.GrigliaSmistamento.Recordset.Move NumeroRecord
    
    Me.WindowState = 2
End Sub

Private Sub LabelLink3_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POIVGamma", "IDProcessoIVGamma", fnNotNullN(Me.GrigliaSmistamento.AllColumns("IDRV_POProcessoIVGamma").Value)
End Sub

Private Sub LabelLink4_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POIVGamma", "IDProcessoIVGamma", fnNotNullN(Me.GrigliaAssegnazione.AllColumns("IDRV_POProcessoIVGamma").Value)
    
End Sub



Private Sub lblLinkOrdine_AfterRunServerApplication(ByVal lIDResultKey As Long)
    If lIDResultKey > 0 Then
        LINK_ORDINE = lIDResultKey
        Me.txtIDOrdine.Value = LINK_ORDINE
        BLOCCA_ORDINE LINK_ORDINE, TheApp.IDUser
'
'        CDAltroCliente.Load Me.cdCliente.KeyFieldID
'        txtDataSmistamento.Value = Me.txtDataOrdine.Value
'        txtNumeroSmistamento.Value = Me.txtNumeroOrdine.Value
'        Me.txtNListaSmistamento.Value = 1
        
        fncGrigliaSmistamento
        
    End If
End Sub

Private Sub lblLinkOrdine_BeforeRunServerApplication()

    If NUOVO_ORDINE_LISTA_PRELIEVO = True Then
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POOrdineL", "IDOggettoOrdinePadre", LINK_ORDINE_PADRE_SEL
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POOrdineL", "NoReturnValue", 1
    Else
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POOrdineL", "NoReturnValue", 0
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POOrdineL", "IDOggettoOrdinePadre", 0
    End If
End Sub

Private Sub txtCodicePedana_LostFocus()
    fncGrigliaSmistamento
End Sub



Private Sub txtDataSmistamento_LostFocus()
    fncGrigliaSmistamento
End Sub

Private Sub txtIDLetteraIntento_Change()
On Error GoTo ERR_txtIDLetteraIntento_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If IsEmpty(Me.txtIDLetteraIntento.Value) Then Me.txtIDLetteraIntento.Value = 0

sSQL = "SELECT * FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & Me.txtIDLetteraIntento.Value

Set rs = CnDMT.OpenResultset(sSQL)

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

Private Sub txtIDOrdine_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If Me.txtIDOrdine.Value = 0 Then

    Me.cboDestinazione.WriteOn 0
    Me.cboVettore.WriteOn 0
    Me.txtDataPartenza.Value = 0
    Me.txtDataOrdineCliente.Value = 0
    Me.txtNumeroOrdineCliente.Text = ""
    Me.txtAnnotazioniOrdine.Text = ""
    Me.txtAnnotazioniInterna.Text = ""
    Me.txtDescrizioneRigaDoc.Text = ""
    Me.cboLuogoPresaMerce.WriteOn 0
    Me.cboVettoreSuccessivo.WriteOn 0
    Me.cboTipoTrasporto.WriteOn 0
    Me.txtDataArrivoMerce.Value = 0
    Me.txtDataArrivoMerceL.Value = 0
    Me.txtOraArrivoMerce.Value = 0
    Me.txtOraArrivoMerceL.Value = 0
    Me.txtIDLetteraIntento.Value = 0
    Me.cboIvaCliente.WriteOn 0
    
    Me.txtIDOrdinePadre.Value = 0
    Me.cboSezionaleOrdOri.WriteOn 0
    Me.txtNumeroOrdineOri.Text = ""
    Me.txtDataOrdineOri.Value = 0
    
    Me.txtNListaPrelievo.Value = 0
    
    
    GET_STATO_ORDINE

    fnGrigliaAssegnazione
    
    GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value
    GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
    GET_TOTALI_ORDINE_DA_SMISTARE
    GET_IMPONIBILE_ORDINE Me.txtIDOrdine.Value
    Exit Sub
    
End If



sSQL = "SELECT Link_Vet_Vettore, Link_nom_ult_sito, Doc_ordine_chiuso, Doc_data_prevista_evasione, "
sSQL = sSQL & "Doc_data_presso_nom, Doc_numero_presso_nom, Doc_annotazioni_variazio, RV_POAnnotazioniInterna, "
sSQL = sSQL & "RV_PODescrizioneCorpoDocEv, RV_POIDLuogoPresaMerce, RV_POIDTrasportatoreSuccessivo, RV_POIDUtenteBlocco, "
sSQL = sSQL & "RV_PODataArrivoMerceLuogo, RV_POOraArrivoMerceLuogo, RV_PODataArrivoMerce, RV_POOraArrivoMerce, "
sSQL = sSQL & "Link_Doc_spedizione, Link_Nom_Lettera_intento, Link_Nom_IVA, Doc_data, Doc_numero, Link_nom_Anagrafica, "
sSQL = sSQL & "RV_PODataOrdinePadre, RV_PONumeroOrdinePadre, RV_POIDOrdinePadre, RV_PONumeroListaPrelievo, Link_Doc_sezionale "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDOrdine.Value

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Exit Sub
Else
    
    Me.cdCliente.Load fnNotNullN(rs!link_nom_anagrafica)
    
    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica WHERE IDAnagrafica=" & Me.cdCliente.KeyFieldID
        .Fill
    End With
    
    Me.txtDataOrdine.Value = fnNotNullN(rs!RV_PODataOrdinePadre) ' fnNotNullN(rs!Doc_Data)
    Me.txtNumeroOrdine.Value = fnNotNullN(rs!RV_PONumeroOrdinePadre) 'fnNotNullN(rs!Doc_Numero)
    Me.txtNListaPrelievo.Value = fnNotNullN(rs!RV_PONumeroListaPrelievo)
    
    Me.cboDestinazione.WriteOn fnNotNullN(rs!Link_nom_ult_sito)
    Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
    Me.txtDataPartenza.Value = fnNotNullN(rs!Doc_data_prevista_evasione)
    Me.txtDataOrdineCliente.Value = fnNotNullN(rs!Doc_data_presso_nom)
    Me.txtNumeroOrdineCliente.Text = fnNotNull(rs!Doc_numero_presso_nom)
    Me.txtAnnotazioniOrdine.Text = fnNotNull(rs!Doc_annotazioni_variazio)
    Me.txtAnnotazioniInterna.Text = fnNotNull(rs!RV_POAnnotazioniInterna)
    Me.txtDescrizioneRigaDoc.Text = fnNotNull(rs!RV_PODescrizioneCorpoDocEv)
    Me.cboLuogoPresaMerce.WriteOn fnNotNullN(rs!RV_POIDLuogoPresaMerce)
    Me.cboVettoreSuccessivo.WriteOn fnNotNullN(rs!RV_POIDTrasportatoreSuccessivo)
    Me.cboTipoTrasporto.WriteOn fnNotNullN(rs!Link_Doc_spedizione)
    Me.txtDataArrivoMerce.Value = fnNotNullN(rs!RV_PODataArrivoMerce)
    Me.txtDataArrivoMerceL.Value = fnNotNullN(rs!RV_PODataArrivoMerceLuogo)
    Me.txtOraArrivoMerce.Text = fnNotNull(rs!RV_POOraArrivoMerce)
    Me.txtOraArrivoMerceL.Text = fnNotNull(rs!RV_POOraArrivoMerceLuogo)
    Me.txtIDLetteraIntento.Value = fnNotNullN(rs!Link_Nom_Lettera_Intento)
    Me.cboIvaCliente.WriteOn fnNotNullN(rs!Link_Nom_IVA)
    
        
    Me.txtIDOrdinePadre = fnNotNullN(rs!RV_POIDOrdinePadre)
    Me.txtNumeroOrdineOri.Text = fnNotNullN(rs!Doc_numero)
    Me.txtDataOrdineOri.Value = fnNotNullN(rs!Doc_data)
    Me.cboSezionaleOrdOri.WriteOn fnNotNullN(rs!Link_Doc_sezionale)
    
    
End If

rs.CloseResultset
Set rs = Nothing

fnGrigliaAssegnazione

LINK_LISTINO_CLIENTE = GET_LINK_LISTINO_CLIENTE(Me.cdCliente.KeyFieldID, Me.cboDestinazione.CurrentID)

GET_STATO_ORDINE

GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

'GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
GET_IMPONIBILE_ORDINE Me.txtIDOrdine.Value
'GET_TOTALI_ORDINE_DA_SMISTARE


GET_INTESTAZIONE_DOCUMENTO Me.cdCliente.KeyFieldID, LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA



End Sub



Private Sub txtLottoVendita_LostFocus()
    fncGrigliaSmistamento
End Sub



Private Sub txtNumeroSmistamento_LostFocus()
    fncGrigliaSmistamento
End Sub

Private Sub txtRaggrOrdine_LostFocus()
    fncGrigliaSmistamento
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

    KillTimer Me.hwnd, 0
    'If frmRiepilogoOrdine.Visible = 0 Then
        Unload frmRiepilogoOrdine
    'End If

    If cmdAvanti = True Then
        If Not (rsGrigliaAss Is Nothing) Then
            rsGrigliaAss.Close
            Set rsGrigliaAss = Nothing
        End If
        If Not (rsGrigliaOrd Is Nothing) Then
            rsGrigliaOrd.CloseResultset
            Set rsGrigliaOrd = Nothing
        End If
        If Not (rsGrigliaSmistamento Is Nothing) Then
            rsGrigliaSmistamento.CloseResultset
            Set rsGrigliaSmistamento = Nothing
        End If
        
        
        BLOCCA_ORDINE LINK_ORDINE, 0
    
        FrmFine.Show
        
    Exit Sub
    End If
    
    BLOCCA_ORDINE LINK_ORDINE, 0

End Sub
Private Sub fnGrigliaAssegnazione()
On Error GoTo ERR_fnGrigliaAssegnazione
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdine=" & Me.txtIDOrdine.Value
    sSQL = sSQL & " AND ((PreConferimento=0) OR (PreConferimento IS NULL)) "
    sSQL = sSQL & " ORDER BY CodicePedana, CodiceArticolo"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        If Not (rsGrigliaAss Is Nothing) Then
            rsGrigliaAss.Close
            Set rsGrigliaAss = Nothing
        End If
        
        Set rsGrigliaAss = New ADODB.Recordset
        rsGrigliaAss.CursorLocation = adUseClient
        rsGrigliaAss.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
        
        With Me.GrigliaAssegnazione
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Clear
                    .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "ID", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "IDRV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "Articolo", "Articolo", dgchar, False, 2000, dgAlignleft, True, True, False
                    Set cl = .ColumnsHeader.Add("Qta_UM", "Quantità", dgDouble, True, 900, dgAlignRight, True, True, False)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 3
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("ImportoUnitarioArticolo", "Importo Art.", dgDouble, True, 900, dgAlignRight, True, True, False)
                        cl.Editable = True
                        cl.BackColor = vbYellow
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Sconto1", "% Sc. 1", dgDouble, True, 900, dgAlignRight, True, True, False)
                        cl.Editable = True
                        cl.BackColor = vbYellow
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Sconto2", "% Sc. 2", dgDouble, True, 900, dgAlignRight, True, True, False)
                        cl.Editable = True
                        cl.BackColor = vbYellow
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                       ' cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                        
                    .ColumnsHeader.Add "Colli", "Colli", dgDouble, True, 1100, dgAlignRight, True, True, False
                    Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, False, 1100, dgAlignRight, True, True, False)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                    Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, False, 1100, dgAlignRight, True, True, False)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                    Set cl = .ColumnsHeader.Add("PesoNetto", "PesoNetto", dgDouble, False, 1100, dgAlignRight, True, True, False)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                        
                    .ColumnsHeader.Add "Pezzi", "Pezzi", dgDouble, False, 1100, dgAlignRight, True, True, False
                    .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "IDImballoVendita", "IDImballo", dgNumeric, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imb.", dgchar, True, 1700, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft, True, True, False
                    Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Importo Imb.", dgDouble, True, 900, dgAlignRight, True, True, False)
                        cl.Editable = True
                        cl.BackColor = vbYellow
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("MerceInclusoImballo", "Imp. incl. Imb.", dgBoolean, True, 1000, dgAligncenter, True, True, False)
                        cl.Editable = True
                        cl.BackColor = vbYellow
                    .ColumnsHeader.Add "CodicePedana", "Pedana", dgchar, True, 1100, dgAlignRight, True, True, False
                    .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, True, 1200, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1700, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1000, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, False, 1000, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "DataDocumento", "Data lav.", dgDate, True, 1500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "OraLavorazione", "Ora lav.", dgchar, True, 1500, dgAlignleft, True, True, False
                    Set cl = .ColumnsHeader.Add("NotaRigaOrdRaggr", "Raggrupp. ord.", dgchar, True, 2500, dgAlignleft, True, True, False)
                        cl.Editable = True
                    .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, False, 2000, dgAlignleft, True, , True
                    .ColumnsHeader.Add "TipoCategoria", "Categoria", dgchar, False, 2000, dgAlignleft, True, , True
                    .ColumnsHeader.Add "Calibro", "Calibro", dgchar, False, 2000, dgAlignleft, True, , True
            
            Set .Recordset = rsGrigliaAss
            .LoadUserSettings
            .Refresh
        End With
    CnDMT.CursorLocation = OLDCursor

    If FraTotaliOrdine.Visible = True Then
        GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
    End If
    
    
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub
Private Sub fncGrigliaOrdine()

End Sub

Private Function GET_STATO_ORDINE() As String
On Error GoTo ERR_GET_STATO_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.lblInfoOrdine.Caption = ""
Me.lblInfoOrdine.BackStyle = 0

If Me.txtIDOrdine.Value = 0 Then Exit Function




    sSQL = "SELECT Doc_ordine_chiuso "
    sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
    sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_STATO_ORDINE = "Non è stato possibile recuperare lo stato dell'ordine"
    Else
        If fnNotNullN(rs!Doc_ordine_chiuso) = 0 Then
            Me.lblInfoOrdine.BackStyle = 1
            Me.lblInfoOrdine.Caption = "ORDINE APERTO"
            Me.lblInfoOrdine.BackColor = vbGreen
        Else
            Me.lblInfoOrdine.BackStyle = 1
            Me.lblInfoOrdine.Caption = "ORDINE CHIUSO"
            Me.lblInfoOrdine.BackColor = vbRed
        End If
    End If
rs.CloseResultset
Set rs = Nothing

'''''''''''''''''''''''CONTROLLO SE ORDINE CONFERMATO'''''''''''''''''''''
sSQL = "SELECT IDOggetto FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value

Set rs = CnDMT.OpenResultset(sSQL)
If Not rs.EOF Then
    Me.lblInfoOrdine.BackStyle = 1
    Me.lblInfoOrdine.Caption = "ORDINE IN CONFERMA"
    Me.lblInfoOrdine.BackColor = vbYellow
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Function
ERR_GET_STATO_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_STATO_ORDINE"
End Function
Private Function GET_TOTALE_ORDINATA(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0010 ON ValoriOggettoPerTipo000F.IDOggetto = ValoriOggettoDettaglio0010.IDOggetto "
sSQL = sSQL & "WHERE ValoriOggettoPerTipo000F.Link_nom_anagrafica=" & Me.cdCliente.KeyFieldID
sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_numero=" & Me.txtNumeroOrdine.Value
sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_data=" & fnNormDate(Me.txtDataOrdine.Text)
sSQL = sSQL & " AND ValoriOggettoDettaglio0010.Link_art_articolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_ORDINATA = 0
Else
    GET_TOTALE_ORDINATA = fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TOTALE_ASSEGNATA(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    sSQL = "SELECT Sum(Qta_UM) as QtaTotale "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDCliente=" & Me.cdCliente.KeyFieldID
    sSQL = sSQL & " AND NumeroOrdine=" & Me.txtNumeroOrdine.Value
    sSQL = sSQL & " AND DataOrdine=" & fnNormDate(Me.txtDataOrdine.Text)
    sSQL = sSQL & " AND IDArticolo=" & IDArticolo
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_ASSEGNATA = 0
Else
    GET_TOTALE_ASSEGNATA = fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fnGrigliaOrdiniDaElaborare()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem
Dim NumeroRiga As Long
Dim I As Integer
NumeroRiga = Me.LVordiniDaElaborare.SelectedItem.Text
Me.LVordiniDaElaborare.ListItems.Clear



sSQL = "SELECT RV_POTMPEvasioneOrdini.NumeroRiga, RV_POTMPEvasioneOrdini.NumeroOrdine, RV_POTMPEvasioneOrdini.DataOrdine, "
sSQL = sSQL & "RV_POTMPEvasioneOrdini.IDCliente, RV_POTMPEvasioneOrdini.IDSitoPerAnagrafica, RV_POTMPEvasioneOrdini.Cliente,"
sSQL = sSQL & "RV_POTMPEvasioneOrdini.SitoPerAnagrafica, RV_POTMPEvasioneOrdini.Vettore, RV_POTMPEvasioneOrdini.IDVettore, RV_POTMPEvasioneOrdini.NumeroListaPrelievo "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE InFatturazione=" & fnNormBoolean(0)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " ORDER BY NumeroRiga"


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Set oItem = Me.LVordiniDaElaborare.ListItems.Add

    'Popola l 'item della listview
    oItem.Text = fnNotNullN(rs!NumeroRiga)
    oItem.SubItems(1) = fnNotNullN(rs!NumeroOrdine)
    oItem.SubItems(2) = fnNotNull(rs!DataOrdine)
    oItem.SubItems(3) = fnNotNull(rs!NumeroListaPrelievo)
    oItem.SubItems(4) = fnNotNull(rs!Cliente)
    oItem.SubItems(5) = fnNotNull(rs!SitoPerAnagrafica)
    oItem.SubItems(6) = fnNotNull(rs!Vettore)
        

rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

For I = 1 To Me.LVordiniDaElaborare.ListItems.Count
    If NumeroRiga = Me.LVordiniDaElaborare.ListItems(I).Text Then
        Set LVordiniDaElaborare.SelectedItem = LVordiniDaElaborare.ListItems(I)
        Me.LVordiniDaElaborare.SelectedItem.EnsureVisible
    End If
Next

End Sub
Private Function GET_ESISTENZA_ORDINE_DA_ELABORARE() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ORDINE_DA_ELABORARE = False
Else
    GET_ESISTENZA_ORDINE_DA_ELABORARE = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Public Function GET_ESISTENZA_ORDINE(IDAnagrafica As Long, DataOrdine As String, NumeroOrdine As Long, NListaPrelievo As Long) As Long
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_DESTINAZIONE_ORDINE As Long

sSQL = "SELECT * "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND "
sSQL = sSQL & "ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo000F.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica "
sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND Oggetto.IDAzienda=" & TheApp.IDFirm
    
If IDAnagrafica > 0 Then
    sSQL_WHERE = sSQL_WHERE & " AND Link_nom_anagrafica=" & IDAnagrafica
End If

If NumeroOrdine > 0 Then
    sSQL_WHERE = sSQL_WHERE & " AND RV_PONumeroOrdinePadre=" & NumeroOrdine
End If

If Len(DataOrdine) > 0 Then
    sSQL_WHERE = sSQL_WHERE & " AND RV_PODataOrdinePadre=" & fnNormDate(DataOrdine)
End If
If (NListaPrelievo > 0) Then
    sSQL_WHERE = sSQL_WHERE & " AND RV_PONumeroListaPrelievo=" & NListaPrelievo
End If

sSQL = sSQL & sSQL_WHERE
    
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ORDINE = 0
    LINK_DESTINAZIONE_ORDINE = 0
Else
    GET_ESISTENZA_ORDINE = fnNotNullN(rs!IDOggetto)
    LINK_DESTINAZIONE_ORDINE = fnNotNullN(rs!Link_nom_ult_sito)
End If

rs.CloseResultset
Set rs = Nothing

LINK_LISTINO_CLIENTE_SOTTO = GET_LINK_LISTINO_CLIENTE(Me.CDAltroCliente.KeyFieldID, LINK_DESTINAZIONE_ORDINE)


End Function

Private Sub fncGrigliaSmistamento()
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
    sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " AND Doc_ordine_chiuso = 0 "
    sSQL = sSQL & " AND Qta_UM>0"
    sSQL = sSQL & " AND ((PreConferimento=0) OR (PreConferimento IS NULL)) "
    
    If Me.CDAltroCliente.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDCliente=" & Me.CDAltroCliente.KeyFieldID
    End If
    
    If Me.txtNumeroSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
    End If
    
    If Me.txtDataSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
    End If
    
    If Me.txtNListaSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PONumeroListaPrelievo=" & Me.txtNListaSmistamento.Value
    End If
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    
    If Len(Me.txtCodicePedana.Text) > 0 Then
        sSQL = sSQL & " AND CodicePedana LIKE " & fnNormString(Me.txtCodicePedana.Text)
    End If
    
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If
    If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
        sSQL = sSQL & " AND CodiceLottoVendita LIKE " & fnNormString(Me.txtLottoVendita.Text)
    End If
    If Len(Trim(Me.txtRaggrOrdine.Text)) > 0 Then
        sSQL = sSQL & " AND NotaRigaOrdRaggr LIKE " & fnNormString("%" & Me.txtRaggrOrdine.Text & "%")
    End If
    sSQL = sSQL & " ORDER BY CodicePedana, CodiceArticolo"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        Set rsGrigliaSmistamento = CnDMT.OpenResultset(sSQL)
            Set rsEvent = rsGrigliaSmistamento.Data
    
        With Me.GrigliaSmistamento
            Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Clear
            
            .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "ID", dgInteger, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft, True, True, False
            Set cl = .ColumnsHeader.Add("Qta_UM", "Quantità", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 3
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 1
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
            Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
            Set cl = .ColumnsHeader.Add("PesoNetto", "PesoNetto", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                
            .ColumnsHeader.Add "Pezzi", "Pezzi", dgDouble, False, 1100, dgAlignRight, True, True, False
           
            
            
            .ColumnsHeader.Add "IDImballoVendita", "IDImballo", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imb.", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodicePedana", "CodicePedana", dgchar, True, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDCliente", "IDCliente", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "RV_PODataOrdinePadre", "Data ord.", dgDate, True, 1500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "RV_PONumeroOrdinePadre", "N° ord.", dgNumeric, True, 1500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "RV_PONumeroListaPrelievo", "N° lista", dgNumeric, True, 1500, dgAlignleft, True, True, False
            
            
            .ColumnsHeader.Add "Link_Doc_sezionale", "IDSezionale", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, False, 2500, dgAlignleft
            .ColumnsHeader.Add "Doc_data", "Data ord. ori.", dgDate, False, 1500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Doc_numero", "N° ord. ori.", dgNumeric, False, 1500, dgAlignleft, True, True, False
            
            
            
            .ColumnsHeader.Add "IDArticolo_Conferito", "IDArticolo_Conferito", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceArticolo_conferito", "Cod. Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Articolo_conferito", "Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDTipoLavorazione", "IDTipoLavorazione", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POCalibro", "IDRV_POCalibro", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Calibro", "Calibro", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POTipoCategoria", "IDRV_POTipoCategoria", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, True, 1700, dgAlignleft, True, True, False
                                                                            
                                                                            
            .ColumnsHeader.Add "CodiceTipoPedana", "Cod. tipo pedana", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceArticoloPedana", "Codice articolo pedana", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "ArticoloPedana", "Articolo pedana", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", dgInteger, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "AnnoProcesso", "Anno processo IV gamma", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NumeroProcesso", "Numero processo IV gamma", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NotaRigaOrdRaggr", "Raggrupp. ord.", dgchar, True, 2500, dgAlignleft, True, True, False
                                                                            
            Set cl = .ColumnsHeader.Add("ImportoUnitarioArticolo", "Importo Art.", dgDouble, False, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Sconto1", "% Sc. 1", dgDouble, False, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Sconto2", "% Sc. 2", dgDouble, False, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
               ' cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

            Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Importo Imb.", dgDouble, False, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("MerceInclusoImballo", "Imp. incl. Imb.", dgBoolean, False, 1000, dgAligncenter, True, True, False)
                cl.Editable = True
                cl.BackColor = vbYellow
                                                                            
            Set .Recordset = rsGrigliaSmistamento.Data
            .LoadUserSettings
            .Refresh
        End With
        
        CnDMT.CursorLocation = OLDCursor



        'If (Me.GrigliaSmistamento.Recordset.BOF) And (Me.GrigliaSmistamento.Recordset.EOF) Then
        '    Me.LabelLink2.Visible = False
        'Else
        '    Me.LabelLink2.Visible = True
        'End If
        
        If FraTotaliOrdineSmist.Visible = True Then
            GET_TOTALI_ORDINE_DA_SMISTARE
        End If
End Sub

Private Sub GET_CLIENTE_ORDINE_PRED()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POParametriUtenteOrd "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser


Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    LINK_CLIENTE_ORD_PRED = fnNotNullN(rs!IDClienteOrdineSmist)
    DATA_ORDINE_PRED = fnNotNull(rs!DataOrdineSmist)
    NUMERO_ORDINE_PRED = fnNotNullN(rs!NumeroOrdineSmist)
    LINK_ORDINE_PRED = fnNotNullN(rs!IDOggettoOrdineSmist)
    Me.CDAltroCliente.Load LINK_CLIENTE_ORD_PRED
    Me.txtDataSmistamento.Text = DATA_ORDINE_PRED
    Me.txtNumeroSmistamento.Value = NUMERO_ORDINE_PRED
    Me.txtNListaSmistamento.Value = 1
    
    rs.CloseResultset
    Set rs = Nothing
    
    FLAG_ASSEGNAZIONE_VELOCE = True
Else

    sSQL = "SELECT IDClienteCoop, DataOrdineCoop, NumeroOrdineCoop "
    sSQL = sSQL & "FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        LINK_CLIENTE_ORD_PRED = 0
        DATA_ORDINE_PRED = ""
        NUMERO_ORDINE_PRED = 0
        LINK_ORDINE_PRED = 0
        Me.CDAltroCliente.Load 0
        Me.txtDataSmistamento.Value = 0
        Me.txtNumeroSmistamento.Value = 0
        Me.txtNListaSmistamento.Value = 0
    Else
        LINK_CLIENTE_ORD_PRED = fnNotNullN(rs!IDClienteCoop)
        DATA_ORDINE_PRED = fnNotNull(rs!DataOrdineCoop)
        NUMERO_ORDINE_PRED = fnNotNullN(rs!NumeroOrdineCoop)
        LINK_ORDINE_PRED = GET_ESISTENZA_ORDINE(LINK_CLIENTE_ORD_PRED, DATA_ORDINE_PRED, NUMERO_ORDINE_PRED, 1)
        Me.CDAltroCliente.Load LINK_CLIENTE_ORD_PRED
        Me.txtDataSmistamento.Text = DATA_ORDINE_PRED
        Me.txtNumeroSmistamento.Value = NUMERO_ORDINE_PRED
        Me.txtNListaSmistamento.Value = 1
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    FLAG_ASSEGNAZIONE_VELOCE = False
End If

End Sub
Private Sub GET_CLIENTE_ORDINE_PREP()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POParametriUtenteOrd "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser


Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtIDOrdine.Value = fnNotNullN(rs!IDOggettoOrdinePrep)
    Me.cdCliente.Load fnNotNullN(rs!IDClienteOrdinePrep)
    Me.txtDataOrdine.Value = fnNotNullN(rs!DataOrdinePrep)
    Me.txtNumeroOrdine.Value = fnNotNullN(rs!NumeroOrdinePrep)
    'Me.txtNListaSmistamento.Value = 1
Else
    Me.txtIDOrdine.Value = 0
    Me.cdCliente.Load 0
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
    'Me.txtNListaSmistamento.Value = 0
End If

rs.CloseResultset
Set rs = Nothing

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
Private Sub AvviaAssegnazioneVeloce()
On Error GoTo ERR_AvviaAssegnazioneVeloce
Dim sSQL As String
Dim sSQL_Assegnazione As String
Dim rsOrd As DmtOleDbLib.adoResultset
Dim rs As ADODB.Recordset
Dim CodiceAssegnazioneVeloce As String
Dim Testo As String
Dim LINK_RIGA_CONFERIMENTO_VELOCE As Long
Dim LINK_PEDANA_VELOCE As Long



    If Me.CDAltroCliente.KeyFieldID = 0 Then
        MsgBox "Inserire il cliente dello smistamento", vbInformation, "Smistamento merce"
        Me.CDAltroCliente.SetFocus
        Exit Sub
    End If
    
    If Me.txtNumeroSmistamento.Value = 0 Then
        MsgBox "Inserire il numero dell'ordine dello smistamento", vbInformation, "Smistamento merce"
        Me.txtNumeroSmistamento.SetFocus
        Exit Sub
    End If
    
    If Me.txtDataSmistamento.Value = 0 Then
        MsgBox "Inserire la data dell'ordine dello smistamento", vbInformation, "Smistamento merce"
        Me.txtDataSmistamento.SetFocus
        Exit Sub
    End If

Select Case Mid(Me.txtAssegnazioneVeloce.Text, 1, 1)

    Case "V"
        LINK_RIGA_CONFERIMENTO_VELOCE = GET_LINK_CONFERIMENTO_RIGA(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
        Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(LINK_RIGA_CONFERIMENTO_VELOCE, 0, 0)
        
        If Len(Testo) > 0 Then
            ERRORE_EVASIONE = Testo
            frmErrore.Show vbModal
            Me.txtAssegnazioneVeloce.Text = ""
            Me.txtAssegnazioneVeloce.SetFocus
            Exit Sub
        End If
        If Me.chkAttitaMaschera.Value = vbUnchecked Then
            sSQL_Assegnazione = " AND CodiceLottoVendita=" & fnNormString(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
        Else

            COMANDO_SPEZZATUTA = 0
            COMANDO_RIPESATURA = 0
            LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO = GET_LINK_ASSEGNAZIONE_MERCE_VELOCE(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
            LINK_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtIDOrdine.Value
            LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO = Me.cdCliente.KeyFieldID
            NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtNumeroOrdine.Value
            DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO = Me.txtDataOrdine.Text
            
            cmdSu_Click
            
            Exit Sub
        End If
    Case "P"
        LINK_PEDANA_VELOCE = GET_LINK_PEDANA(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
        
        Testo = GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(0, 0, LINK_PEDANA_VELOCE)
        
        If Len(Testo) > 0 Then
            ERRORE_EVASIONE = Testo
            frmErrore.Show vbModal
            Me.txtAssegnazioneVeloce.Text = ""
            Me.txtAssegnazioneVeloce.SetFocus
            Exit Sub
        End If
        
        If Me.chkAttitaMaschera.Value = vbUnchecked Then
            sSQL_Assegnazione = " AND CodicePedana=" & fnNormString(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
        Else
            AVVIO_VELOCE_RIPESATURA = 1
            cmdRipesatura_Click
            AVVIO_VELOCE_RIPESATURA = 0
            Exit Sub
        End If
            
    Case "O"
        GET_ORDINE_DA_ASSEGNAZIONE_VELOCE Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text))
        Exit Sub
        
        
    Case Else
        Select Case Me.txtAssegnazioneVeloce.Text
            Case CODICE_CONFERMA_ORDINE
                GET_CONFERMA_ORDINE
                Me.txtAssegnazioneVeloce.Text = ""
                Me.txtAssegnazioneVeloce.SetFocus
                Exit Sub
            Case RIPRISTINA_PARAMETRI
                GET_CLIENTE_ORDINE_PRED
                GET_CLIENTE_ORDINE_PREP
                
                fnGrigliaAssegnazione
                fncGrigliaSmistamento
                
                Me.txtAssegnazioneVeloce.Text = ""
                Me.txtAssegnazioneVeloce.SetFocus
                Exit Sub
            Case CODICE_GESTIONE_ERRORI
                GET_CLIENTE_ORDINE_PRED
                GET_CLIENTE_ORDINE_PREP
                Me.txtAssegnazioneVeloce.Text = ""
                Me.txtAssegnazioneVeloce.SetFocus
                Exit Sub
            Case Else
                ERRORE_EVASIONE = "CODICE ERRATO"
                frmErrore.Show vbModal
                Me.txtAssegnazioneVeloce.Text = ""
                Me.txtAssegnazioneVeloce.SetFocus
                Exit Sub
        End Select
End Select

If Me.txtIDOrdine.Value = 0 Then
    ERRORE_EVASIONE = "INSERIRE L'ORDINE DA PREPARARE"
    frmErrore.Show vbModal
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus
    Exit Sub
End If

sSQL = "SELECT RV_POAssegnazioneMerce.*, RV_POCaricoMerceRighe.IDArticolo AS IDArticoloConferito, "
sSQL = sSQL & "RV_POCaricoMerceRighe.CodiceArticolo AS CodiceArticolo_Conferito, RV_POCaricoMerceRighe.Articolo AS Articolo_Conferito,"
sSQL = sSQL & "ValoriOggettoPerTipo000F.Doc_ordine_chiuso, Anagrafica.Anagrafica AS Cliente, Anagrafica.Nome AS NomeCliente, "
sSQL = sSQL & "ValoriOggettoPerTipo000F.Doc_data, ValoriOggettoPerTipo000F.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON ValoriOggettoPerTipo000F.Doc_numero = RV_POAssegnazioneMerce.NumeroOrdine AND "
sSQL = sSQL & "ValoriOggettoPerTipo000F.Doc_data = RV_POAssegnazioneMerce.DataOrdine LEFT OUTER JOIN "
sSQL = sSQL & "Anagrafica ON RV_POAssegnazioneMerce.IDCliente = Anagrafica.IDAnagrafica LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta ON "
sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE (ValoriOggettoPerTipo000F.Doc_ordine_chiuso = 0) "

If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDCliente=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_Numero=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_data=" & fnNormDate(Me.txtDataSmistamento.Text)
End If

sSQL = sSQL & sSQL_Assegnazione

Set rsOrd = CnDMT.OpenResultset(sSQL)

If rsOrd.EOF = True Then
    ERRORE_EVASIONE = "CODICE INESISTENTE"
    frmErrore.Show vbModal
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus
    
    rsOrd.CloseResultset
    Set rsOrd = Nothing
    Exit Sub
End If

While Not rsOrd.EOF
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & fnNotNullN(rsOrd!IDRV_POAssegnazioneMerce)
    
    Set rs = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenDynamic, adLockPessimistic
    
    
    While Not rs.EOF
        If GET_CONTROLLO_RIGA_SMISTAMENTO(fnNotNullN(rs!IDRV_POAssegnazioneMerce)) = True Then
                        
            rs!IDOggettoOrdine = Me.txtIDOrdine.Value
            rs!IDCliente = Me.cdCliente.KeyFieldID
            rs!NumeroOrdine = Me.txtNumeroOrdine.Value
            rs!DataOrdine = Me.txtDataOrdine.Text
            rs!NumeroListaPrelievo = Me.txtNListaPrelievo.Value
            rs!IDOggettoOrdinePadre = Me.txtIDOrdinePadre.Value
            
            If PAR_NonCalcImpDaAssVeloce = 1 Then
                If PREZZI_ARTICOLI_DA_ORDINE = 0 Then
                    GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                    GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                    rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
                Else
                    If (GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDImballoVendita), Me.txtIDOrdinePadre.Value, rs) = False) Then
                        GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                        GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                        rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
                    Else
                        If RETURN_SEL_PREZZO_IMB_DA_ORD = 0 Then
                            GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
                            rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
                        End If
                    End If
                End If
            Else
                rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
            End If
            
            'GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
            'GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rs!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), rs
            'rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.txtIDOrdinePadre.Value)
            
            rs.Update
        End If
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
'    If PREZZI_ARTICOLI_DA_ORDINE = 1 Then
'        Testo = "ATTENZIONE!!!" & vbCrLf
'        Testo = Testo & "È stato impostato di prelevare gli importi dall'ordine, pertanto prima di procedere alla conferma dell'ordine eseguire la prezzatura veloce dal comando 'PREZZATURA DA ORDINE'"
'
'        MsgBox Testo, vbInformation, "Prezzatura da ordine"
'    End If
    
    
    DoEvents
    
    fnGrigliaAssegnazione
    'fncGrigliaOrdine
    fncGrigliaSmistamento
    
    DoEvents

rsOrd.MoveNext
Wend
Screen.MousePointer = 0
Me.txtAssegnazioneVeloce.Text = ""

rsOrd.CloseResultset
Set rsOrd = Nothing


PlaySound App.Path & "\ding.wav", 0, SND_FILENAME Or SND_ASYNC


GET_RIEPILOGO_ORDINE Me.txtIDOrdine.Value

GET_TOTALI_ORDINE_DA_PREPARARE Me.txtIDOrdine.Value
GET_TOTALI_ORDINE_DA_SMISTARE
Exit Sub
ERR_AvviaAssegnazioneVeloce:
    MsgBox Err.Description, vbCritical, "AvviaAssegnazioneVeloce"
End Sub

Private Function GET_CONTROLLO_RIGA_SMISTAMENTO(IDAssegnazioneMerce As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Ordine_Smistamento As Long


Link_Ordine_Smistamento = GET_ESISTENZA_ORDINE(Me.CDAltroCliente.KeyFieldID, Me.txtDataSmistamento.Text, Me.txtNumeroSmistamento.Value, Me.txtNListaSmistamento.Value)
sSQL = "SELECT IDOggettoOrdine "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_RIGA_SMISTAMENTO = False
Else
    If fnNotNullN(rs!IDOggettoOrdine) = Link_Ordine_Smistamento Then
        GET_CONTROLLO_RIGA_SMISTAMENTO = True
    Else
        GET_CONTROLLO_RIGA_SMISTAMENTO = False
    End If
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AGGIORNA_ORDINE_CLIENTE(IDDestinazione As Long, IDVettore As Long, IDTipoDestinazione As Long)
'On Error GoTo ERR_FN_CREA_ORDINE
Dim ObjDoc As DmtDocs.cDocument
Dim cDefault As Collection


If Not (ObjDoc Is Nothing) Then
    Set ObjDoc = Nothing
End If
Screen.MousePointer = 11
Set ObjDoc = New DmtDocs.cDocument
    Set ObjDoc.Connection = CnDMT
    ObjDoc.ReadWithTO Me.txtIDOrdine.Value, 15

    If IDDestinazione > 0 Then
        ObjDoc.ReadDataFromCliFoSite IDDestinazione, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    ObjDoc.Field "RV_PODataArrivoMerce", Me.txtDataArrivoMerce.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POOraArrivoMerce", Me.txtOraArrivoMerce.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "Link_Doc_spedizione", IDTipoDestinazione, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
    If IDTipoDestinazione = 3 Then
        If IDVettore > 0 Then
            ObjDoc.ReadDataFromCarrier IDVettore, MainCarrier, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
    End If
    
    ObjDoc.Field "Doc_data_prevista_evasione", Me.txtDataPartenza.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Doc_data_presso_nom", Me.txtDataOrdineCliente.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Doc_numero_presso_nom", Me.txtNumeroOrdineCliente.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Doc_annotazioni_variazio", Mid(Me.txtAnnotazioniOrdine.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POAnnotazioniInterna", Mid(Me.txtAnnotazioniInterna.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_PODescrizioneCorpoDocEv", Mid(Me.txtDescrizioneRigaDoc.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDLuogoPresaMerce", Me.cboLuogoPresaMerce.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_PODataArrivoMerceLuogo", Me.txtDataArrivoMerceL.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POOraArrivoMerceLuogo", Me.txtOraArrivoMerceL.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDTrasportatoreSuccessivo", Me.cboVettoreSuccessivo.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_IVA", Me.cboIvaCliente.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_lettera_intento", Me.txtIDLetteraIntento.Value, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Update

    Set ObjDoc = Nothing
    
Screen.MousePointer = 0
Exit Sub
ERR_FN_CREA_ORDINE:
    MsgBox Err.Description, vbCritical, "FN_CREA_ORDINE"
    Screen.MousePointer = 0
End Sub
Function TimerProc1()
    fnGrigliaOrdiniDaElaborare
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

Private Sub StampaDocumento(TipoScelta As Integer)
On Error GoTo ERR_StampaDocumento
Dim IDReport As Long
Dim IDTipoOggettoPRG As Long

Set oReport = New dmtReportLib.dmtReport
    'parametri di accesso al database SQL Server
    Set oReport.Connection = TheApp.Database.Connection
    oReport.Password = TheApp.Password
    oReport.User = TheApp.User

'Imposta l'idfiliale di appartenenza del documento da stampare
    oReport.BranchID = TheApp.Branch  'IDFiliale

'Imposta l'identificativo del tipo di documento
    IDTipoOggettoPRG = fncIDTipoOggettoPrg
    
    oReport.DocTypeID = IDTipoOggettoPRG
    
    If TipoScelta = 0 Then
        oReport.Where = "IDOggettoOrdine=" & Me.txtIDOrdine.Value
    Else
        oReport.Where = GET_SQL_PER_STAMPA
    End If
    
    

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


If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCliente=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND NumeroOrdine=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine=" & fnNormDateString(Me.txtDataSmistamento.Text)
End If

If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
End If

If Len(Me.txtCodicePedana.Text) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND CodicePedana=" & fnNormString(Me.txtCodicePedana.Text)
End If

If Me.CDSocio.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If

GET_SQL_PER_STAMPA = sSQL_GENERALE

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
Private Function CHECK_ABILITAZIONE_DIAMANTE() As Boolean
Dim I As Integer
Dim swChk As DmtSwChk.SwCheck

Set swChk = New DmtSwChk.SwCheck

Set swChk.Connection = CnDMT.InternalConnection
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
Private Function GET_ORDINE_DA_ASSEGNAZIONE_VELOCE(CodiceOrdine As String) As Long
On Error GoTo ERR_GET_ORDINE_DA_ASSEGNAZIONE_VELOCE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsConf As DmtOleDbLib.adoResultset
Dim IDUtenteBlocco As Long
Dim Testo As String
Dim LINK_TIPO_OGGETTO_ORDINE_COOP As Long
Dim LINK_TIPO_OGGETTO_ORDINE As Long
Dim LINK_FUNZIONE_ORDINE_COOP As Long
Dim LINK_FUNZIONE_OGGETTO_ORDINE As Long


'''''VARIABILI PER SCORPORARE IL CODICE ORDINE PASSATO''''''''''''''''''
Dim STRINGA_DATA_ORDINE As String
Dim GIORNO_ORDINE As String
Dim MESE_ORDINE As String
Dim ANNO_ORDINE As String
Dim STRINGA_NUMERO_ORDINE As String
Dim DATA_DOCUMENTO As String
Dim NUMERO_DOCUMENTO As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

BLOCCA_ORDINE Me.txtIDOrdine.Value, 0


LINK_TIPO_OGGETTO_ORDINE_COOP = fnGetTipoOggetto("RV_POOrdineL")
LINK_FUNZIONE_ORDINE_COOP = GET_FUNZIONE(LINK_TIPO_OGGETTO_ORDINE_COOP)
LINK_TIPO_OGGETTO_ORDINE = 15
LINK_FUNZIONE_OGGETTO_ORDINE = 128

STRINGA_DATA_ORDINE = Mid(CodiceOrdine, 1, 8)
GIORNO_ORDINE = Mid(STRINGA_DATA_ORDINE, 1, 2)
MESE_ORDINE = Mid(STRINGA_DATA_ORDINE, 3, 2)
ANNO_ORDINE = Mid(STRINGA_DATA_ORDINE, 5, 4)

STRINGA_NUMERO_ORDINE = Mid(CodiceOrdine, 9, Len(CodiceOrdine))

DATA_DOCUMENTO = GIORNO_ORDINE & "/" & MESE_ORDINE & "/" & ANNO_ORDINE
NUMERO_DOCUMENTO = Val(STRINGA_NUMERO_ORDINE)

sSQL = "SELECT ValoriOggettoPerTipo000F.IDOggetto, ValoriOggettoPerTipo000F.Doc_numero, ValoriOggettoPerTipo000F.Doc_data, "
sSQL = sSQL & "ValoriOggettoPerTipo000F.Link_nom_anagrafica, ValoriOggettoPerTipo000F.Doc_ordine_chiuso, "
sSQL = sSQL & "ValoriOggettoPerTipo000F.IDTipoOggetto, Oggetto.IDFunzione "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Doc_data=" & fnNormDate(DATA_DOCUMENTO)
sSQL = sSQL & " AND Doc_numero=" & fnNormNumber(NUMERO_DOCUMENTO)

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    ''''CONTROLLO DELL'ORDINE BLOCCATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    IDUtenteBlocco = CONTROLLO_ORDINE_BLOCCATO(fnNotNullN(rs!IDOggetto), TheApp.IDUser)
    
    If IDUtenteBlocco > 0 Then
        rs.CloseResultset
        Set rs = Nothing
        
        Testo = ""
        Testo = Testo & "L'ordine risulta bloccato dall'utente " & GET_UTENTE(IDUtenteBlocco)
        ERRORE_EVASIONE = Testo
        frmErrore.Show vbModal
        Me.txtAssegnazioneVeloce.Text = ""
        Me.txtAssegnazioneVeloce.SetFocus
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    '''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE STANDARD''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM Semaforo "
    sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
    sSQL = sSQL & " AND IDTipoOggetto=" & fnNotNullN(rs!IDTipoOggetto)
    sSQL = sSQL & "  AND IDFunzione=" & fnNotNullN(rs!IDFunzione)
    
    Set rsConf = CnDMT.OpenResultset(sSQL)
    If Not rsConf.EOF Then
        rsConf.CloseResultset
        Set rsConf = Nothing
        ERRORE_EVASIONE = "ATTENZIONE!!!!" & vbCrLf
        ERRORE_EVASIONE = ERRORE_EVASIONE & "l'ordine risulta aperto dall'utente " & GET_UTENTE(fnNotNullN(rs!IDUtente))
        frmErrore.Show vbModal
        Me.txtIDOrdine.Value = 0
        Me.txtAssegnazioneVeloce.Text = ""
        Me.txtAssegnazioneVeloce.SetFocus
        Exit Function
    End If
    
    rsConf.CloseResultset
    Set rsConf = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE GREEN TOP''''''''''''''''''''''''''''''''''
    sSQL = "SELECT * FROM Semaforo "
    sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
    sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO_ORDINE_COOP
    sSQL = sSQL & "  AND IDFunzione=" & LINK_FUNZIONE_ORDINE_COOP
    
    Set rsConf = CnDMT.OpenResultset(sSQL)
    If Not rsConf.EOF Then
        ERRORE_EVASIONE = "ATTENZIONE!!!!" & vbCrLf
        ERRORE_EVASIONE = ERRORE_EVASIONE & "l'ordine risulta aperto dall'utente " & GET_UTENTE(fnNotNullN(rsConf!IDUtente))

        rsConf.CloseResultset
        Set rsConf = Nothing
        frmErrore.Show vbModal
        Me.txtIDOrdine.Value = 0
        Me.txtAssegnazioneVeloce.Text = ""
        Me.txtAssegnazioneVeloce.SetFocus
        Exit Function
    End If
    
    rsConf.CloseResultset
    Set rsConf = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''CONTROLLO DELL'ORDINE CHIUSO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If fnNotNullN(rs!Doc_ordine_chiuso) = 1 Then
        rs.CloseResultset
        Set rs = Nothing
        
        ERRORE_EVASIONE = "ATTENZIONE!!!!" & vbCrLf
        ERRORE_EVASIONE = ERRORE_EVASIONE & "l'ordine risulta chiuso"
        frmErrore.Show vbModal
        Me.txtIDOrdine.Value = 0
        Me.txtAssegnazioneVeloce.Text = ""
        Me.txtAssegnazioneVeloce.SetFocus
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    '''''''''''''''''''''''CONTROLLO SE ORDINE CONFERMATO'''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT IDOggetto FROM RV_POTMPEvasioneOrdini "
    sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
    
    Set rsConf = CnDMT.OpenResultset(sSQL)
    If Not rsConf.EOF Then
        rsConf.CloseResultset
        Set rsConf = Nothing
        ERRORE_EVASIONE = "ATTENZIONE!!!!" & vbCrLf
        ERRORE_EVASIONE = ERRORE_EVASIONE & "l'ordine risulta confermato"
        frmErrore.Show vbModal
        Me.txtIDOrdine.Value = 0
        Me.txtAssegnazioneVeloce.Text = ""
        Me.txtAssegnazioneVeloce.SetFocus
        Exit Function
    End If
    
    rsConf.CloseResultset
    Set rsConf = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    
    BLOCCA_ORDINE fnNotNullN(rs!IDOggetto), TheApp.IDUser
    
    Me.txtIDOrdine.Value = fnNotNullN(rs!IDOggetto)
    Me.cdCliente.Load fnNotNullN(rs!link_nom_anagrafica)
    Me.txtDataOrdine.Value = fnNotNullN(rs!Doc_data)
    Me.txtNumeroOrdine.Value = fnNotNullN(rs!Doc_numero)
Else
    Me.txtIDOrdine.Value = 0
    Me.cdCliente.Load 0
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
End If



rs.CloseResultset
Set rs = Nothing

fnGrigliaAssegnazione

Me.txtAssegnazioneVeloce.Text = ""
Me.txtAssegnazioneVeloce.SetFocus

Exit Function
ERR_GET_ORDINE_DA_ASSEGNAZIONE_VELOCE:
    ERRORE_EVASIONE = Err.Description
    frmErrore.Show vbModal
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus
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
Private Sub GET_PARAMETRI_EVASIONE_ORDINI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    CODICE_CONFERMA_ORDINE = ""
    RIPRISTINA_PARAMETRI = ""
    CODICE_GESTIONE_ERRORI = ""
    DISATTIVA_SCALATA_COMM_TRASP = 0
Else
    CODICE_CONFERMA_ORDINE = fnNotNull(rs!CodiceConfermaOrdine)
    RIPRISTINA_PARAMETRI = fnNotNull(rs!CodiceRispristinoOperazioni)
    CODICE_GESTIONE_ERRORI = fnNotNull(rs!CodiceAnnullaOperazione)
    DISATTIVA_SCALATA_COMM_TRASP = fnNotNullN(rs!DisattivaScalataCommTipoPedana)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_CONTROLLO_BLOCCO_ASSEGNAZIONI(IDRigaConferimento As Long, IDAssegnazioneMerce As Long, IDPedana As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsConf As DmtOleDbLib.adoResultset
Dim Testo As String


GET_CONTROLLO_BLOCCO_ASSEGNAZIONI = ""

Dim LINK_TIPO_OGGETTO_ASSEGNAZIONE As Long
Dim LINK_FUNZIONE_ASSEGNAZIONE As Long

LINK_TIPO_OGGETTO_ASSEGNAZIONE = fnGetTipoOggetto("RV_POAssegnazioneMerce")
LINK_FUNZIONE_ASSEGNAZIONE = GET_FUNZIONE(LINK_TIPO_OGGETTO_ASSEGNAZIONE)

If IDPedana > 0 Then
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        sSQL = "SELECT * FROM Semaforo "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
        sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO_ASSEGNAZIONE
        sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE_ASSEGNAZIONE
        
        Set rsConf = CnDMT.OpenResultset(sSQL)
        
        If Not rsConf.EOF Then
            rsConf.CloseResultset
            Set rsConf = Nothing
            rs.CloseResultset
            Set rs = Nothing
            Testo = "ATTENZIONE!!!!" & vbCrLf
            Testo = Testo & "Alcune lavorazioni della pedana sono bloccate da altri utenti"
            
            GET_CONTROLLO_BLOCCO_ASSEGNAZIONI = Testo
            
            Exit Function
        End If
        
        rsConf.CloseResultset
        Set rsConf = Nothing
        
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
    
Else
    sSQL = "SELECT * FROM Semaforo "
    sSQL = sSQL & "WHERE IDOggetto=" & IDRigaConferimento
    sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO_ASSEGNAZIONE
    sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE_ASSEGNAZIONE
    
    Set rsConf = CnDMT.OpenResultset(sSQL)
    If Not rsConf.EOF Then
        
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "La lavorazione è bloccata dall'utente " & GET_UTENTE(fnNotNullN(rsConf!IDUtente))
        
        GET_CONTROLLO_BLOCCO_ASSEGNAZIONI = Testo
    
    End If
    
    rsConf.CloseResultset
    Set rsConf = Nothing

End If



sSQL = ""

End Function
Private Function GET_LINK_CONFERIMENTO_RIGA(CodiceVendita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE CodiceLottoVendita=" & fnNormString(CodiceVendita)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CONFERIMENTO_RIGA = 0
Else
    GET_LINK_CONFERIMENTO_RIGA = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ASSEGNAZIONE_MERCE_VELOCE(CodiceVendita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POAssegnazioneMerce "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & " WHERE CodiceLottoVendita=" & fnNormString(CodiceVendita)
sSQL = sSQL & " AND "


sSQL = "SELECT IDRV_POAssegnazioneMerce FROM RV_POIEOrdineSmistamento "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Doc_ordine_chiuso=0"
sSQL = sSQL & " AND Qta_UM>0"
sSQL = sSQL & " AND CodiceLottoVendita=" & fnNormString(CodiceVendita)

If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDCliente=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
End If
If Me.txtNListaSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroListaPrelievo=" & txtNListaSmistamento.Value
End If

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ASSEGNAZIONE_MERCE_VELOCE = 0
Else
    GET_LINK_ASSEGNAZIONE_MERCE_VELOCE = fnNotNullN(rs!IDRV_POAssegnazioneMerce)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Function GET_LINK_PEDANA(CodicePedana As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POPedana "
sSQL = sSQL & "WHERE Codice=" & fnNormString(CodicePedana)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PEDANA = 0
Else
    GET_LINK_PEDANA = fnNotNullN(rs!IDRV_POPedana)
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
Private Function GET_LINK_LISTINO_AZIENDA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM PersonalizzazionePerFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LISTINO_AZIENDA = 0
Else
    GET_LINK_LISTINO_AZIENDA = fnNotNullN(rs!IDListinoDiBase)
End If

rs.CloseResultset
Set rs = Nothing
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
    IDArticoloPadre = GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
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

'''''IMPORTO UNITARIO DAL LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''''''
If ImportoUnitario = 0 Then
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
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

GET_PREZZO_UNITARIO_ARTICOLO = ImportoUnitario
Exit Function
ERR_GET_PREZZO_UNITARIO_ARTICOLO:
    GET_PREZZO_UNITARIO_ARTICOLO = 0
    
End Function
Private Function GET_CONTROLLO_IMPORTI_A_ZERO(IDOggettoOrdine As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_IMPORTI_A_ZERO = False


''''CONTROLLO DEGLI IMPORTI UNITARI ARTICOLO A ZERO''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDRV_POAssegnazioneMerce "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & " WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & " AND ((ImportoUnitarioArticolo=0) OR (ImportoUnitarioArticolo IS NULL))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_IMPORTI_A_ZERO = False
Else
    GET_CONTROLLO_IMPORTI_A_ZERO = True
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GET_CONTROLLO_IMPORTI_A_ZERO = False Then
    ''''CONTROLLO DEGLI IMPORTI UNITARI IMBALLO A ZERO''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT IDRV_POAssegnazioneMerce "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & " WHERE IDOggettoOrdine=" & IDOggettoOrdine
    sSQL = sSQL & " AND ((ImportoUnitarioImballo=0) OR (ImportoUnitarioImballo IS NULL))"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLO_IMPORTI_A_ZERO = False
    Else
        GET_CONTROLLO_IMPORTI_A_ZERO = True
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End If



End Function
Private Sub GET_CONFERMA_ORDINE()
Dim sSQL As String
Dim rsConf As DmtOleDbLib.adoResultset
Dim Testo As String

If Me.txtIDOrdine.Value = 0 Then
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus
    Exit Sub
End If

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE STANDARD''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & 15
sSQL = sSQL & "  AND IDFunzione=" & 128

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    ERRORE_EVASIONE = "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente)
    
    rsConf.CloseResultset
    Set rsConf = Nothing

    frmErrore.Show vbModal
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus

    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE GREEN TOP'''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdine.Value
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POOrdineL")
sSQL = sSQL & "  AND IDFunzione=" & GET_FUNZIONE(fnGetTipoOggetto("RV_POOrdineL"))

Set rsConf = CnDMT.OpenResultset(sSQL)
If Not rsConf.EOF Then
    ERRORE_EVASIONE = "L'ordine è bloccato dall'utente " & GET_UTENTE(rsConf!IDUtente)
    rsConf.CloseResultset
    Set rsConf = Nothing
    frmErrore.Show vbModal
    Me.txtAssegnazioneVeloce.Text = ""
    Me.txtAssegnazioneVeloce.SetFocus
    
    Exit Sub
End If

rsConf.CloseResultset
Set rsConf = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'If GET_CONTROLLO_IMPORTI_A_ZERO(Me.txtIDOrdine.Value) = True Then
'    ERRORE_EVASIONE = "ATTENZIONE!!!!!" & vbCrLf
'    ERRORE_EVASIONE = ERRORE_EVASIONE & "Nell'ordine confermato risultano importi articolo o importi imballo a zero " & vbCrLf
'
'    frmErrore.Show vbModal
'    Me.txtAssegnazioneVeloce.Text = ""
'    Me.txtAssegnazioneVeloce.SetFocus
'
'    Exit Sub
'End If



KillTimer Me.hwnd, 0

rsGrigliaAss.UpdateBatch


If GET_ESISTENZA_ORDINE_DA_ELABORARE = False Then
    sSQL = "INSERT INTO RV_POTMPEvasioneOrdini ("
    sSQL = sSQL & "NumeroRiga, IDOggetto, IDAzienda, NumeroOrdine, DataOrdine, DaRegistrare, "
    sSQL = sSQL & "IDCliente, IDSitoPerAnagrafica, Cliente, SitoPerAnagrafica, "
    sSQL = sSQL & "IDVettore, Vettore, DescrizioneCorpoDocEv, IDLuogoPresaMerce, IDVettoreSuccessivo, InFatturazione ) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_POTMPEvasioneOrdini", "NumeroRiga") & ", "
    sSQL = sSQL & Me.txtIDOrdine.Value & ", "
    sSQL = sSQL & TheApp.IDFirm & ","
    sSQL = sSQL & Me.txtNumeroOrdine.Value & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataOrdine.Text) & ", "
    sSQL = sSQL & fnNormBoolean(0) & ", "
    sSQL = sSQL & Me.cdCliente.KeyFieldID & ", "
    sSQL = sSQL & Me.cboDestinazione.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cdCliente.Description) & ", "
    sSQL = sSQL & fnNormString(Me.cboDestinazione.Text) & ", "
    sSQL = sSQL & Me.cboVettore.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cboVettore.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtDescrizioneRigaDoc.Text) & ", "
    sSQL = sSQL & Me.cboLuogoPresaMerce.CurrentID & ", "
    sSQL = sSQL & Me.cboVettoreSuccessivo.CurrentID & ", "
    sSQL = sSQL & fnNormBoolean(0) & ")"
    
    CnDMT.Execute sSQL
    
End If

AGGIORNA_ORDINE_CLIENTE Me.cboDestinazione.CurrentID, Me.cboVettore.CurrentID, Me.cboTipoTrasporto.CurrentID

BLOCCA_ORDINE Me.txtIDOrdine.Value, 0

LINK_ORDINE = 0
Me.txtIDOrdine.Value = 0
Me.cdCliente.Load 0
Me.txtNumeroOrdine.Value = 0
Me.txtDataOrdine.Value = 0

fnGrigliaOrdiniDaElaborare
fnGrigliaAssegnazione

Me.txtAssegnazioneVeloce.Text = ""
Me.txtAssegnazioneVeloce.SetFocus


lngTimerID = SetTimer(Me.hwnd, 0, 5000, AddressOf TimerProc)
End Sub

Private Sub GET_RIEPILOGO_ORDINE(IDOrdine As Long)
On Error Resume Next
If frmRiepilogoOrdine.Visible = True Then
    frmRiepilogoOrdine.txtNumeroPedana.Text = frmRiepilogoOrdine.GET_NUMERO_PEDANE_ORDINE(FrmMain.txtIDOrdine.Value)
    frmRiepilogoOrdine.txtNumeroColli.Text = frmRiepilogoOrdine.GET_NUMERO_COLLI_ORDINE(FrmMain.txtIDOrdine.Value)
End If

Me.SetFocus
End Sub
Private Function GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine As Long, PesoPedane As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_PEDANE_LAVORATE = 0


sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & " AND ((PreConferimento=0) OR (PreConferimento IS NULL)) "
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + 1
    PesoPedane = PesoPedane + GET_TROVA_PESO_PEDANA(fnNotNullN(rs!IDRV_POPedana))
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function



Private Function GET_TOTALI_ORDINE_DA_PREPARARE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim PesoPedana As Double
Dim PesoLordo As Double

PesoPedana = 0


sSQL = "SELECT SUM(Colli) as NumeroColli,"
sSQL = sSQL & "SUM(PesoLordo) as PesoLordo, "
sSQL = sSQL & "SUM(PesoNetto) as PesoNetto, "
sSQL = sSQL & "SUM(Tara) as Tara, "
sSQL = sSQL & "SUM(Pezzi) as Pezzi "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & " AND ((PreConferimento=0) OR (PreConferimento IS NULL)) "
Set rs = CnDMT.OpenResultset(sSQL)
    
If Not rs.EOF Then
    PesoLordo = fnNotNullN(rs!PesoLordo)
    Me.txtTotaleColliOrdPrep.Text = FormatNumber(fnNotNullN(rs!NumeroColli), 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdPrep.Text = FormatNumber(fnNotNullN(rs!PesoLordo), 2, , , vbTrue)
    Me.txtTotaleTaraOrdPrep.Text = FormatNumber(fnNotNullN(rs!Tara), 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdPrep.Text = FormatNumber(fnNotNullN(rs!PesoNetto), 2, , , vbTrue)
    Me.txtTotalePezziOrdPrep.Text = FormatNumber(fnNotNullN(rs!Pezzi), 2, , , vbTrue)
Else
    PesoLordo = 0
    Me.txtTotaleColliOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotaleTaraOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePezziOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
End If

rs.CloseResultset
Set rs = Nothing

Me.txtTotalePedaneOrdPrep.Text = FormatNumber(GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine, PesoPedana), 2, , , vbTrue)
Me.txtTotalePesoOrdPrep.Text = FormatNumber(PesoPedana + PesoLordo, 2, , , vbTrue)

End Function

Private Sub GET_TOTALI_ORDINE_DA_SMISTARE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
        
sSQL = "SELECT SUM(RV_POAssegnazioneMerce.Colli) as NumeroColli,"
sSQL = sSQL & "SUM(RV_POAssegnazioneMerce.PesoLordo) as PesoLordo, "
sSQL = sSQL & "SUM(RV_POAssegnazioneMerce.PesoNetto) as PesoNetto, "
sSQL = sSQL & "SUM(RV_POAssegnazioneMerce.Tara) as Tara, "
sSQL = sSQL & "SUM(RV_POAssegnazioneMerce.Pezzi) as Pezzi "
sSQL = sSQL & "FROM Oggetto INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000F ON Oggetto.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto AND "
sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo000F.IDTipoOggetto RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON Oggetto.IDOggetto = RV_POAssegnazioneMerce.IDOggettoOrdine LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta ON "
sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE Oggetto.IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_ordine_chiuso = 0 "

If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Link_nom_anagrafica=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
End If

If Me.txtNListaSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PONumeroListaPrelievo=" & Me.txtNListaSmistamento.Value
End If


If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDArticolo=" & Me.CDArticolo.KeyFieldID
End If

If Len(Me.txtCodicePedana.Text) > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.CodicePedana=" & fnNormString(Me.txtCodicePedana.Text)
End If

If Me.CDSocio.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If
If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
    sSQL = sSQL & " AND CodiceLottoVendita LIKE " & fnNormString(Me.txtLottoVendita.Text)
End If
If Len(Trim(Me.txtRaggrOrdine.Text)) > 0 Then
    sSQL = sSQL & " AND NotaRigaOrdRaggr LIKE " & fnNormString(Me.txtRaggrOrdine.Text)
End If

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtTotaleColliOrdSmist.Text = FormatNumber(fnNotNullN(rs!NumeroColli), 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdSmist.Text = FormatNumber(fnNotNullN(rs!PesoLordo), 2, , , vbTrue)
    Me.txtTotaleTaraOrdSmist.Text = FormatNumber(fnNotNullN(rs!Tara), 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdSmist.Text = FormatNumber(fnNotNullN(rs!PesoNetto), 2, , , vbTrue)
    Me.txtTotalePezziOrdSmist.Text = FormatNumber(fnNotNullN(rs!Pezzi), 2, , , vbTrue)
Else
    Me.txtTotaleColliOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotaleTaraOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePezziOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
End If


rs.CloseResultset
Set rs = Nothing

Me.txtTotalePedaneOrdSmist.Text = FormatNumber(GET_NUMERO_PEDANE_ORD_SMIST, 2, , , vbTrue)

End Sub

Private Function GET_NUMERO_PEDANE_ORD_SMIST() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_PEDANE_ORD_SMIST = 0

sSQL = "SELECT RV_POAssegnazioneMerce.IDRV_POPedana "
sSQL = sSQL & "FROM Oggetto INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000F ON Oggetto.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto AND "
sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo000F.IDTipoOggetto RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON Oggetto.IDOggetto = RV_POAssegnazioneMerce.IDOggettoOrdine LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta ON "
sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE Oggetto.IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Doc_ordine_chiuso = 0 "

If Me.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Link_nom_anagrafica=" & Me.CDAltroCliente.KeyFieldID
End If

If Me.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
End If

If Me.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
End If

If Me.txtNListaSmistamento.Value > 0 Then
    sSQL = sSQL & " AND ValoriOggettoPerTipo000F.RV_PONumeroListaPrelievo=" & Me.txtNListaSmistamento.Value
End If

If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDArticolo=" & Me.CDArticolo.KeyFieldID
End If

If Len(Me.txtCodicePedana.Text) > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.CodicePedana=" & fnNormString(Me.txtCodicePedana.Text)
End If

If Me.CDSocio.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If

sSQL = sSQL & " GROUP BY RV_POAssegnazioneMerce.IDRV_POPedana"

Set rs = CnDMT.OpenResultset(sSQL)


While Not rs.EOF
    If (fnNotNullN(rs!IDRV_POPedana) > 0) Then
        GET_NUMERO_PEDANE_ORD_SMIST = GET_NUMERO_PEDANE_ORD_SMIST + 1
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PARAMETRI_FILIALI_PER_TOTALI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT VisualizzaTotaliOrdinePrep, VisualizzaTotaliOrdineSmist "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & 0

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    FLAG_VIS_TOT_ORD_PREP = 0
    FLAG_VIS_TOT_ORD_SMIST = 0
Else
    FLAG_VIS_TOT_ORD_PREP = Abs(fnNotNullN(rs!VisualizzaTotaliOrdinePrep))
    FLAG_VIS_TOT_ORD_SMIST = Abs(fnNotNullN(rs!VisualizzaTotaliOrdineSmist))
End If

rs.CloseResultset
Set rs = Nothing

If FLAG_VIS_TOT_ORD_PREP = 1 Then
    cmdTotaliOrdine_Click
End If
If FLAG_VIS_TOT_ORD_SMIST = 1 Then
    cmdTotaliOrdSmist_Click
End If
End Sub
Private Function GET_PREZZO_IMBALLO_INCLUSO_2(IDArticolo As Long, IDCliente As Long) As Long
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
'    IDArticoloPadre = IDArticolo 'GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
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
'                GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rs!RV_POImportoImballoInArticolo)
'            End If
'
'            rs.CloseResultset
'            Set rs = Nothing
'            Exit Function
'        End If
'    End If
'End If
'

If GET_PREZZO_IMBALLO_INCLUSO_2 = 0 Then
    sSQL = "SELECT PrezzoInclusoImballo "
    sSQL = sSQL & "FROM RV_POConfigurazioneClienteImb "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticolo
    
    
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
Private Sub GET_PREZZI_DA_ORDINE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezziArticoloDaOrdine, PrezziImballoDaOrdine, PrezzoInclusoImballoDaOrdine, "
sSQL = sSQL & "AbilitaCategoriaImportoOrdine, AbilitaCalibroImportoOrdine, NonAbilitareImportoImballo, "
sSQL = sSQL & "NonCalcImpDaAssVeloce, CalcImpAConfOrd, NonVisMsgImpZeroConfOrd, AssNewPedDaAssSingola, NonCalcPrezzoDueRefArtOrd, "
sSQL = sSQL & "AttivaCommissioniDaOrdine, RicalcolaCommTipoPedDaOrdInEvasione, VisElencoRigheOrdineSeNonTroviAssociazione "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & 0

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    PREZZI_ARTICOLI_DA_ORDINE = 0
    PREZZI_IMBALLI_DA_ORDINE = 0
    PREZZO_INCLUSO_IMBALLO_DA_ORDINE = 0
    TROVA_PREZZI_ORD_CAT = 0
    TROVA_PREZZI_ORD_CAL = 0
    TROVA_PREZZI_NO_IMB = 0
    PAR_NonCalcImpDaAssVeloce = 0
    PAR_CalcImpAConfOrd = 0
    PAR_NonVisMsgImpZeroConfOrd = 0
    PAR_AssNewPedDaAssSingola = 0
    PAR_NonCalcPrezzoDueRefArtOrd = 0
    ATTIVA_COMMISSIONI_DA_ORDINE = 0
    RIC_COMM_TIPO_PED_DA_ORD = 0
    VIS_ELECO_RIGHE_ORD = 0
Else
    PREZZI_ARTICOLI_DA_ORDINE = fnNotNullN(rs!PrezziArticoloDaOrdine)
    PREZZI_IMBALLI_DA_ORDINE = fnNotNullN(rs!PrezziImballoDaOrdine)
    PREZZO_INCLUSO_IMBALLO_DA_ORDINE = fnNotNullN(rs!PrezzoInclusoImballoDaOrdine)
    TROVA_PREZZI_ORD_CAT = fnNotNullN(rs!AbilitaCategoriaImportoOrdine)
    TROVA_PREZZI_ORD_CAL = fnNotNullN(rs!AbilitaCalibroImportoOrdine)
    TROVA_PREZZI_NO_IMB = fnNotNullN(rs!NonAbilitareImportoImballo)
    PAR_NonCalcImpDaAssVeloce = fnNotNullN(rs!NonCalcImpDaAssVeloce)
    PAR_CalcImpAConfOrd = fnNotNullN(rs!CalcImpAConfOrd)
    PAR_NonVisMsgImpZeroConfOrd = fnNotNullN(rs!NonVisMsgImpZeroConfOrd)
    PAR_AssNewPedDaAssSingola = fnNotNullN(rs!AssNewPedDaAssSingola)
    PAR_NonCalcPrezzoDueRefArtOrd = fnNotNullN(rs!NonCalcPrezzoDueRefArtOrd)
    ATTIVA_COMMISSIONI_DA_ORDINE = fnNotNullN(rs!AttivaCommissioniDaOrdine)
    RIC_COMM_TIPO_PED_DA_ORD = fnNotNullN(rs!RicalcolaCommTipoPedDaOrdInEvasione)
    VIS_ELECO_RIGHE_ORD = fnNotNullN(rs!VisElencoRigheOrdineSeNonTroviAssociazione)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
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

Set rs = CnDMT.OpenResultset(sSQL)

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

Set rs = CnDMT.OpenResultset(sSQL)

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

Set rs = CnDMT.OpenResultset(sSQL)

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

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE = 0
Else

    GET_LINK_IVA_CLIENTE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function

End Function
Private Sub GET_CONFIGURAZIONE_IMPORTI_ARTICOLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, rstmp As ADODB.Recordset)
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
'            If TROVA_PREZZI_NO_IMB = 0 Then
'                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'            End If
'        End If
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
'                rstmp!ImportoUnitarioArticolo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'
'                rstmp!Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
'                rstmp!Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
'                ImportoUnitario = rstmp!ImportoUnitarioArticolo
'            End If
'
'            rs.CloseResultset
'            Set rs = Nothing
'        End If
'    End If
'End If
'
'If ImportoUnitario > 0 Then Exit Sub
'
'ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
'ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal
'ObjDoc.ReadDataFromPriceList IDListino
'ObjDoc.ReadDataFromDiscountsList

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

rstmp!ImportoUnitarioArticolo = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))

rstmp!Sconto1 = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglioLocal))
rstmp!Sconto2 = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglioLocal))



End Sub
Private Sub GET_CONFIGURAZIONE_IMPORTI_IMBALLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, rstmp As ADODB.Recordset)
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
'
'        If GESTIONE_ORDINE_VIVAIO = 0 Then
'            If TROVA_PREZZI_NO_IMB = 0 Then
'                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
'            End If
'        End If
'
'        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
'
'        If (TROVA_PREZZI_ORD_CAT = 1) Then
'            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
'        End If
'        If (TROVA_PREZZI_ORD_CAL = 1) Then
'            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
'        End If
'
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
'                    rstmp!ImportoUnitarioImballo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'                    ImportoUnitario = fnNotNullN(rstmp!ImportoUnitarioImballo)
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


'ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 2 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
'ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal
'ObjDoc.ReadDataFromPriceList IDListino
'ObjDoc.ReadDataFromDiscountsList

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

rstmp!ImportoUnitarioImballo = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))



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
Private Function GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE(IDOrdine As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Link_Doc_Agente FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE = True
Else
    If fnNotNullN(rs!Link_Doc_agente) = 0 Then
        GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE = True
    Else
        If GET_LINK_REGOLA_PROVV_AGE(fnNotNullN(rs!Link_Doc_agente), TheApp.IDFirm) = 0 Then
            GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE = False
        Else
            GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE = True
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE"
    GET_CONTROLLO_REGOLA_PROV_AGE_ORDINE = True
End Function
Private Function GET_LINK_AGENTE_ORDINE(IDOrdine As Long) As Long
On Error GoTo ERR_GET_LINK_AGENTE_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Link_Doc_Agente FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_AGENTE_ORDINE = 0
Else
    If fnNotNullN(rs!Link_Doc_agente) = 0 Then
        GET_LINK_AGENTE_ORDINE = 0
    Else
        GET_LINK_AGENTE_ORDINE = fnNotNullN(rs!Link_Doc_agente)
    End If
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_LINK_AGENTE_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_LINK_AGENTE_ORDINE"
    GET_LINK_AGENTE_ORDINE = 0
End Function



Private Function GET_LINK_REGOLA_PROVV_AGE(IDAnagraficaAgente As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_REGOLA_PROVV_AGE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRegolaProvv FROM RegolaProvvAgente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaAgente
sSQL = sSQL & " AND IDAzienda=" & IDAzienda
sSQL = sSQL & " AND Predefinita=" & fnNormBoolean(1)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGOLA_PROVV_AGE = 0
Else
    GET_LINK_REGOLA_PROVV_AGE = fnNotNullN(rs!IDRegolaProvv)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_REGOLA_PROVV_AGE:
    MsgBox Err.Description, vbCritical, "GET_LINK_REGOLA_PROVV_AGE"
    GET_LINK_REGOLA_PROVV_AGE = 0
End Function

Private Function GET_PESO_PEDANE_LAVORATE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_PESO_PEDANE_LAVORATE = 0

sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_PESO_PEDANE_LAVORATE = GET_PESO_PEDANE_LAVORATE + GET_TROVA_PESO_PEDANA(fnNotNullN(rs!IDRV_POPedana))
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TROVA_PESO_PEDANA(IDPedana As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POPedana "
sSQL = sSQL & " WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TROVA_PESO_PEDANA = 0
Else
    GET_TROVA_PESO_PEDANA = fnNotNullN(rs!PesoPedana)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PEDANA_PER_RIPESATURA(AvvioVeloce As Boolean)
On Error GoTo ERR_GET_PEDANA_PER_RIPESATURA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Screen.MousePointer = 11

If Not (rsPedana Is Nothing) Then
    If rsPedana.State > 0 Then
        rsPedana.Close
    End If
    Set rsPedana = Nothing
End If

Set rsPedana = New ADODB.Recordset
rsPedana.CursorLocation = adUseClient

rsPedana.Fields.Append "IDPedana", adInteger, , adFldIsNullable
rsPedana.Fields.Append "CodicePedana", adVarChar, 250, adFldIsNullable
rsPedana.Fields.Append "PesoPedana", adDouble, , adFldIsNullable
rsPedana.Fields.Append "PesoMerceLorda", adDouble, , adFldIsNullable
rsPedana.Fields.Append "Registra", adSmallInt, , adFldIsNullable
rsPedana.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsPedana.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsPedana.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsPedana.Fields.Append "TaraTotaleImballi", adDouble, , adFldIsNullable
rsPedana.Fields.Append "NumeroColli", adDouble, , adFldIsNullable

rsPedana.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND Qta_UM>0"

If AvvioVeloce = False Then
    If Me.CDAltroCliente.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDCliente=" & Me.CDAltroCliente.KeyFieldID
    End If
    
    If Me.txtNumeroSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PONumeroOrdinePadre=" & Me.txtNumeroSmistamento.Value
    End If
    
    If Me.txtDataSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PODataOrdinePadre=" & fnNormDate(Me.txtDataSmistamento.Text)
    End If
    
    If Me.txtNListaSmistamento.Value > 0 Then
        sSQL = sSQL & " AND RV_PONumeroListaPrelievo=" & Me.txtNListaSmistamento.Value
    End If
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    
    If Len(Me.txtCodicePedana.Text) > 0 Then
        sSQL = sSQL & " AND CodicePedana LIKE " & fnNormString(Me.txtCodicePedana.Text)
    End If
    
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If
    If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
        sSQL = sSQL & " AND CodiceLottoVendita LIKE " & fnNormString(Me.txtLottoVendita.Text)
    End If
    If Len(Trim(Me.txtRaggrOrdine.Text)) > 0 Then
        sSQL = sSQL & " AND NotaRigaOrdRaggr LIKE " & fnNormString("%" & Me.txtRaggrOrdine.Text & "%")
    End If
Else
    sSQL = sSQL & " AND CodicePedana LIKE " & fnNormString(Mid(Me.txtAssegnazioneVeloce.Text, 2, Len(Me.txtAssegnazioneVeloce.Text)))
End If

sSQL = sSQL & " ORDER BY CodicePedana, CodiceArticolo"

NUMERO_PEDANA_PESATURA = 0
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    rsPedana.Filter = "IDPedana=" & fnNotNullN(rs!IDRV_POPedana)
    'rsPedana.Filter = rsPedana.Filter & " AND IDArticolo=" & fnNotNullN(rs!IDArticolo)
    
    If rsPedana.EOF Then
        rsPedana.AddNew
            rsPedana!IDPedana = fnNotNullN(rs!IDRV_POPedana)
            rsPedana!CodicePedana = fnNotNullN(rs!CodicePedana)
            rsPedana!PesoPedana = fnNotNullN(rs!PesoPedana) 'GET_PESO_PEDANA(rsPedana!IDPedana)
            'rsPedana!PesoMerceLorda = GET_PESO_MERCE_PEDANA(rsPedana!IDPedana)
            rsPedana!Registra = 0
            rsPedana!CodiceArticolo = fnNotNull(rs!CodiceArticolo) 'GET_CODICE_ARTICOLO_PEDANA(rsPedana!IDPedana)
            'rsPedana!TaraTotaleImballi = GET_TARA_MERCE_PEDANA(rsPedana!IDPedana)
            
            NUMERO_PEDANA_PESATURA = NUMERO_PEDANA_PESATURA + 1
        'rsPedana.Update
    End If
    
    rsPedana!PesoMerceLorda = fnNotNullN(rsPedana!PesoMerceLorda) + fnNotNullN(rs!PesoLordo)
    rsPedana!TaraTotaleImballi = fnNotNullN(rsPedana!TaraTotaleImballi) + fnNotNullN(rs!Tara)
    rsPedana!NumeroColli = fnNotNullN(rsPedana!NumeroColli) + fnNotNullN(rs!Colli)
    
    rsPedana.Update
    rsPedana.Filter = vbNullString
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Screen.MousePointer = 0

Exit Function
ERR_GET_PEDANA_PER_RIPESATURA:
MsgBox Err.Description, vbCritical, "GET_PEDANA_PER_RIPESATURA"
Screen.MousePointer = 0
End Function
Private Function GET_PESO_PEDANA(IDPedana As Long) As Double
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
Private Function GET_PESO_MERCE_PEDANA(IDPedana As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(PesoLordo) AS TotalePesoMerce "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_MERCE_PEDANA = 0
Else
    GET_PESO_MERCE_PEDANA = fnNotNullN(rs!TotalePesoMerce)
End If


rs.CloseResultset
Set rs = Nothing


    
End Function
Private Function GET_TARA_MERCE_PEDANA(IDPedana As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Tara) AS TotaleTara "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TARA_MERCE_PEDANA = 0
Else
    GET_TARA_MERCE_PEDANA = fnNotNullN(rs!TotaleTara)
End If


rs.CloseResultset
Set rs = Nothing


    
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

Private Function GET_PEDANA_PER_RIPESATURA_ORDINE(AvvioVeloce As Boolean)
On Error GoTo ERR_GET_PEDANA_PER_RIPESATURA_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Screen.MousePointer = 11

If Not (rsPedana Is Nothing) Then
    If rsPedana.State > 0 Then
        rsPedana.Close
    End If
    Set rsPedana = Nothing
End If

Set rsPedana = New ADODB.Recordset
rsPedana.CursorLocation = adUseClient

rsPedana.Fields.Append "IDPedana", adInteger, , adFldIsNullable
rsPedana.Fields.Append "CodicePedana", adVarChar, 250, adFldIsNullable
rsPedana.Fields.Append "TaraPedana", adDouble, , adFldIsNullable
rsPedana.Fields.Append "PesoPedana", adDouble, , adFldIsNullable
rsPedana.Fields.Append "PesoRealeLordo", adDouble, , adFldIsNullable
rsPedana.Fields.Append "PesoRealeNetto", adDouble, , adFldIsNullable
rsPedana.Fields.Append "TaraTotaleImballi", adDouble, , adFldIsNullable
rsPedana.Fields.Append "NumeroColli", adDouble, , adFldIsNullable
rsPedana.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable

rsPedana.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & Me.txtIDOrdine.Value
sSQL = sSQL & " ORDER BY CodicePedana"


NUMERO_PEDANA_PESATURA = 0

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    
    rsPedana.Filter = "IDPedana=" & fnNotNullN(rs!IDRV_POPedana)
    
    If rsPedana.EOF Then
        rsPedana.AddNew
        rsPedana!IDPedana = fnNotNullN(rs!IDRV_POPedana)
        rsPedana!CodicePedana = fnNotNullN(rs!CodicePedana)
        rsPedana!TaraPedana = GET_PESO_PEDANA(rsPedana!IDPedana)
        rsPedana!PesoRealeLordo = 0
        rsPedana!PesoRealeNetto = 0
        rsPedana!PesoPedana = rsPedana!TaraPedana
        rsPedana!CodiceArticolo = GET_CODICE_ARTICOLO_PEDANA(rs!IDRV_POPedana)
        
        NUMERO_PEDANA_PESATURA = NUMERO_PEDANA_PESATURA + 1
    End If
    
    rsPedana!PesoPedana = rsPedana!PesoPedana + fnNotNullN(rs!PesoLordo)
    rsPedana!TaraTotaleImballi = fnNotNullN(rsPedana!TaraTotaleImballi) + fnNotNullN(rs!Tara)
    rsPedana!NumeroColli = fnNotNullN(rsPedana!NumeroColli) + fnNotNullN(rs!Colli)
    
    rsPedana.Update
    
    rsPedana.Filter = vbNullString
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Screen.MousePointer = 0

Exit Function
ERR_GET_PEDANA_PER_RIPESATURA_ORDINE:
MsgBox Err.Description, vbCritical, "GET_PEDANA_PER_RIPESATURA"
Screen.MousePointer = 0
End Function
Private Function GET_NUMERO_LAVORAZIONI_PER_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_POAssegnazioneMerce) AS NumeroPedane "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & " WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_LAVORAZIONI_PER_PEDANA = 1
Else
    GET_NUMERO_LAVORAZIONI_PER_PEDANA = fnNotNullN(rs!NumeroPedane)
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
On Error Resume Next
ObjDoc.ClearValues

 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica

ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal



End Sub
Private Function GET_CODICE_ARTICOLO_PEDANA(IDPedana As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CODICE_ARTICOLO_PEDANA = ""

sSQL = "SELECT CodiceArticolo "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana
sSQL = sSQL & " GROUP BY CodiceArticolo"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If Len(GET_CODICE_ARTICOLO_PEDANA) > 0 Then
        GET_CODICE_ARTICOLO_PEDANA = GET_CODICE_ARTICOLO_PEDANA & " - "
    End If
    GET_CODICE_ARTICOLO_PEDANA = GET_CODICE_ARTICOLO_PEDANA & fnNotNull(rs!CodiceArticolo)
rs.MoveNext
Wend


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

Set rs = CnDMT.OpenResultset(sSQL)

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
Private Sub ParametroGestioneOrdineVivaio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivaGestioneOrdineVivaio, IDArticoloCommConf, IDRV_POTipoTrattenutaAggiuntivaVivaio "
sSQL = sSQL & "FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GESTIONE_ORDINE_VIVAIO = fnNotNullN(rs!AttivaGestioneOrdineVivaio)
    LINK_ARTICOLO_COMM_CONF = fnNotNullN(rs!IDArticoloCommConf)
    LINK_TIPO_TRATT_AGG_COMM_CONF = fnNotNullN(rs!IDRV_POTipoTrattenutaAggiuntivaVivaio)
Else
    GESTIONE_ORDINE_VIVAIO = 0
    LINK_ARTICOLO_COMM_CONF = 0
    LINK_TIPO_TRATT_AGG_COMM_CONF = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GESTIONE_COMMISSIONE_PER_CONFERIMENTO(IDOggettoOrdine As Long, Aggiungi As Boolean)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "DELETE FROM RV_POCaricoMerceAddebiti "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & fnNotNullN(rs!IDRV_POAssegnazioneMerce)
    CnDMT.Execute sSQL
    
    If Aggiungi = True Then
        AGGIUNGI_COMMISSIONE_PER_LAVORAZIONE IDOggettoOrdine, fnNotNullN(rs!IDArticolo), fnNotNullN(rs!Qta_UM), fnNotNullN(rs!IDRV_POAssegnazioneMerce), GET_LINK_CONFERIMENTO_TESTA(fnNotNullN(rs!IDRV_POCaricoMerceRighe))
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub AGGIUNGI_COMMISSIONE_PER_LAVORAZIONE(IDOggettoOrdine As Long, IDArticoloPadre As Long, QuantitaAssegnazione As Long, IDAssegnazioneMerce As Long, IDTestataConferimento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroCombinazioni As Long
Dim ImportoUnitario As Double
Dim ImportoListinoOrdine As Double

Dim Sconto1 As Double
Dim Sconto2 As Double
Dim IDRegolaProvv As Long
Dim IDAgenteOrdine As Long
Dim PercentualeProvv As Double
Dim ImportoCommissione As Double

IDAgenteOrdine = GET_LINK_AGENTE_ORDINE(IDOggettoOrdine)

If (IDOggettoOrdine = 0) Then Exit Sub

IDRegolaProvv = GET_LINK_REGOLA_PROVV_AGE(IDAgenteOrdine, TheApp.IDFirm)

If IDRegolaProvv = 0 Then Exit Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
'sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
'sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
'sSQL = sSQL & " AND RV_POTipoRiga=1 "
'sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'
'Set rs = CnDMT.OpenResultset(sSQL)
'
'If rs.EOF Then
'    NumeroCombinazioni = 0
'Else
'    NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
'End If
'
'rs.CloseResultset
'Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'If NumeroCombinazioni = 1 Then
    ImportoListinoOrdine = 0
    ImportoUnitario = 0
    Sconto1 = 0
    Sconto2 = 0
    
    'sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
    'sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
    'sSQL = sSQL & " AND RV_POTipoRiga=1 "
    'sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
'    sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
'    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo

    

    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If Not rs.EOF Then
'        ImportoUnitario = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
'        Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
'        Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
        
        ImportoListinoOrdine = fnNotNullN(rs!RV_POImportoUnitarioListino)
        ImportoUnitario = fnNotNullN(rs!ImportoUnitarioArticolo)
        Sconto1 = fnNotNullN(rs!Sconto1)
        Sconto2 = fnNotNullN(rs!Sconto2)
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If ((ImportoListinoOrdine > ImportoUnitario) And (Sconto1 = 0) And (Sconto2 = 0)) Then
        
        Sconto1 = 0
        Sconto2 = 0
        Sconto1 = (1 - (ImportoUnitario / ImportoListinoOrdine)) * 100
        Sconto1 = fnRoundChange(Sconto1, 1, 3)
    Else
        If ((Sconto1 > 0) Or (Sconto2 > 0)) Then
            ImportoUnitario = ImportoUnitario - ((ImportoUnitario / 100) * Sconto1)
            ImportoUnitario = ImportoUnitario - ((ImportoUnitario / 100) * Sconto2)
        End If
        
        
    End If
    
    
    If ImportoUnitario > 0 Then
        
        PercentualeProvv = GET_PERCENTUALE_PROVV(IDRegolaProvv, Sconto1, Sconto2)
        
        If PercentualeProvv > 0 Then
            ImportoCommissione = (ImportoUnitario / 100) * PercentualeProvv
            
            If ((LINK_ARTICOLO_COMM_CONF > 0) And (LINK_TIPO_TRATT_AGG_COMM_CONF > 0)) Then
                
                sSQL = "INSERT INTO RV_POCaricoMerceAddebiti ("
                sSQL = sSQL & "IDRV_POCaricoMerceAddebiti, IDRV_POCaricoMerceTesta, IDArticolo, IDRV_POTipoTrattenutaAggiuntiva, "
                sSQL = sSQL & "Quantita, ImportoUnitario, IDIva, TotaleRigaNettoIva, ImpostaRiga, TotaleRigaLordoIva, "
                sSQL = sSQL & "IDRV_POAssegnazioneMerce"
                sSQL = sSQL & ") "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnGetNewKey("RV_POCaricoMerceAddebiti", "IDRV_POCaricoMerceAddebiti") & ", "
                sSQL = sSQL & IDTestataConferimento & ", "
                sSQL = sSQL & LINK_ARTICOLO_COMM_CONF & ", "
                sSQL = sSQL & LINK_TIPO_TRATT_AGG_COMM_CONF & ", "
                sSQL = sSQL & fnNormNumber(QuantitaAssegnazione) & ", "
                sSQL = sSQL & fnNormNumber(ImportoCommissione) & ", "
                sSQL = sSQL & "0" & ", "
                sSQL = sSQL & fnNormNumber(QuantitaAssegnazione * ImportoCommissione) & ", "
                sSQL = sSQL & "0" & ", "
                sSQL = sSQL & fnNormNumber(QuantitaAssegnazione * ImportoCommissione) & ", "
                sSQL = sSQL & IDAssegnazioneMerce
                sSQL = sSQL & ")"
                
                CnDMT.Execute sSQL
                
            End If
            
        End If
        
    End If

End Sub
Private Function GET_PERCENTUALE_PROVV(IDRegolaProvv As Long, Sconto1 As Double, Sconto2 As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEConfigRegolaProvv "
sSQL = sSQL & "WHERE IDRegolaProvv=" & IDRegolaProvv
sSQL = sSQL & " AND Sconto1=" & fnNormNumber(Sconto1)
sSQL = sSQL & " AND Sconto2=" & fnNormNumber(Sconto2)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PERCENTUALE_PROVV = 0
Else
    GET_PERCENTUALE_PROVV = fnNotNullN(rs!ValoreProvv)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_CONFERIMENTO_TESTA(IDRigaConferimento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceTesta FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CONFERIMENTO_TESTA = 0
Else
    GET_LINK_CONFERIMENTO_TESTA = fnNotNullN(rs!IDRV_POCaricoMerceTesta)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub ELIMINA_COMMISSIONE_LAVORAZIONE(IDAssegnazioneMerce As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POCaricoMerceAddebiti "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce
CnDMT.Execute sSQL
    
End Sub
Private Sub GET_IMPONIBILE_ORDINE(IDOggettoOrdine As Long)
On Error GoTo ERR_GET_IMPONIBILE_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Totale As Double
Dim TotaleUnitario As Double

sSQL = "SELECT IDOggettoOrdine, Qta_UM, Colli, ImportoUnitarioImballo, Sconto2, Sconto1, MerceInclusoImballo, ImportoUnitarioArticolo "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)
Totale = 0

While Not rs.EOF
    TotaleUnitario = (fnNotNullN(rs!Qta_UM) * fnNotNullN(rs!ImportoUnitarioArticolo))
    TotaleUnitario = TotaleUnitario - ((TotaleUnitario / 100)) * fnNotNullN(rs!Sconto1)
    TotaleUnitario = TotaleUnitario - ((TotaleUnitario / 100)) * fnNotNullN(rs!Sconto2)
    If (fnNotNullN(rs!MerceInclusoImballo) = 0) Then
        TotaleUnitario = TotaleUnitario + (fnNotNullN(rs!Colli) * fnNotNullN(rs!ImportoUnitarioImballo))
    End If
    
    Totale = Totale + TotaleUnitario
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Me.txtImponibileOrdPrep.Text = FormatNumber(Totale, 2, , , vbTrue)

Exit Sub
ERR_GET_IMPONIBILE_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_IMPONIBILE_ORDINE"
    Totale = 0
    Me.txtImponibileOrdPrep.Text = FormatNumber(Totale, 2, , , vbTrue)
End Sub
Private Function GET_LINK_ORDINE_PADRE(IDOggettoOrdine As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto, RV_POIDOrdinePadre "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ORDINE_PADRE = 0
Else
    GET_LINK_ORDINE_PADRE = fnNotNullN(rs!RV_POIDOrdinePadre)
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

CnDMT.Execute sSQL

Exit Sub
ERR_AGGIORNA_NUMERAZIONE_ORDINE:
    MsgBox Err.Description, vbCritical, "AGGIORNA_NUMERAZIONE_ORDINE"
End Sub
Private Function GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(IDArticolo As Long, IDImballo As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset) As Boolean
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
                rstmp!ImportoUnitarioArticolo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                rstmp!RV_POImportoUnitarioListino = fnNotNullN(rs!RV_POImportoUnitarioListino)
                rstmp!Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
                rstmp!Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
                rstmp!NotaRigaOrdRaggr = fnNotNull(rs!RV_PONotaRigaOrdRaggr)
                
                If (IDImballo = fnNotNullN(rs!RV_POIDImballo)) Then
                    If fnNotNullN(rs!RV_POLinkRiga) > 0 Then
                        RETURN_SEL_PREZZO_IMB_DA_ORD = 1
                        GET_PREZZO_IMBALLO_DA_ORDINE IDImballo, fnNotNullN(rs!RV_POLinkRiga), IDOggettoOrdine, rstmp
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
            If (PAR_NonCalcPrezzoDueRefArtOrd = 0) Then
                If Screen.MousePointer = 11 Then
                    Screen.MousePointer = 0
                    AttivaCursore = True
                End If
                
                LINK_ARTICOLO_ORDINE = rstmp!IDArticolo
                LINK_ORDINE_PER_PREZZO = IDOggettoOrdine
                Set RECORDSET_RETURN_PER_PREZZO = rstmp
                MODALITA_RECUPERO_RIGA_ORD = 0
    
                frmCorpoOrdine.Show vbModal
                
                If CONFERMA_SEL_PREZZO_DA_ORD = 1 Then
                    Result = True
                End If
                If AttivaCursore = True Then
                    Screen.MousePointer = 11
                    DoEvents
                End If
            Else
                Result = True
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
                Set RECORDSET_RETURN_PER_PREZZO = rstmp
                MODALITA_RECUPERO_RIGA_ORD = 0
                LINK_LAVORAZIONE_PER_PREZZO_ORD = rstmp!IDRV_POAssegnazioneMerce
                
                frmCorpoOrdine2.Show vbModal
                
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

Private Sub GET_PREZZO_IMBALLO_DA_ORDINE(IDImballo As Long, linkRiga As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=2 "
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    rstmp!ImportoUnitarioImballo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
    rstmp!MerceInclusoImballo = Abs(fnNotNullN(rs!RV_POImportoImballoInArticolo))
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
Private Sub PREZZATURA_ORDINE(IDOrdine As Long)
On Error GoTo ERR_PREZZATURA_ORDINE
Dim sSQL As String
Dim rs As ADODB.Recordset

'sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
'sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOrdine
'
'Set rs = New ADODB.Recordset
'rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.GrigliaAssegnazione.UpdatePosition = False
rsGrigliaAss.MoveFirst

While Not rsGrigliaAss.EOF
    Screen.MousePointer = 11
    DoEvents
    If PREZZI_ARTICOLI_DA_ORDINE = 0 Then
        GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rsGrigliaAss!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGrigliaAss!Qta_UM), rsGrigliaAss
        GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rsGrigliaAss!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGrigliaAss!Qta_UM), rsGrigliaAss
        rsGrigliaAss!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDImballoVendita), Me.cdCliente.KeyFieldID)
    Else
        If (GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(fnNotNullN(rsGrigliaAss!IDArticolo), fnNotNullN(rsGrigliaAss!IDImballoVendita), Me.txtIDOrdinePadre.Value, rsGrigliaAss) = False) Then
            GET_CONFIGURAZIONE_IMPORTI_ARTICOLO Me.cdCliente.KeyFieldID, fnNotNullN(rsGrigliaAss!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGrigliaAss!Qta_UM), rsGrigliaAss
            GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rsGrigliaAss!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGrigliaAss!Qta_UM), rsGrigliaAss
            rsGrigliaAss!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rsGrigliaAss!IDImballoVendita), Me.cdCliente.KeyFieldID)
        Else
            If RETURN_SEL_PREZZO_IMB_DA_ORD = 0 Then
                GET_CONFIGURAZIONE_IMPORTI_IMBALLO Me.cdCliente.KeyFieldID, fnNotNullN(rsGrigliaAss!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGrigliaAss!Qta_UM), rsGrigliaAss
                rsGrigliaAss!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rsGrigliaAss!IDImballoVendita), Me.cdCliente.KeyFieldID)
            End If
        End If
    End If
    
    Screen.MousePointer = 0
    DoEvents

rsGrigliaAss.MoveNext
Wend
Me.GrigliaAssegnazione.UpdatePosition = True
'rs.Close
'Set rs = Nothing
Exit Sub
ERR_PREZZATURA_ORDINE:
    MsgBox Err.Description, vbCritical, "PREZZATURA_ORDINE"
    Screen.MousePointer = 0
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
    'Me.txtIDPedana.Value = 0
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
Private Function GET_CONTROLLO_CLIENTE_BLOCCATO(IDAnagrafica As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_CLIENTE_BLOCCATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_CLIENTE_BLOCCATO = False

sSQL = "SELECT BloccoEmissioneDoc "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!BloccoEmissioneDoc) = 1 Then GET_CONTROLLO_CLIENTE_BLOCCATO = True
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_CLIENTE_BLOCCATO:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_CLIENTE_BLOCCATO"

End Function
Private Function GET_CONTROLLO_QTA_LIQ_A_ZERO(IDOggettoOrdine As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Quantita As Double
GET_CONTROLLO_QTA_LIQ_A_ZERO = False


''''CONTROLLO DEGLI IMPORTI UNITARI ARTICOLO A ZERO''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT RV_POAssegnazioneMerce.IDRV_POAssegnazioneMerce, RV_POAssegnazioneMerce.IDArticolo, Articolo.RV_POIDUnitaDiMisuraLiq, "
sSQL = sSQL & "RV_POAssegnazioneMerce.Colli, RV_POAssegnazioneMerce.PesoLordo, RV_POAssegnazioneMerce.PesoNetto, "
sSQL = sSQL & "RV_POAssegnazioneMerce.Tara, RV_POAssegnazioneMerce.Pezzi, RV_POAssegnazioneMerce.Qta_UM "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce INNER JOIN "
sSQL = sSQL & "Articolo ON RV_POAssegnazioneMerce.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & " WHERE IDOggettoOrdine=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If GET_CONTROLLO_QTA_LIQ_A_ZERO = False Then
        Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
            Case 1
                Quantita = fnNotNullN(rs!Colli)
            Case 2
                Quantita = fnNotNullN(rs!PesoLordo)
            Case 3
                Quantita = fnNotNullN(rs!PesoNetto)
            Case 4
                Quantita = fnNotNullN(rs!Tara)
            Case 5
                Quantita = fnNotNullN(rs!Pezzi)
            Case Else
                Quantita = fnNotNullN(rs!Qta_UM)
        End Select
    End If
    If Quantita = 0 Then
        GET_CONTROLLO_QTA_LIQ_A_ZERO = True
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

End Function
