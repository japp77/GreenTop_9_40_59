VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAnalizzaOrdine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANALIZZA ANDAMENTO ORDINE"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalizzaOrdine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   600
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   720
      TabIndex        =   44
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   0
      ScaleHeight     =   9135
      ScaleWidth      =   12390
      TabIndex        =   0
      Top             =   0
      Width           =   12390
      Begin VB.TextBox txtAndamentoOrdineDett 
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
         Height          =   220
         Index           =   0
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   2670
         Visible         =   0   'False
         Width           =   375
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
         Height          =   220
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   1550
         Width           =   375
      End
      Begin VB.TextBox txtCodiceArticoloPadre 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtArticoloPadre 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2280
         Visible         =   0   'False
         Width           =   4695
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroPedanaPadre 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroColliPadre 
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaArticoloPadre 
         Height          =   285
         Index           =   0
         Left            =   10200
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroPedaneLavorate 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroColliLavorati 
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaArticoloLavorati 
         Height          =   285
         Index           =   0
         Left            =   10200
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDiffPedana 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDiffNumeroColli 
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDiffQtaArticolo 
         Height          =   285
         Index           =   0
         Left            =   10200
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePedaneOrd 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleColliOrd 
         Height          =   285
         Left            =   3360
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePedaneLav 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleColliLav 
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePedaneDiff 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleColliDiff 
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoOrd 
         Height          =   285
         Left            =   5040
         TabIndex        =   18
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoLav 
         Height          =   285
         Left            =   5040
         TabIndex        =   19
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoDiff 
         Height          =   285
         Left            =   5040
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleTaraOrd 
         Height          =   285
         Left            =   6720
         TabIndex        =   21
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleTaraLav 
         Height          =   285
         Left            =   6720
         TabIndex        =   22
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleTaraDiff 
         Height          =   285
         Left            =   6720
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoNettoOrd 
         Height          =   285
         Left            =   8400
         TabIndex        =   24
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoNettoLav 
         Height          =   285
         Left            =   8400
         TabIndex        =   25
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoNettoDiff 
         Height          =   285
         Left            =   8400
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePezziOrd 
         Height          =   285
         Left            =   10080
         TabIndex        =   27
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePezziLav 
         Height          =   285
         Left            =   10080
         TabIndex        =   28
         Top             =   840
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePezziDiff 
         Height          =   285
         Left            =   10080
         TabIndex        =   29
         Top             =   1200
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin MSComctlLib.ProgressBar PBOrdine 
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         ToolTipText     =   "Andamento ordine"
         Top             =   1520
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar PBOrdineDettaglio 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "Andamento ordine del prodotto"
         Top             =   2640
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   240
         X2              =   11880
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   11880
         Y1              =   1850
         Y2              =   1850
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordinato"
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
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Lavorato"
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
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Differenza"
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
         Index           =   7
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "N° pedane"
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
         Index           =   8
         Left            =   1680
         TabIndex        =   40
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "N° colli"
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
         Index           =   9
         Left            =   3360
         TabIndex        =   39
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Peso lordo"
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
         Index           =   10
         Left            =   5040
         TabIndex        =   38
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Tara"
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
         Index           =   11
         Left            =   6720
         TabIndex        =   37
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Peso netto"
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
         Index           =   12
         Left            =   8400
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pezzi"
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
         Index           =   13
         Left            =   10080
         TabIndex        =   35
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblCodiceArticolo 
         Alignment       =   2  'Center
         Caption         =   "Codice articolo"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblArticolo 
         Alignment       =   2  'Center
         Caption         =   "Articolo ordinato"
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
         Left            =   2040
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label lblNumeroPedane 
         Alignment       =   2  'Center
         Caption         =   "N° Pedane"
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
         Left            =   6840
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblNumeroColli 
         Alignment       =   2  'Center
         Caption         =   "N° Colli"
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
         Left            =   8520
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblQuantitaArticolo 
         Alignment       =   2  'Center
         Caption         =   "Q.tà articolo"
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
         Left            =   10200
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAnalizzaOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TOP_INIZIALE As Long
Private Const TOP_ADD As Long = 360
Private Const TOP_ADD_INIZIO As Long = 120

Private IControl As Integer

Private rsTmpPed As ADODB.Recordset

Private Sub Form_Activate()
On Error Resume Next
    Me.SetFocus
    GET_TOTALI_ORDINE LINK_ORDINE_SELEZIONATO
    GET_TOTALE_LAVORAZIONE_ORDINE LINK_ORDINE_SELEZIONATO
    GET_DIFFERENZA_TOTALE
    
    TOP_INIZIALE = 1920
    
    IControl = 0
    
    ELABORAZIONE_DATI_PER_ARTICOLO_ORDINE LINK_ORDINE_SELEZIONATO
    
    If TOP_INIZIALE >= Me.Pic1.Height Then
        Me.Pic1.Height = TOP_INIZIALE + 40
    End If
    Form_Resize
    Screen.MousePointer = 0
    Form_Click
End Sub

Private Sub GET_TOTALI_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(RV_POQuantitaPedanaEffettiva) as NumeroPedane, "
sSQL = sSQL & "SUM(Art_Numero_Colli) as NumeroColli, "
sSQL = sSQL & "SUM(Art_Peso) as PesoLordo, "
sSQL = sSQL & "SUM(Art_Tara) as Tara, "
sSQL = sSQL & "SUM(Art_quantita_pezzi) as Pezzi "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotalePedaneOrd.Value = 0
    Me.txtTotaleColliOrd.Value = 0
    Me.txtTotalePesoLordoOrd.Value = 0
    Me.txtTotaleTaraOrd.Value = 0
    Me.txtTotalePesoLordoOrd.Value = 0
    Me.txtTotalePezziOrd.Value = 0
Else
    Me.txtTotalePedaneOrd.Value = fnNotNullN(rs!NumeroPedane)
    Me.txtTotaleColliOrd.Value = fnNotNullN(rs!NumeroColli)
    Me.txtTotalePesoLordoOrd.Value = fnNotNullN(rs!PesoLordo)
    Me.txtTotaleTaraOrd.Value = fnNotNullN(rs!Tara)
    Me.txtTotalePesoNettoOrd.Value = fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara)
    Me.txtTotalePezziOrd.Value = fnNotNullN(rs!Pezzi)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_TOTALE_LAVORAZIONE_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Colli) as NumeroColli, "
sSQL = sSQL & "SUM(PesoLordo) as PesoLordo, "
sSQL = sSQL & "SUM(Tara) as Tara, "
sSQL = sSQL & "SUM(Pezzi) as Pezzi, "
sSQL = sSQL & "SUM(PesoNetto) as PesoNetto "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotalePedaneLav.Value = 0
    Me.txtTotaleColliLav.Value = 0
    Me.txtTotalePesoLordoLav.Value = 0
    Me.txtTotaleTaraLav.Value = 0
    Me.txtTotalePesoLordoLav.Value = 0
    Me.txtTotalePezziLav.Value = 0
Else
    Me.txtTotalePedaneLav.Value = GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine)
    Me.txtTotaleColliLav.Value = fnNotNullN(rs!NumeroColli)
    Me.txtTotalePesoLordoLav.Value = fnNotNullN(rs!PesoLordo)
    Me.txtTotaleTaraLav.Value = fnNotNullN(rs!Tara)
    Me.txtTotalePesoNettoLav.Value = fnNotNullN(rs!PesoNetto)
    Me.txtTotalePezziLav.Value = fnNotNullN(rs!Pezzi)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine As Long) As Double
'Dim sSQL As String
'Dim rs As DmtOleDbLib.adoResultset
'
'GET_NUMERO_PEDANE_LAVORATE = 0
'
'sSQL = "SELECT IDRV_POPedana "
'sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
'sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
'If chkRicerca.Value = vbChecked Then
'    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
'End If
'sSQL = sSQL & " GROUP BY IDRV_POPedana"
'
'Set rs = Cn.OpenResultset(sSQL)
'
'While Not rs.EOF
'    GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + 1
'rs.MoveNext
'Wend
'
'rs.CloseResultset
'Set rs = Nothing
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroColliPerPedana As Double
Dim NumeroColliSpezzati As Double
GET_NUMERO_PEDANE_LAVORATE = 0

'sSQL = "SELECT IDRV_POPedana "
'sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
'sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
'sSQL = sSQL & " GROUP BY IDRV_POPedana"
'
'Set rs = Cn.OpenResultset(sSQL)
'
'While Not rs.EOF
'    GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE = GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE + 1
'rs.MoveNext
'Wend
'
'rs.CloseResultset
'Set rs = Nothing
CREA_RECORDSET_PEDANA_TMP

sSQL = "SELECT RV_POAssegnazioneMerce.IDRV_POPedana, RV_POAssegnazioneMerce.IDImballoVendita, RV_POAssegnazioneMerce.Colli, RV_POPedana.IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce LEFT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POAssegnazioneMerce.IDRV_POPedana = RV_POPedana.IDRV_POPedana "
sSQL = sSQL & " WHERE RV_POAssegnazioneMerce.IDOggettoOrdinePadre=" & IDOggettoOrdine
'If chkRicerca.Value = vbChecked Then
'    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
'End If
'sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine
'sSQL = sSQL & " GROUP BY RV_POAssegnazioneMerce.IDRV_POPedana"

Set rs = Cn.OpenResultset(sSQL)
    
While Not rs.EOF
    NumeroColliPerPedana = GET_COLLI_PER_TIPO_PEDANA(fnNotNullN(rs!IDRV_POTipoPedana), fnNotNullN(rs!IDImballoVendita))
    If GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA(fnNotNullN(rs!IDRV_POPedana)) = False Then
        If NumeroColliPerPedana = 0 Then
            GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + 1
        Else
            NumeroColliSpezzati = (fnNotNullN(rs!Colli) / NumeroColliPerPedana)
            GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + NumeroColliSpezzati
        End If
    Else
        If NumeroColliPerPedana > 0 Then
            NumeroColliSpezzati = FormatNumber((fnNotNullN(rs!Colli) / NumeroColliPerPedana), 2)
            GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + NumeroColliSpezzati
        End If
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsTmpPed.Close
Set rsTmpPed = Nothing
End Function
Private Sub GET_DIFFERENZA_TOTALE()
Dim NumeroPedaneOrdinate As Double
Dim NumeroPedaneLavorate As Double
Dim Andamento As Double

    Me.txtTotaleColliDiff.Value = Me.txtTotaleColliLav.Value - Me.txtTotaleColliOrd.Value
    Me.txtTotalePedaneDiff.Value = Me.txtTotalePedaneLav.Value - Me.txtTotalePedaneOrd.Value
    Me.txtTotalePesoLordoDiff.Value = Me.txtTotalePesoLordoLav.Value - Me.txtTotalePesoLordoOrd.Value
    Me.txtTotaleTaraDiff.Value = Me.txtTotaleTaraLav.Value - Me.txtTotaleTaraOrd.Value
    Me.txtTotalePesoNettoDiff.Value = Me.txtTotalePesoNettoLav.Value - Me.txtTotalePesoNettoOrd.Value
    Me.txtTotalePezziDiff.Value = Me.txtTotalePezziLav.Value - Me.txtTotalePezziOrd.Value
        
'''''''''''''''''''''''''''CALCOLO DELLE PERCENTUALE DI ANDAMENTO''''''''''''''''''''''''''''''''''''''''''''''''

    Me.PBOrdine.Visible = False
    Me.txtAndamentoOrdine.Visible = False
    Me.PBOrdine.Value = 0
    Me.txtAndamentoOrdine.Text = "0%"
    
    NumeroPedaneOrdinate = Me.txtTotalePedaneOrd.Value
    If NumeroPedaneOrdinate = 0 Then Exit Sub
    
    NumeroPedaneLavorate = Me.txtTotalePedaneLav.Value

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
Private Sub ELABORAZIONE_DATI_PER_ARTICOLO_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDValoriOggettoDettaglio, Link_Art_articolo, Art_codice, Art_descrizione, "
sSQL = sSQL & "IDValoriOggettoDettaglio, IDOggetto, RV_POQuantitaPedana,  "
sSQL = sSQL & "Art_numero_colli, Art_peso, Art_tara, Art_quantita_pezzi,  "
sSQL = sSQL & "Link_art_unita_di_misura , Art_quantita_totale, RV_POQuantitaPedanaEffettiva "
sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine

If LINK_ART_ORD_PADRE_SEL > 0 Then
    sSQL = sSQL & " AND Link_Art_articolo=" & LINK_ART_ORD_PADRE_SEL
End If
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    
    IControl = IControl + 1
    
    CREA_RECORDSET_PEDANA_TMP
    ''''''INTESTAZIONI
    Load Me.lblCodiceArticolo(IControl)
    With Me.lblCodiceArticolo(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
    End With
    
    Load Me.lblArticolo(IControl)
    With Me.lblArticolo(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
    End With
    
    Load Me.lblNumeroPedane(IControl)
    With Me.lblNumeroPedane(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
    End With
    
    Load Me.lblNumeroColli(IControl)
    With Me.lblNumeroColli(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
    End With

    Load Me.lblQuantitaArticolo(IControl)
    With Me.lblQuantitaArticolo(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    TOP_INIZIALE = TOP_INIZIALE + TOP_ADD
    
    ''''''''''DATI DELL'ARTICOLO PADRE'''''''''''''''''''''''''''''''''''''''''
    Load Me.txtCodiceArticoloPadre(IControl)
    With Me.txtCodiceArticoloPadre(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Text = fnNotNull(rs!Art_codice)
    End With

    Load Me.txtArticoloPadre(IControl)
    With Me.txtArticoloPadre(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Text = fnNotNull(rs!Art_descrizione)
    End With

    Load Me.txtNumeroPedanaPadre(IControl)
    With Me.txtNumeroPedanaPadre(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = fnNotNullN(rs!RV_POQuantitaPedanaEffettiva)
    End With

    Load Me.txtNumeroColliPadre(IControl)
    With Me.txtNumeroColliPadre(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = fnNotNullN(rs!Art_numero_colli)
    End With

    Load Me.txtQtaArticoloPadre(IControl)
    With Me.txtQtaArticoloPadre(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = fnNotNullN(rs!Art_quantita_totale)
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    TOP_INIZIALE = TOP_INIZIALE + TOP_ADD
    
    ''''''''''''DATI TOTALI DELLA LAVORAZIONE'''''''''''''''''''''''''''''''''''''''''
    Load Me.PBOrdineDettaglio(IControl)
    With Me.PBOrdineDettaglio(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = 0
    End With
    
    Load Me.txtAndamentoOrdineDett(IControl)
    With Me.txtAndamentoOrdineDett(IControl)
        .Top = TOP_INIZIALE + 30
        .Visible = True
        .ZOrder 0
        .Text = "0%"
    End With
    
    Load Me.txtNumeroPedaneLavorate(IControl)
    With Me.txtNumeroPedaneLavorate(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO(IDOggettoOrdine, fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!IDValoriOggettoDettaglio))
    End With
    
    Load Me.txtNumeroColliLavorati(IControl)
    With Me.txtNumeroColliLavorati(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO(IDOggettoOrdine, fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!IDValoriOggettoDettaglio))
    End With
    
    Load Me.txtQtaArticoloLavorati(IControl)
    With Me.txtQtaArticoloLavorati(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO(IDOggettoOrdine, fnNotNullN(rs!Link_Art_articolo), fnGetUMCoop(fnNotNullN(rs!Link_Art_unita_di_misura)), fnNotNullN(rs!IDValoriOggettoDettaglio))
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    TOP_INIZIALE = TOP_INIZIALE + TOP_ADD
    
    ''''''''''''DIFFERENZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Load Me.txtDiffPedana(IControl)
    With Me.txtDiffPedana(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = Me.txtNumeroPedaneLavorate(IControl).Value - Me.txtNumeroPedanaPadre(IControl).Value
    End With
    
    
    Load Me.txtDiffNumeroColli(IControl)
    With Me.txtDiffNumeroColli(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = Me.txtNumeroColliLavorati(IControl).Value - Me.txtNumeroColliPadre(IControl).Value
    End With
    
    Load Me.txtDiffQtaArticolo(IControl)
    With Me.txtDiffQtaArticolo(IControl)
        .Top = TOP_INIZIALE
        .Visible = True
        .ZOrder 0
        .Value = Me.txtQtaArticoloLavorati(IControl).Value - Me.txtQtaArticoloPadre(IControl).Value
    End With
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    TOP_INIZIALE = TOP_INIZIALE + TOP_ADD
    
    '''''''CHIUDI LINEA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Load Me.Line1(IControl)
    With Me.Line1(IControl)
        
        .Visible = True
        .Y1 = TOP_INIZIALE
        .Y2 = TOP_INIZIALE
        
        
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TOP_INIZIALE = TOP_INIZIALE + TOP_ADD_INIZIO
    
    GET_ANDAMENTO_RIGA_ORDINE IControl

rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing



End Sub
Private Function GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO(IDOggettoOrdine As Long, IDArticoloPadre As Long, IDRigaOrdine As Long) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim rstmp As ADODB.Recordset
Dim NumeroColliPerPedana As Long
Dim IDTipoPedana As Double
Dim NumeroColliSpezzati As Double

GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = 0

'sSQL = "SELECT IDRV_POPedana, ImballoVendita, Colli "
'sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
'sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
'sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine
'sSQL = sSQL & " GROUP BY IDRV_POPedana"

sSQL = "SELECT RV_POAssegnazioneMerce.IDRV_POPedana, RV_POAssegnazioneMerce.IDImballoVendita, RV_POAssegnazioneMerce.Colli, RV_POPedana.IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce LEFT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POAssegnazioneMerce.IDRV_POPedana = RV_POPedana.IDRV_POPedana "
sSQL = sSQL & " WHERE RV_POAssegnazioneMerce.IDOggettoOrdinePadre=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine
'sSQL = sSQL & " GROUP BY RV_POAssegnazioneMerce.IDRV_POPedana"

Set rs = Cn.OpenResultset(sSQL)
    
While Not rs.EOF
    NumeroColliPerPedana = GET_COLLI_PER_TIPO_PEDANA(fnNotNullN(rs!IDRV_POTipoPedana), fnNotNullN(rs!IDImballoVendita))
    If GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA(fnNotNullN(rs!IDRV_POPedana)) = False Then
        If NumeroColliPerPedana = 0 Then
            GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO + 1
        Else
            NumeroColliSpezzati = (fnNotNullN(rs!Colli) / NumeroColliPerPedana)
            GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO + NumeroColliSpezzati
        End If
    Else
        If NumeroColliPerPedana > 0 Then
            NumeroColliSpezzati = FormatNumber((fnNotNullN(rs!Colli) / NumeroColliPerPedana), 2)
            GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO + NumeroColliSpezzati
        End If
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
    
End Function


Private Function GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO(IDOggettoOrdine As Long, IDArticoloPadre As Long, IDRigaOrdine As Long) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO = 0


'sSQL = "SELECT * FROM RV_POArticoloFiglioOrdine "
'sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloPadre

'Set rsArt = Cn.OpenResultset(sSQL)

'While Not rsArt.EOF

    sSQL = "SELECT SUM(Colli) as NumeroColli "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
    'sSQL = sSQL & " AND IDArticolo=" & IDArticoloPadre 'fnNotNullN(rsArt!IDArticoloFiglio)
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine
    Set rs = Cn.OpenResultset(sSQL)
        
    If Not rs.EOF Then
        GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO = GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!NumeroColli)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
'rsArt.MoveNext
'Wend

'rsArt.CloseResultset
'Set rsArt = Nothing
End Function
Private Function GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO(IDOggettoOrdine As Long, IDArticoloPadre As Long, IDUMCoop As Long, IDRigaOrdine As Long) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset

GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = 0


'sSQL = "SELECT * FROM RV_POArticoloFiglioOrdine "
'sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloPadre

'Set rsArt = Cn.OpenResultset(sSQL)

'While Not rsArt.EOF

    sSQL = "SELECT SUM(Colli) as NumeroColli, "
    sSQL = sSQL & "SUM(PesoLordo) as PesoLordo, "
    sSQL = sSQL & "SUM(PesoNetto) as PesoNetto, "
    sSQL = sSQL & "SUM(Tara) as Tara, "
    sSQL = sSQL & "SUM(Pezzi) as Pezzi "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
    'sSQL = sSQL & " AND IDArticolo=" & IDArticoloPadre 'fnNotNullN(rsArt!IDArticoloFiglio)
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine
    Set rs = Cn.OpenResultset(sSQL)
        
    If Not rs.EOF Then
        Select Case IDUMCoop
            Case 1
                GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!NumeroColli)
            Case 2
                GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!PesoLordo)
            Case 3
                GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!PesoNetto)
            Case 4
                GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!Tara)
            Case 5
                GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO = GET_QUANTITA_LAVORATA_PER_ARTICOLO_ORDINATO + fnNotNullN(rs!Pezzi)
        End Select
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
'rsArt.MoveNext
'Wend

'rsArt.CloseResultset
'Set rsArt = Nothing
End Function

Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
    Else
        fnGetUMCoop = 0
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


Private Sub Form_Click()
    Screen.MousePointer = 0
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
    
    
End Sub

Private Sub Form_Resize()
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
            If (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height) <= 32767 Then
                .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            Else
                .Max = 32727
            End If
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
    End If

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
Private Sub GET_ANDAMENTO_RIGA_ORDINE(IControl As Integer)
Dim NumeroPedaneOrdinate As Double
Dim NumeroPedaneLavorate As Double
Dim Andamento As Long

Me.txtAndamentoOrdineDett(IControl).BackColor = Me.BackColor

NumeroPedaneOrdinate = Me.txtNumeroPedanaPadre(IControl).Value

If NumeroPedaneOrdinate = 0 Then Exit Sub

Me.PBOrdineDettaglio(IControl).Max = NumeroPedaneOrdinate

NumeroPedaneLavorate = Me.txtNumeroPedaneLavorate(IControl).Value

If NumeroPedaneOrdinate < NumeroPedaneLavorate Then
    Me.PBOrdineDettaglio(IControl).Value = NumeroPedaneOrdinate
Else
    Me.PBOrdineDettaglio(IControl).Value = NumeroPedaneLavorate
End If

Andamento = FormatNumber(((NumeroPedaneLavorate / NumeroPedaneOrdinate) * 100), 0)

Me.txtAndamentoOrdineDett(IControl).Text = Andamento & "%"
If Andamento <= 50 Then
    Me.txtAndamentoOrdineDett(IControl).BackColor = Me.BackColor
Else
    Me.txtAndamentoOrdineDett(IControl).BackColor = &H8000000D
End If
End Sub

Private Function GET_COLLI_PER_TIPO_PEDANA(IDTipoPedana As Long, IDArticoloImballo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LOCAL_QUANTITA As Double
Dim LOCAL_QUANTITA_LAVORATA As Double
Dim Testo As String

GET_COLLI_PER_TIPO_PEDANA = 0

sSQL = "SELECT Quantita FROM RV_POTipoPedanaImballo "
sSQL = sSQL & "WHERE IDRV_POTipoPedana=" & IDTipoPedana
sSQL = sSQL & " AND IDArticolo=" & IDArticoloImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_COLLI_PER_TIPO_PEDANA = 0
Else
    GET_COLLI_PER_TIPO_PEDANA = fnNotNullN(rs!Quantita)
End If

rs.CloseResultset
Set rs = Nothing


End Function

