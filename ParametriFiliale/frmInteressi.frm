VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmInteressi 
   Caption         =   "INTERESSI"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInteressi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElimina_Interessi 
      Caption         =   "Elimina"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalva_Interessi 
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuovo_Interessi 
      Caption         =   "Nuovo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin DMTEDITNUMLib.dmtNumber txtNumeroGiorniAnno 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "0"
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
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtTassoInteressi 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "0"
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
      Appearance      =   1
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTDATETIMELib.dmtDate txtDataInizioInteressi 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   503
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
      Appearance      =   1
   End
   Begin DmtGridCtl.DmtGrid GrigliaInteressi 
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
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
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "0"
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
      DecimalPlaces   =   5
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin VB.Label Label13 
      Caption         =   "Data inizio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "% Annuale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "N° giorni anno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "% Giornaliera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmInteressi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
