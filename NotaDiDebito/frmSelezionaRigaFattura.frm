VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmSelezionaRigaFattura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ricerca riga fatturata"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16185
   Icon            =   "frmSelezionaRigaFattura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   16185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraParziali 
      Caption         =   "Parziali"
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
      Height          =   2895
      Left            =   3120
      TabIndex        =   36
      Top             =   6240
      Width           =   2895
      Begin DMTEDITNUMLib.dmtNumber txtPezzi 
         Height          =   315
         Left            =   1200
         TabIndex        =   37
         Top             =   1800
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
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTara 
         Height          =   315
         Left            =   1200
         TabIndex        =   38
         Top             =   1080
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
         Height          =   315
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
         Height          =   315
         Left            =   1200
         TabIndex        =   40
         Top             =   720
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtColli 
         Height          =   315
         Left            =   1200
         TabIndex        =   41
         Top             =   360
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
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label13 
         Caption         =   "Colli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Pezzi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Peso lordo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Tara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Peso netto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame FraTotali 
      Caption         =   "Totali"
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
      Height          =   2895
      Left            =   0
      TabIndex        =   25
      Top             =   6240
      Width           =   3015
      Begin DMTEDITNUMLib.dmtNumber txtColliOriginali 
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Top             =   360
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
         DecimalPlaces   =   5
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordoOriginale 
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   720
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoNettoOriginale 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Top             =   1440
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPezziOriginali 
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   1800
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
         DecimalPlaces   =   5
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTaraOriginale 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
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
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label13 
         Caption         =   "Colli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Pezzi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Peso lordo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Tara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Peso netto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valori da accreditare"
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
      Height          =   2775
      Left            =   9960
      TabIndex        =   13
      Top             =   6360
      Width           =   6135
      Begin VB.CheckBox chkRiportaTuttoDoc 
         Caption         =   "Riporta tutto il documento"
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
         TabIndex        =   20
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdElabora 
         Caption         =   "ELABORA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantitaDaAccreditare 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtCurrency txtImportoDaAccreditare 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   " 0"
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
         CurrencyDecimalPlaces=   4
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleQuantita 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   503
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
      Begin DMTDataCmb.DMTCombo cboTipoVariazione 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2355
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
      Begin VB.Label Label2 
         Caption         =   "Tipo variazione"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3720
         X2              =   3720
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità Sel."
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
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Importo unitario"
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Quantità"
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4575
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   8070
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
   Begin VB.Frame Frame1 
      Caption         =   "CRITERI DI RICERCA"
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   16095
      Begin VB.TextBox txtPrefissoDocColl 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboTipoOggetto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3255
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
         Height          =   615
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1085
         PropCodice      =   $"frmSelezionaRigaFattura.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmSelezionaRigaFattura.frx":47872
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmSelezionaRigaFattura.frx":478D2
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
      End
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "RICERCA"
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
         Left            =   6960
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin DMTDATETIMELib.dmtDate txtDataRicerca 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         Appearance      =   1
      End
      Begin VB.TextBox txtNumeroRicerca 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin DMTEDITNUMLib.dmtNumber txtImportoUnitario 
         Height          =   315
         Left            =   11280
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         Caption         =   "Importo unitario"
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
         Left            =   11280
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Prefisso"
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
         Index           =   10
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Data doc."
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
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "N° doc."
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
         Left            =   4320
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo documento (F10)"
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmSelezionaRigaFattura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub cmdElabora_Click()

If Me.chkRiportaTuttoDoc = vbUnchecked Then
'    If (Me.txtQuantitaDaAccreditare.Value > Me.txtTotaleQuantita.Value) Then
'        MsgBox "La quantita da accreditare non può essere maggiore del totale della quantità selezionata", vbCritical, "Errore immissione dati"
'        Me.txtQuantitaDaAccreditare.Value = Me.txtTotaleQuantita.Value
'        Me.txtQuantitaDaAccreditare.SetFocus
'        Exit Sub
'    End If

    Quantita_da_accreditare = Me.txtQuantitaDaAccreditare.Value
    Colli_da_accreditare = Me.txtColli.Text
    PesoLordo_da_accreditare = Me.txtPesoLordo.Value
    Tara_da_accreditare = Me.txtTara.Value
    PesoNetto_da_accreditare = Me.txtPesoNetto.Value
    Pezzi_da_accreditare = Me.txtPezzi.Text
    
    Importo_da_accreditare = Me.txtImportoDaAccreditare.Value
    
    Quantita_Totale_Selezionata = Me.txtTotaleQuantita.Value
    Colli_Totali_Selezionati = Me.txtColliOriginali.Value
    PesoLordo_Totali_Selezionati = Me.txtPesoLordoOriginale.Value
    Tara_Totali_Selezionati = Me.txtTaraOriginale.Value
    PesoNetto_Totali_Selezionati = Me.txtPesoNettoOriginale.Value
    Pezzi_Totali_Selezionati = Me.txtPezziOriginali.Value
    TIPO_VARIAZIONE_DA_WIZARD = Me.cboTipoVariazione.CurrentID
    
    Elaborazione_da_wizard = True
    
    Rif_UltimoTipoOggetto = Me.cboTipoOggetto.ListIndex
    Rif_UltimoNumeroDoc = Me.txtNumeroRicerca.Text
    Rif_UltimoDataDoc = Me.txtDataRicerca.Text
    Rif_UltimoPrefissoDoc = Me.txtPrefissoDocColl.Text
    
    
    
    RiportaTuttoDocumento = False
    
Else

    Rif_UltimoTipoOggetto = Me.cboTipoOggetto.ListIndex
    Rif_UltimoNumeroDoc = Me.txtNumeroRicerca.Text
    Rif_UltimoDataDoc = Me.txtDataRicerca.Text
    Rif_UltimoPrefissoDoc = Me.txtPrefissoDocColl.Text

    Elaborazione_da_wizard = True
    RiportaTuttoDocumento = True
End If

Unload Me
    
End Sub

Private Sub cmdRicerca_Click()
On Error GoTo ERR_cmdRicerca_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NomeTabellaTesta As String
Dim NomeTabellaCorpo As String
Dim OggettoDocumento As String

If Me.cboTipoOggetto.Text = "" Then
    MsgBox "Inserire il tipo di documento", vbInformation, "Inserimento parametri"
Exit Sub
End If

If Len(Me.txtNumeroRicerca.Text) = 0 Then
    MsgBox "Inserire il numero del documento", vbInformation, "Inserimento parametri"
Exit Sub
End If

If Len(Me.txtDataRicerca.Text) = 0 Then
    MsgBox "Inserire la data del documento", vbInformation, "Inserimento parametri"
Exit Sub
End If

Cn.Execute "DELETE FROM RV_POTMPArticoliNotaDiCredito WHERE IDUtente=" & TheApp.IDUser

Select Case Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.ListIndex)
    
    Case 114
        InserimentoDatiFatturaAccompagnatoria
    Case 4
        InserimentoDatiFatturaDifferita
        
End Select

Exit Sub
ERR_cmdRicerca_Click:
    MsgBox Err.Description, vbCritical, "cmdRicerca_Click"
End Sub
Private Sub InserimentoDatiFatturaDifferita()
On Error GoTo ERR_InserimentoDatiFatturaDifferita
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsDDT As DmtOleDbLib.adoResultset
Dim IDOggettoDiff As Long

Dim OggettoDocumento As String
Dim IDOggettoDifferita As Long
Dim LetteraSezionale As String

Dim IDOggetto As Long
Dim Testo As String

Screen.MousePointer = 11
DoEvents

OggettoDocumento = "Rif. Fattura differita n: "


sSQL = "SELECT ValoriOggettoPerTipo0004.IDOggetto, ValoriOggettoPerTipo0004.Doc_prefisso "
'sSQL = sSQL & "FROM ValoriOggettoPerTipo0004 "
sSQL = sSQL & "FROM ValoriOggettoPerTipo0004 INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo0004.IDOggetto = Oggetto.IDOggetto AND "
sSQL = sSQL & "ValoriOggettoPerTipo0004.IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Doc_numero=" & fnNormString(Me.txtNumeroRicerca.Text)
sSQL = sSQL & " AND  Doc_data=" & fnNormDate(Me.txtDataRicerca.Text)
sSQL = sSQL & " AND  Doc_prefisso=" & fnNormString(Me.txtPrefissoDocColl.Text)
sSQL = sSQL & " AND Link_Nom_anagrafica=" & frmMain.cdAnagrafica.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    Rif_UltimoIDOggetto = fnNotNullN(rs!IDOggetto)
    IDOggettoDiff = fnNotNullN(rs!IDOggetto)
    If IDOggettoDiff > 0 Then
    
        LetteraSezionale = Trim(fnNotNull(rs!Doc_prefisso))
        
        OggettoDocumento = "Rif. Fattura differita n: " & LetteraSezionale & "/"
        
        sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_Data, ValoriOggettoPerTipo0002.Doc_Numero "
        sSQL = sSQL & "FROM FlussoOggettiCollegati INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON FlussoOggettiCollegati.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto AND "
        sSQL = sSQL & "FlussoOggettiCollegati.IDTipoOggetto = ValoriOggettoPerTipo0002.IDTipoOggetto INNER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 ON ValoriOggettoPerTipo0002.IDOggetto = ValoriOggettoDettaglio0004.IDOggetto AND "
        sSQL = sSQL & "ValoriOggettoPerTipo0002.IDTipoOggetto = ValoriOggettoDettaglio0004.IDTipoOggetto "
        sSQL = sSQL & "WHERE (FlussoOggettiCollegati.IDTipoOggettoCollegato = 4) "
        sSQL = sSQL & " AND (FlussoOggettiCollegati.IDOggettoCollegato = " & IDOggettoDiff & ") "
        sSQL = sSQL & " AND (ValoriOggettoDettaglio0004.RV_POTipoRiga = 1)"
        
        If Me.CDArticolo.KeyFieldID > 0 Then
             sSQL = sSQL & " AND (Link_Art_articolo=" & Me.CDArticolo.KeyFieldID & ")"
        End If
        
        Set rsDDT = Cn.OpenResultset(sSQL)
        
        While Not rsDDT.EOF
            sSQL = "INSERT INTO RV_POTMPArticoliNotaDiCredito ("
            sSQL = sSQL & "IDUtente, IDArticolo, CodiceArticolo, DescrizioneArticolo, Quantita, Colli, PesoLordo, Tara, PesoNetto, Pezzi, "
            sSQL = sSQL & "ImportoUnitario, CodiceLotto, IDOggetto, IDTipoOggetto, IDValoriOggettoDettaglio, IDOggettoCollegato, IDIvaVendita, "
            sSQL = sSQL & "IDSocio, CodiceSocio, Socio, NomeSocio, DataConferimento,IDAnagraficaFatturazione, IDCollegamentoConferimento, Selezionato,"
            sSQL = sSQL & "DescrizioneDocumento, AliquotaIva, IDUnitaDiMisura, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "NumeroDocumento, DataDocumento, IDCollegamentoAssegnazioneMerce, IDCollegamentoProcessoLavorazione, UnitaDiMisuraVendita, "
            sSQL = sSQL & "RigaRiscontroPeso, IDRV_POProcessoLavorazione, IDRV_POProcessoLavorazioneRighe, IDRV_POLineaProduzione, IDRV_POTipoUtilizzoLinea, "
            sSQL = sSQL & "ID_Art_dettaglio_prog, Sconto1, Sconto2, QuantitaConferita,"
            sSQL = sSQL & "IDArticoloImballo, CodiceImballo, DescrizioneImballo"
            sSQL = sSQL & ") "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & TheApp.IDUser & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!Link_art_articolo) & ", "
            sSQL = sSQL & fnNormString(rsDDT!art_codice) & ", "
            sSQL = sSQL & fnNormString(rsDDT!art_descrizione) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_quantita_totale) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_numero_colli) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_peso) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_tara) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsDDT!Art_peso) - fnNotNullN(rsDDT!Art_tara)) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_quantita_pezzi) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsDDT!Art_prezzo_unitario_neutro)) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_POCodiceLotto) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!IDOggetto) & ", "
            IDOggetto = fnNotNullN(rs!IDOggetto)
            sSQL = sSQL & fnNotNullN(rsDDT!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!IDValoriOggettoDettaglio) & ", "
            sSQL = sSQL & "0" & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Link_Art_Iva) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDSocio) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_POCodiceSocio) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_POSocio) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_PONomeSocio) & ", "
            sSQL = sSQL & fnNormDate(rsDDT!RV_PODataConferimento) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDAnagraficaFatturazione) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDConferimentoRighe) & ", "
            sSQL = sSQL & fnNormBoolean(0) & ", "
            sSQL = sSQL & fnNormString(OggettoDocumento & GetNumeroDocumentoModificato(Me.txtNumeroRicerca.Text) & " del " & Me.txtDataRicerca.Text) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_aliquota_Iva) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Link_Art_unita_di_misura) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDTipoLavorazione) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDTipoCategoria) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDCalibro) & ", "
            sSQL = sSQL & fnNormString(rsDDT!Doc_Numero) & ", "
            sSQL = sSQL & fnNormDate(rsDDT!Doc_data) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDAssegnazioneMerce) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDProcessoIVGamma) & ", "
            sSQL = sSQL & fnNormString(GET_DESCRIZIONE_UM(fnNotNullN(rsDDT!Link_Art_unita_di_misura))) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_PORigaRiscontroPeso) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDProcessoLavorazione) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDProcessoLavorazioneRighe) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDLineaProduzione) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDTipoUtilizzoLinea) & ", "
            sSQL = sSQL & fnNotNullN(rsDDT!ID_Art_dettaglio_prog) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_sco_in_percentuale_1) & ", "
            sSQL = sSQL & fnNormNumber(rsDDT!Art_sco_in_percentuale_2) & ", "
            Select Case GET_UM_COOP(fnNotNullN(rsDDT!RV_POIDConferimentoRighe))
                Case 1
                    sSQL = sSQL & fnNormNumber(rsDDT!Art_numero_colli) & ", "
                Case 2
                    sSQL = sSQL & fnNormNumber(rsDDT!Art_peso) & ", "
                Case 3
                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsDDT!Art_peso) - fnNotNullN(rsDDT!Art_tara)) & ", "
                Case 4
                    sSQL = sSQL & fnNormNumber(rsDDT!Art_tara) & ", "
                Case 5
                    sSQL = sSQL & fnNormNumber(rsDDT!Art_quantita_pezzi) & ", "
                Case Else
                    sSQL = sSQL & fnNormNumber(0) & ", "
            End Select
            sSQL = sSQL & fnNotNullN(rsDDT!RV_POIDImballo) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_POCodiceImballo) & ", "
            sSQL = sSQL & fnNormString(rsDDT!RV_PODescrizioneImballo)
            sSQL = sSQL & ")"
            
            Cn.Execute sSQL
                
        rsDDT.MoveNext
        Wend
    End If
End If

rs.CloseResultset
Set rs = Nothing

fncGriglia

Screen.MousePointer = 0

Me.txtTotaleQuantita.Value = 0
Me.txtQuantitaDaAccreditare.Value = 0
Me.txtImportoDaAccreditare.Value = 0

Me.Griglia.SetFocus

If IDOggetto > 0 Then
    If GET_CONTROLLO_UTILIZZO_PRECEDENTE(sTabellaDettaglio, Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.ListIndex), IDOggetto) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Una o più righe di questo documento sono già state collegate ad una nota di credito o debito "
        Testo = Testo & "Vuoi continuare?"
        
        MsgBox Testo, vbCritical, "Controllo"
        
    End If
End If
Exit Sub
ERR_InserimentoDatiFatturaDifferita:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "InserimentoDatiFatturaDifferita"
End Sub
Private Sub InserimentoDatiFatturaAccompagnatoria()
On Error GoTo ERR_InserimentoDatiFatturaAccompagnatoria
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NomeTabellaTesta As String
Dim NomeTabellaCorpo As String
Dim OggettoDocumento As String
Dim IDOggetto As Long
Dim Testo As String

OggettoDocumento = "Rif. Fattura accompagnatoria n: "
NomeTabellaTesta = "ValoriOggettoPerTipo0072"
NomeTabellaCorpo = "ValoriOggettoDettaglio0001"

IDOggetto = 0

Screen.MousePointer = 11


sSQL = "SELECT " & NomeTabellaCorpo & ".*, " & NomeTabellaTesta & ".Doc_data, " & NomeTabellaTesta & ".doc_numero "
'sSQL = sSQL & "FROM " & NomeTabellaCorpo & " INNER JOIN "
'sSQL = sSQL & NomeTabellaTesta & " ON " & NomeTabellaCorpo & ".IDOggetto = " & NomeTabellaTesta & ".IDOggetto "
sSQL = sSQL & "FROM " & NomeTabellaCorpo & " INNER JOIN "
sSQL = sSQL & NomeTabellaTesta & " ON " & NomeTabellaCorpo & ".IDOggetto =" & NomeTabellaTesta & ".IDOggetto AND "
sSQL = sSQL & NomeTabellaCorpo & ".IDTipoOggetto = " & NomeTabellaTesta & ".IDTipoOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON " & NomeTabellaTesta & ".IDOggetto = Oggetto.IDOggetto AND "
sSQL = sSQL & NomeTabellaTesta & ".IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Doc_numero=" & fnNormString(Me.txtNumeroRicerca.Text)
sSQL = sSQL & " AND  Doc_data=" & fnNormDate(Me.txtDataRicerca.Text)
sSQL = sSQL & " AND  Doc_prefisso=" & fnNormString(Me.txtPrefissoDocColl.Text)
sSQL = sSQL & " AND Link_Nom_anagrafica=" & frmMain.cdAnagrafica.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1"


If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL = sSQL & " AND Link_Art_articolo=" & Me.CDArticolo.KeyFieldID
End If
If Me.txtImportoUnitario.Value > 0 Then
    sSQL = sSQL & " AND Art_pre_uni_net_sco_net_iva=" & fnNormNumber(Me.txtImportoUnitario.Value)
End If
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Rif_UltimoIDOggetto = fnNotNullN(rs!IDOggetto)
End If

While Not rs.EOF
    sSQL = "INSERT INTO RV_POTMPArticoliNotaDiCredito ("
    sSQL = sSQL & "IDUtente, IDArticolo, CodiceArticolo, DescrizioneArticolo, Quantita, Colli, PesoLordo, Tara, PesoNetto, Pezzi, "
    sSQL = sSQL & "ImportoUnitario, CodiceLotto, IDOggetto, IDTipoOggetto, IDValoriOggettoDettaglio, IDOggettoCollegato, IDIvaVendita, "
    sSQL = sSQL & "IDSocio, CodiceSocio, Socio, NomeSocio, DataConferimento, IDAnagraficaFatturazione, IDCollegamentoConferimento, Selezionato,"
    sSQL = sSQL & "DescrizioneDocumento, AliquotaIva, IDUnitaDiMisura, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
    sSQL = sSQL & "NumeroDocumento, DataDocumento, IDCollegamentoAssegnazioneMerce, IDCollegamentoProcessoLavorazione, UnitaDiMisuraVendita, "
    sSQL = sSQL & "RigaRiscontroPeso, IDRV_POProcessoLavorazione, IDRV_POProcessoLavorazioneRighe, IDRV_POLineaProduzione, IDRV_POTipoUtilizzoLinea, "
    sSQL = sSQL & "ID_Art_dettaglio_prog, Sconto1, Sconto2, QuantitaConferita ,"
    sSQL = sSQL & "IDArticoloImballo, CodiceImballo, DescrizioneImballo"
    sSQL = sSQL & ") "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & fnNotNullN(rs!Link_art_articolo) & ", "
    sSQL = sSQL & fnNormString(rs!art_codice) & ", "
    sSQL = sSQL & fnNormString(rs!art_descrizione) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_quantita_totale) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_numero_colli) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_peso) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_tara) & ", "
    sSQL = sSQL & fnNormNumber(fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_quantita_pezzi) & ", "
    sSQL = sSQL & fnNormNumber(fnNotNullN(rs!Art_prezzo_unitario_neutro)) & ", "
    sSQL = sSQL & fnNormString(rs!RV_POCodiceLotto) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDOggetto) & ", "
    IDOggetto = fnNotNullN(rs!IDOggetto)
    sSQL = sSQL & fnNotNullN(rs!IDTipoOggetto) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDValoriOggettoDettaglio) & ", "
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & fnNormNumber(rs!Link_Art_Iva) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDSocio) & ", "
    sSQL = sSQL & fnNormString(rs!RV_POCodiceSocio) & ", "
    sSQL = sSQL & fnNormString(rs!RV_POSocio) & ", "
    sSQL = sSQL & fnNormString(rs!RV_PONomeSocio) & ", "
    sSQL = sSQL & fnNormDate(rs!RV_PODataConferimento) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDAnagraficaFatturazione) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDConferimentoRighe) & ", "
    sSQL = sSQL & fnNormBoolean(0) & ", "
    'sSQL = sSQL & fnNormString(OggettoDocumento & GetNumeroDocumentoModificato(Me.txtNumeroRicerca.Text) & " del " & Me.txtDataRicerca.Text) & ", "
    sSQL = sSQL & fnNormString(rs!RV_PODescrizioneDocumento) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_aliquota_Iva) & ", "
    sSQL = sSQL & fnNotNullN(rs!Link_Art_unita_di_misura) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDTipoLavorazione) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDTipoCategoria) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDCalibro) & ","
    sSQL = sSQL & fnNormString(rs!Doc_Numero) & ", "
    sSQL = sSQL & fnNormDate(rs!Doc_data) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDAssegnazioneMerce) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDProcessoIVGamma) & ", "
    sSQL = sSQL & fnNormString(GET_DESCRIZIONE_UM(fnNotNullN(rs!Link_Art_unita_di_misura))) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_PORigaRiscontroPeso) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDProcessoLavorazione) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDProcessoLavorazioneRighe) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDLineaProduzione) & ", "
    sSQL = sSQL & fnNotNullN(rs!RV_POIDTipoUtilizzoLinea) & ", "
    sSQL = sSQL & fnNotNullN(rs!ID_Art_dettaglio_prog) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_sco_in_percentuale_1) & ", "
    sSQL = sSQL & fnNormNumber(rs!Art_sco_in_percentuale_2) & ", "
    Select Case GET_UM_COOP(fnNotNullN(rs!RV_POIDConferimentoRighe))
        Case 1
            sSQL = sSQL & fnNormNumber(rs!Art_numero_colli) & ", "
        Case 2
            sSQL = sSQL & fnNormNumber(rs!Art_peso) & ", "
        Case 3
            sSQL = sSQL & fnNormNumber(fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) & ", "
        Case 4
            sSQL = sSQL & fnNormNumber(rs!Art_tara) & ", "
        Case 5
            sSQL = sSQL & fnNormNumber(rs!Art_quantita_pezzi) & ", "
        Case Else
            sSQL = sSQL & fnNormNumber(0) & ", "
    End Select
    sSQL = sSQL & fnNotNullN(rs!RV_POIDImballo) & ", "
    sSQL = sSQL & fnNormString(rs!RV_POCodiceImballo) & ", "
    sSQL = sSQL & fnNormString(rs!RV_PODescrizioneImballo)
    sSQL = sSQL & ")"
    
    Cn.Execute sSQL

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

fncGriglia

Screen.MousePointer = 0

Me.txtTotaleQuantita.Value = 0
Me.txtQuantitaDaAccreditare.Value = 0
Me.txtImportoDaAccreditare.Value = 0

Me.Griglia.SetFocus

If IDOggetto > 0 Then
    If GET_CONTROLLO_UTILIZZO_PRECEDENTE(sTabellaDettaglio, Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.ListIndex), IDOggetto) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Una o più righe di questo documento sono già state collegate ad una nota di credito o debito "
        Testo = Testo & "Vuoi continuare?"
        
        MsgBox Testo, vbCritical, "Controllo"
    End If
End If

Exit Sub
ERR_InserimentoDatiFatturaAccompagnatoria:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "InserimentoDatiFatturaAccompagnatoria"
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyF10 Then
        Label1_Click 0
    End If
End Sub

Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If Griglia.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If Griglia.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("Selezionato").Value), 2
            End If
        End If
    End If
    
    If GET_CONTROLLO_DATI_CONGRUENTI = True Then
        rsGriglia.UpdateBatch
        CALCOLA_QUANTITA_TOTALE
        GET_CALCOLO_PARZIALI Me.txtQuantitaDaAccreditare.Value
    Else
        MsgBox "Impossibile elaborare questo articolo", vbCritical, "Selezione articoli"
        sbSelectSelectedRow False, 2
    End If
    Me.Griglia.SetFocus
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Selezionato").Value), 2
        End If
    End If
    If GET_CONTROLLO_DATI_CONGRUENTI = True Then
        rsGriglia.UpdateBatch
        CALCOLA_QUANTITA_TOTALE
        GET_CALCOLO_PARZIALI Me.txtQuantitaDaAccreditare.Value
    Else
        MsgBox "Impossibile elaborare questo articolo", vbCritical, "Selezione articoli"
        sbSelectSelectedRow False, 2
    End If
    
    Me.Griglia.SetFocus
    
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
            rsGriglia.Fields("Selezionato").Value = Abs(CLng(Selected))
            'sbCheckSelected
            Me.Griglia.Refresh
        End If
End Sub
Private Sub Form_Load()
'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

    'Inizializza la combo dei tipi di documento

Me.cboTipoOggetto.Clear

Me.cboTipoOggetto.AddItem "Fattura differita"
Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.NewIndex) = 4

Me.cboTipoOggetto.AddItem "Fattura accompagnatoria"
Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.NewIndex) = 114

Me.cboTipoOggetto.ListIndex = Rif_UltimoTipoOggetto
Me.txtNumeroRicerca.Text = Rif_UltimoNumeroDoc
Me.txtDataRicerca.Text = Rif_UltimoDataDoc
Me.txtPrefissoDocColl.Text = Rif_UltimoPrefissoDoc
    
With Me.CDArticolo
    Set .Application = TheApp
    Set .Database = TheApp.Database
    .HwndContainer = Me.hwnd
    .CodeField = "CodiceArticolo"
    .DescriptionField = "Articolo"
    .KeyField = "IDArticolo"
    .TableName = "Articolo"
    .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm
    .MenuFunctions("EseguiGestione").Enabled = True
    .PropCodice.Caption = "Codice"
    .PropDescrizione.Caption = "Descrizione"
    .CodeCaption4Find = "Codice Articolo"
    .DescriptionCaption4Find = "Descrizione Articolo"
    .IDExecuteFunction = 6 'Articoli
    .CodeIsNumeric = False
End With

With Me.cboTipoVariazione
    Set .Database = TheApp.Database.Connection
    .AddFieldKey "IDRV_POTipoVariazione"
    .DisplayField = "TipoVariazione"
    .SQL = "SELECT * FROM RV_POTipoVariazione"
End With



Quantita_da_accreditare = 0
Importo_da_accreditare = 0
Quantita_Totale_Selezionata = 0


Elaborazione_da_wizard = False


Cn.Execute "DELETE FROM RV_POTMPArticoliNotaDiCredito WHERE IDUtente=" & TheApp.IDUser
fncGriglia


End Sub
Public Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    

    sSQL = "SELECT * FROM RV_POTMPArticoliNotaDiCredito "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            'Set rsEvent = rsGriglia2.Data
    
        
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                                
                    .ColumnsHeader.Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
                    Set cl = .ColumnsHeader.Add("Selezionato", "Selezionato", dgBoolean, True, 1200, dgAlignleft)
                        cl.Editable = True
                    If Me.cboTipoOggetto.ItemData(Me.cboTipoOggetto.ListIndex) = 4 Then
                        .ColumnsHeader.Add "NumeroDocumento", "N° Doc. D.d.t.", dgchar, True, 1500, dgAlignleft
                        .ColumnsHeader.Add "DataDocumento", "Data doc.", dgchar, True, 1500, dgAlignleft
                    
                    End If
                        
                    .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "DescrizioneArticolo", "Articolo", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "IDArticoloImballo", "IDArticoloImballo", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceImballo", "Codice imballo", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "DescrizioneImballo", "Descrizione imballo", dgchar, False, 1500, dgAlignleft
                    .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgInteger, False, 500, dgAlignRight
                    .ColumnsHeader.Add "UnitaDiMisuraVendita", "U.M.", dgchar, True, 1500, dgAlignleft
                   
                    Set cl = .ColumnsHeader.Add("Quantita", "Q.tà", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 5
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    
                    Set cl = .ColumnsHeader.Add("ImportoUnitario", "Importo unitario", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 6
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Sconto1", "% Sc. 1", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Sconto2", "% Sc. 2", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    .ColumnsHeader.Add "Socio", "Socio", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "NomeSocio", "Nome", dgchar, False, 1500, dgAlignleft
                    .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1500, dgAlignleft
                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
End Sub

Private Sub CALCOLA_QUANTITA_TOTALE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Importo As Double

sSQL = "SELECT SUM(Quantita) as Totale, "
sSQL = sSQL & "SUM(Colli) as TotaleColli, "
sSQL = sSQL & "SUM(PesoLordo) as TotalePesoLordo, "
sSQL = sSQL & "SUM(Tara) as TotaleTara, "
sSQL = sSQL & "SUM(PesoNetto) as TotalePesoNetto, "
sSQL = sSQL & "SUM(Pezzi) as TotalePezzi, "
sSQL = sSQL & "ImportoUnitario "
sSQL = sSQL & "FROM RV_POTMPArticoliNotaDiCredito "
sSQL = sSQL & "WHERE Selezionato=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " GROUP BY ImportoUnitario"
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotaleQuantita.Value = 0
    Me.txtImportoDaAccreditare.Value = 0
    Me.txtColliOriginali.Value = 0
    Me.txtPesoLordoOriginale.Value = 0
    Me.txtTaraOriginale.Value = 0
    Me.txtPesoLordoOriginale.Value = 0
    Me.txtPezziOriginali.Value = 0
    
Else
    Me.txtTotaleQuantita.Value = fnNotNullN(rs!Totale)
    Me.txtImportoDaAccreditare.Value = fnNotNullN(rs!ImportoUnitario)
    Me.txtColliOriginali.Value = fnNotNullN(rs!TotaleColli)
    Me.txtPesoLordoOriginale.Value = fnNotNullN(rs!TotalePesoLordo)
    Me.txtTaraOriginale.Value = fnNotNullN(rs!TotaleTara)
    Me.txtPesoNettoOriginale.Value = fnNotNullN(rs!TotalePesoNetto)
    Me.txtPezziOriginali.Value = fnNotNullN(rs!TotalePezzi)
End If

Me.txtQuantitaDaAccreditare.Value = Me.txtTotaleQuantita.Value

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function GET_CONTROLLO_DATI_CONGRUENTI() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTMPArticoliNotaDiCredito "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Selezionato=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_DATI_CONGRUENTI = True
Else
    If ((rs!CodiceArticolo = Me.Griglia.AllColumns("CodiceArticolo").Value) And (rs!ImportoUnitario = Me.Griglia.AllColumns("ImportoUnitario").Value) And (rs!IDUnitaDiMisura = Me.Griglia.AllColumns("IDUnitaDiMisura").Value) And (rs!Sconto1 = Me.Griglia.AllColumns("Sconto1").Value) And (rs!Sconto2 = Me.Griglia.AllColumns("Sconto2").Value)) Then
        GET_CONTROLLO_DATI_CONGRUENTI = True
    Else
        GET_CONTROLLO_DATI_CONGRUENTI = False
    End If
End If
End Function
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

Private Sub Label1_Click(Index As Integer)
    If Index = 0 Then
        frmSelezionaFattura.Show vbModal
    End If

End Sub
Private Function GET_UM_COOP(IDRigaConferimento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_UM_COOP = 0

sSQL = "SELECT IDUnitaDiMisura FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_UM_COOP = fnNotNullN(rs!IDUnitaDiMisura)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_CALCOLO_PARZIALI(Quantita As Double)
If Me.txtTotaleQuantita.Value = 0 Then
    Me.txtColli.Value = 0
    Me.txtPesoLordo.Value = 0
    Me.txtTara.Value = 0
    Me.txtPesoNetto.Value = 0
    Me.txtPezzi.Value = 0
Else
    Me.txtColli.Value = (Quantita / Me.txtTotaleQuantita.Value) * Me.txtColliOriginali.Value
     Me.txtPesoLordo.Value = (Quantita / Me.txtTotaleQuantita.Value) * Me.txtPesoLordoOriginale.Value
     Me.txtTara.Value = (Quantita / Me.txtTotaleQuantita.Value) * Me.txtTaraOriginale.Value
    Me.txtPesoNetto.Value = (Quantita / Me.txtTotaleQuantita.Value) * Me.txtPesoNettoOriginale.Value
    Me.txtPezzi.Value = (Quantita / Me.txtTotaleQuantita.Value) * Me.txtPezziOriginali.Value
End If
End Sub

Private Sub txtQuantitaDaAccreditare_Change()
    GET_CALCOLO_PARZIALI Me.txtQuantitaDaAccreditare.Value
End Sub
Private Function GET_CONTROLLO_UTILIZZO_PRECEDENTE(TabellaRighe As String, IDTipoOggetto As Long, IDOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDValoriOggettoDettaglio FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE RV_POIDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND RV_POIDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_UTILIZZO_PRECEDENTE = False
Else
    GET_CONTROLLO_UTILIZZO_PRECEDENTE = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_UM(IDUM As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT UnitaDiMisura FROM UnitaDiMisura"
sSQL = sSQL & " WHERE IDUnitaDiMisura=" & IDUM


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_UM = ""
Else
    GET_DESCRIZIONE_UM = fnNotNull(rs!UnitaDiMisura)
End If

rs.CloseResultset
Set rs = Nothing
End Function

