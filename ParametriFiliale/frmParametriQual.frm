VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form frmParametriQual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametri per gestione certificati"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParametriQual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "COMPORTAMENTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4695
      Left            =   120
      TabIndex        =   38
      Top             =   1680
      Width           =   7335
      Begin VB.CheckBox Check7 
         Caption         =   "Non riportare riferimento certificato nel DDT"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   4200
         Width           =   7095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Attiva selezione anagrafiche automatizzato"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3960
         Width           =   7095
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Attiva selezione lotto per varietà articolo del contratto"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   7095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Forza vettore da contratto"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Width           =   7095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Riporta verttore da contratto"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3240
         Width           =   7095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Forza destinazione diversa da contratto"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3000
         Width           =   7095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Riporta destinazione diversa da contratto"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2760
         Width           =   7095
      End
      Begin DMTEDITNUMLib.dmtNumber txtColliPred 
         Height          =   330
         Left            =   5280
         TabIndex        =   19
         Top             =   450
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DmtSearchAccount.DmtSearchACS ACSAnaDest 
         Height          =   585
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   1032
         WidthCode       =   700
         WidthDescription=   3200
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
         CaptionDescription=   "Anagrafica di destinazione per soci diretti"
         CaptionCode     =   "Codice"
         IDSearchTypeConto=   6
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboCatAnaSocioDiretto 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   1080
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
      Begin DMTDataCmb.DMTCombo cboCatAnaProdAcq 
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   1680
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
      Begin DMTDataCmb.DMTCombo cboCatAnaSocioCoop 
         Height          =   315
         Left            =   3720
         TabIndex        =   44
         Top             =   1080
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
      Begin DMTDataCmb.DMTCombo cboCatAnaNoProd 
         Height          =   315
         Left            =   3720
         TabIndex        =   46
         Top             =   1680
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
      Begin DmtCodDescCtl.DmtCodDesc CDArticoloScarto 
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1085
         PropCodice      =   $"frmParametriQual.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmParametriQual.frx":47869
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmParametriQual.frx":478D0
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
      Begin DMTEDITNUMLib.dmtNumber txtMesiDataRevoca 
         Height          =   330
         Left            =   5280
         TabIndex        =   55
         Top             =   2280
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Mesi per data revoca"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   56
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria anagrafica non produttore"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   47
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria anagrafica socio cooperativa"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   45
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria anagrafica produttore Acq."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria anagrafica socio diretto"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Numero colli pred."
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   39
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Frame FraTab 
      Caption         =   "PARAMETRI QUALITATIVI"
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
      Height          =   1635
      Index           =   8
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin DMTEDITNUMLib.dmtCurrency txtQual01 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtQual02 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual03 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtCurrency txtQual04 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual15 
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual14 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual13 
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual12 
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual16 
         Height          =   285
         Left            =   5160
         TabIndex        =   16
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual09 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual10 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual11 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual08 
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual07 
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual06 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQual05 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQualPrz16 
         Height          =   285
         Left            =   5880
         TabIndex        =   17
         Top             =   1080
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   37
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   36
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   35
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4440
         TabIndex        =   33
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3720
         TabIndex        =   32
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   3000
         TabIndex        =   31
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   2280
         TabIndex        =   30
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   1560
         TabIndex        =   29
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   900
         TabIndex        =   28
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   5160
         TabIndex        =   26
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   4440
         TabIndex        =   25
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   3720
         TabIndex        =   24
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   46
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15% +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   47
         Left            =   5160
         TabIndex        =   22
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDocument 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15 € +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   6000
         TabIndex        =   21
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmParametriQual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
    QUAL01 = txtQual01.Value
    QUAL02 = txtQual02.Value
    QUAL03 = txtQual03.Value
    QUAL04 = txtQual04.Value
    QUAL05 = txtQual05.Value
    QUAL06 = txtQual06.Value
    QUAL07 = txtQual07.Value
    QUAL08 = txtQual08.Value
    QUAL09 = txtQual09.Value
    QUAL10 = txtQual10.Value
    QUAL11 = txtQual11.Value
    QUAL12 = txtQual12.Value
    QUAL13 = txtQual13.Value
    QUAL14 = txtQual14.Value
    QUAL15 = txtQual15.Value
    QUAL16 = txtQual16.Value
    QUALPRZ16 = txtQualPrz16.Value
    NUMERO_COLLI_PRED_CERT = Me.txtColliPred.Value
    IDAnagraficaDestinazionePerCertificato = Me.ACSAnaDest.IDAnagrafica
    frmMain.cboCategoriaAnagraficaSocio.WriteOn Me.cboCatAnaSocioCoop.CurrentID
    IDCategoriaAnagraficaSocioDiretto = Me.cboCatAnaSocioDiretto.CurrentID
    IDCategoriaAnagraficaProdAcq = Me.cboCatAnaProdAcq.CurrentID
    IDCategoriaAnagraficaNoProd = Me.cboCatAnaNoProd.CurrentID
    IDArticoloScartoPerCertificato = Me.CDArticoloScarto.KeyFieldID
    RiportaDestinazioneDaContrattoCertificato = Me.Check1.Value
    RiportaVettoreDaContrattoCertificato = Me.Check3.Value
    ForzaDestinazioneDaContrattoCertificato = Me.Check2.Value
    ForzaVettoreDaContrattoCertificato = Me.Check4.Value
    AttivaSelezioneSocioCertPerVarieta = Me.Check5.Value
    AttivaSelezioneAnaVeloceInCert = Me.Check6.Value
    NumeroMesiPerDataRevocaCertificato = txtMesiDataRevoca.Value
    NonRiportareRifCerticatoInDDT = Me.Check7.Value
    CONFERMA_PARAMETRI_QUAL = True
    Unload Me
Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Command1_Click"
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load

    INIT_CONTROLLI

    CONFERMA_PARAMETRI_QUAL = False
    
    txtQual01.Value = QUAL01
    txtQual02.Value = QUAL02
    txtQual03.Value = QUAL03
    txtQual04.Value = QUAL04
    txtQual05.Value = QUAL05
    txtQual06.Value = QUAL06
    txtQual07.Value = QUAL07
    txtQual08.Value = QUAL08
    txtQual09.Value = QUAL09
    txtQual10.Value = QUAL10
    txtQual11.Value = QUAL11
    txtQual12.Value = QUAL12
    txtQual13.Value = QUAL13
    txtQual14.Value = QUAL14
    txtQual15.Value = QUAL15
    txtQual16.Value = QUAL16
    txtQualPrz16.Value = QUALPRZ16
    Me.txtColliPred.Value = NUMERO_COLLI_PRED_CERT
    Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, IDAnagraficaDestinazionePerCertificato
    Me.cboCatAnaSocioCoop.WriteOn frmMain.cboCategoriaAnagraficaSocio.CurrentID
    
    Me.cboCatAnaSocioDiretto.WriteOn IDCategoriaAnagraficaSocioDiretto
    Me.cboCatAnaProdAcq.WriteOn IDCategoriaAnagraficaProdAcq
    Me.cboCatAnaNoProd.WriteOn IDCategoriaAnagraficaNoProd
    Me.CDArticoloScarto.Load IDArticoloScartoPerCertificato
    
    Me.Check1.Value = RiportaDestinazioneDaContrattoCertificato
    Me.Check3.Value = RiportaVettoreDaContrattoCertificato
    Me.Check2.Value = ForzaDestinazioneDaContrattoCertificato
    Me.Check4.Value = ForzaVettoreDaContrattoCertificato
    Me.Check5.Value = AttivaSelezioneSocioCertPerVarieta
    Me.Check6.Value = AttivaSelezioneAnaVeloceInCert
    Me.Check7.Value = NonRiportareRifCerticatoInDDT
    txtMesiDataRevoca.Value = NumeroMesiPerDataRevocaCertificato
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub INIT_CONTROLLI()
    Set Me.ACSAnaDest.Connection = TheApp.Database.Connection
    ACSAnaDest.ApplicationName = App.Title
    ACSAnaDest.Client = App.EXEName
    ACSAnaDest.IDFirm = TheApp.IDFirm
    ACSAnaDest.IDUser = TheApp.IDUser
    ACSAnaDest.UserName = TheApp.User
    ACSAnaDest.SearchType = DmtSearchCustomers
    ACSAnaDest.HwndContainer = Me.hwnd
    
    With Me.cboCatAnaNoProd
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica IS NULL"
        .Fill
    End With
    
    With Me.cboCatAnaProdAcq
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica IS NULL"
        .Fill
    End With
    
    With Me.cboCatAnaSocioCoop
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica IS NULL"
        .Fill
    End With
    
    With Me.cboCatAnaSocioDiretto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica IS NULL"
        .Fill
    End With
    
    With Me.CDArticoloScarto
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Descrizione"
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
    
End Sub

