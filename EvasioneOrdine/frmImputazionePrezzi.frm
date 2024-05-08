VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmImputazionePrezzi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPUTAZIONE PREZZI ORDINE"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   19515
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImputazionePrezzi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   19515
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   0
      ScaleHeight     =   10275
      ScaleWidth      =   19395
      TabIndex        =   8
      Top             =   0
      Width           =   19455
      Begin VB.CommandButton Command2 
         Caption         =   "AGGIORNA DA ORDINE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13440
         TabIndex        =   38
         Top             =   3200
         Width           =   5895
      End
      Begin VB.CommandButton cmdAggiorna 
         Caption         =   "AGGIORNA DA IMPUTAZIONE PREZZI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   37
         Top             =   3200
         Width           =   5895
      End
      Begin VB.Frame fraPedaneOrdine 
         Caption         =   "PEDANE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10215
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   5775
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
            Height          =   315
            Left            =   4800
            Picture         =   "frmImputazionePrezzi.frx":4781A
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Espandi maschera"
            Top             =   0
            Visible         =   0   'False
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
            Left            =   5280
            Picture         =   "frmImputazionePrezzi.frx":47DA4
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Espandi maschera"
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.ListView LVPedana 
            Height          =   9735
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   17171
            View            =   2
            Arrange         =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame FraDettaglioArticoli 
         Caption         =   "ORDINE"
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
         Height          =   6495
         Left            =   5880
         TabIndex        =   10
         Top             =   3720
         Width           =   13455
         Begin DMTEDITNUMLib.dmtNumber txtTotaleDocumento 
            Height          =   255
            Left            =   11040
            TabIndex        =   24
            Top             =   6120
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   450
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   65535
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
         Begin DMTEDITNUMLib.dmtNumber txtTotaleImballo 
            Height          =   255
            Left            =   8520
            TabIndex        =   22
            Top             =   6120
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   450
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   65535
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
         Begin DMTEDITNUMLib.dmtNumber txtTotaleMerce 
            Height          =   255
            Left            =   6000
            TabIndex        =   20
            Top             =   6120
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   450
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   65535
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
         Begin DmtGridCtl.DmtGrid GrigliaCorpo 
            Height          =   5535
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   9763
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
         Begin VB.Label Label4 
            Caption         =   "TOTALE "
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
            Left            =   11040
            TabIndex        =   23
            Top             =   5880
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "TOTALE IMBALLO"
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
            Left            =   8520
            TabIndex        =   21
            Top             =   5880
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "TOTALE MERCE"
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
            Left            =   6000
            TabIndex        =   19
            Top             =   5880
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "IMPUTAZIONE PREZZI"
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
         Left            =   5880
         TabIndex        =   11
         Top             =   960
         Width           =   13455
         Begin VB.CheckBox chkMerceInclusoImballo 
            Caption         =   "Imposta a tutte le righe l'indicazione di merce incluso imballo"
            Height          =   255
            Left            =   3600
            TabIndex        =   35
            Top             =   1560
            Width           =   9135
         End
         Begin VB.CheckBox chkAggiornaPrezzoImballoDaListino 
            Caption         =   "Aggiorna l'importo unitario imballo da listino quando è zero"
            Height          =   255
            Left            =   3600
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1080
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaPrezziAZero 
            Caption         =   "Aggiorna Importo unitario della merce quando è zero"
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   1320
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaDaListino 
            Caption         =   "Aggiorna l'importo unitario della merce come da listino selezionato"
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   1800
            Width           =   9015
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
            Height          =   375
            Left            =   3600
            TabIndex        =   3
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioImballo 
            Height          =   375
            Left            =   5880
            TabIndex        =   4
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
         Begin DMTDataCmb.DMTCombo cboListino 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
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
         Begin DMTEDITNUMLib.dmtNumber txtSconto1 
            Height          =   375
            Left            =   8160
            TabIndex        =   5
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
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
         Begin DMTEDITNUMLib.dmtNumber txtSconto2 
            Height          =   375
            Left            =   10440
            TabIndex        =   6
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
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
         Begin DMTDataCmb.DMTCombo cboListinoMerce 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1080
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
         Begin VB.Label Label3 
            Caption         =   "Listino merce"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 2 %"
            Height          =   255
            Index           =   3
            Left            =   10440
            TabIndex        =   18
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 1 %"
            Height          =   255
            Index           =   2
            Left            =   8160
            TabIndex        =   17
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Listino imballi"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario imballo"
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario articolo"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraRicerca 
         Caption         =   "PARAMETRI"
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
         Height          =   975
         Left            =   5880
         TabIndex        =   9
         Top             =   0
         Width           =   13455
         Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
            Height          =   615
            Left            =   120
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":4832E
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":48386
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":483E6
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
         Begin DmtCodDescCtl.DmtCodDesc CDImballo 
            Height          =   615
            Left            =   4920
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":48440
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":48497
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":484F6
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
            Left            =   9240
            TabIndex        =   25
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":48550
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":4859E
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":485F8
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
      Begin DMTDataCmb.DMTCombo cboPedana 
         Height          =   315
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
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
   End
End
Attribute VB_Name = "frmImputazionePrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'VARIABILI PER LA CONFIGURAZIONE DELL'OGGETTO DOCUMENTO PER PRELEVARE I PREZZI
Private ObjDoc As DmtDocs.cDocument
Private sTabellaTestataLocal As String
Private sTabellaDettaglioLocal As String
Private sTabellaIVALocal As String
Private sTabellaScadenzeLocal As String


Private Link_TipoImballo As Long
Private rsGriglia As ADODB.Recordset

Private Sub cboPedana_Click()
    fnGrigliaAssegnazione
End Sub

Private Sub CDArticolo_ChangeElement()
Dim sSQL As String
Dim IPedana As Long
Dim sSQL_WHERE As String
Dim SQL_Pedana As String
sSQL = ""

    sSQL = sSQL & "(IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDImballoVendita=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If

    If Me.LVPedana.ListItems.Count > 0 Then
        IDPedana = 0
        For I = 1 To Me.LVPedana.ListItems.Count
            If Me.LVPedana.ListItems(I).Checked = True Then
                SQL_Pedana = ""
                
                If IDPedana = 0 Then
                    IDPedana = 1
                Else
                    SQL_Pedana = SQL_Pedana & " OR "
                End If
                                              
                SQL_Pedana = SQL_Pedana & sSQL & " AND IDRV_POPedana=" & fnNotNullN(Me.LVPedana.ListItems(I).Text) & ")"
                sSQL_WHERE = sSQL_WHERE & SQL_Pedana
            End If
        Next
        
    End If
    

    If IDPedana = 0 Then
        rsGriglia.Filter = Mid(sSQL, 2, Len(sSQL))
   
    Else
        rsGriglia.Filter = sSQL_WHERE
    End If
    
    'sSQL
    
    rsGriglia.Sort = "CodicePedana, CodiceArticolo"
    rsGriglia.Requery
    
    If Not (rsGriglia.EOF And rsGriglia.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings
End Sub


Private Sub CDImballo_ChangeElement()
Dim sSQL As String
Dim IPedana As Long
Dim sSQL_WHERE As String
Dim SQL_Pedana As String
sSQL = ""

    sSQL = sSQL & "(IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDImballoVendita=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If

    If Me.LVPedana.ListItems.Count > 0 Then
        IDPedana = 0
        For I = 1 To Me.LVPedana.ListItems.Count
            If Me.LVPedana.ListItems(I).Checked = True Then
                SQL_Pedana = ""
                
                If IDPedana = 0 Then
                    IDPedana = 1
                Else
                    SQL_Pedana = SQL_Pedana & " OR "
                End If
                                              
                SQL_Pedana = SQL_Pedana & sSQL & " AND IDRV_POPedana=" & fnNotNullN(Me.LVPedana.ListItems(I).Text) & ")"
                sSQL_WHERE = sSQL_WHERE & SQL_Pedana
            End If
        Next
        
    End If
    

    If IDPedana = 0 Then
        rsGriglia.Filter = Mid(sSQL, 2, Len(sSQL))
   
    Else
        rsGriglia.Filter = sSQL_WHERE
    End If
    
    'sSQL
    
    rsGriglia.Sort = "CodicePedana, CodiceArticolo"
    rsGriglia.Requery
    If Not (rsGriglia.EOF And rsGriglia.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings
End Sub

Private Sub CDSocio_ChangeElement()
Dim sSQL As String
Dim IPedana As Long
Dim sSQL_WHERE As String
Dim SQL_Pedana As String
sSQL = ""

    sSQL = sSQL & "(IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDImballoVendita=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If

    If Me.LVPedana.ListItems.Count > 0 Then
        IDPedana = 0
        For I = 1 To Me.LVPedana.ListItems.Count
            If Me.LVPedana.ListItems(I).Checked = True Then
                SQL_Pedana = ""
                
                If IDPedana = 0 Then
                    IDPedana = 1
                Else
                    SQL_Pedana = SQL_Pedana & " OR "
                End If
                                              
                SQL_Pedana = SQL_Pedana & sSQL & " AND IDRV_POPedana=" & fnNotNullN(Me.LVPedana.ListItems(I).Text) & ")"
                sSQL_WHERE = sSQL_WHERE & SQL_Pedana
            End If
        Next
        
    End If
    

    If IDPedana = 0 Then
        rsGriglia.Filter = Mid(sSQL, 2, Len(sSQL))
   
    Else
        rsGriglia.Filter = sSQL_WHERE
    End If
    
    'sSQL
    
    rsGriglia.Sort = "CodicePedana, CodiceArticolo"
    rsGriglia.Requery
    
    If Not (rsGriglia.EOF And rsGriglia.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    
    Me.GrigliaCorpo.LoadUserSettings
End Sub

Private Sub cmdAggiorna_Click()
On Error GoTo ERR_cmdAggiorna_Click
Dim PrezzoImballo As Double

If Me.chkAggiornaDaListino.Value = vbChecked Then
    If Me.cboListinoMerce.CurrentID = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "È stato indicato di aggiornare i prezzi da un listino, ma il listino non è stato selezionato" & vbCrLf
        Testo = Testo & "Impossibile continuare"
        MsgBox Testo, vbInformation, "Aggiorna listino"
        
        Exit Sub
        
    End If
End If


    If Not (rsGriglia.BOF And rsGriglia.EOF) Then
        rsGriglia.MoveFirst
        
        Me.GrigliaCorpo.UpdatePosition = False
        
        
        While Not rsGriglia.EOF
            Screen.MousePointer = 11
            DoEvents
            If Me.chkAggiornaDaListino.Value = vbChecked Then
                rsGriglia("ImportoUnitarioArticolo").Value = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsGriglia!IDArticolo), cboListinoMerce.CurrentID)
            Else
            
                If Me.txtImportoUnitarioArticolo.Value > 0 Then
                    If Me.chkAggiornaPrezziAZero.Value = vbChecked Then
                        If fnNotNullN(rsGriglia("ImportoUnitarioArticolo").Value) = 0 Then
                            rsGriglia("ImportoUnitarioArticolo") = Me.txtImportoUnitarioArticolo.Value
                        End If
                    Else
                        rsGriglia("ImportoUnitarioArticolo") = Me.txtImportoUnitarioArticolo.Value
                    End If
                End If
            End If
            If Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked Then
                If Me.cboListino.CurrentID > 0 Then
                    PrezzoImballo = GET_PREZZO_IMBALLO(fnNotNullN(rsGriglia!IDImballoVendita), Me.cboListino.CurrentID)
                    If PrezzoImballo = 0 Then
                        If Me.txtImportoUnitarioImballo.Value > 0 Then
                            If fnNotNullN(rsGriglia("ImportoUnitarioImballo").Value) = 0 Then
                                rsGriglia("ImportoUnitarioImballo").Value = Me.txtImportoUnitarioImballo.Value
                            End If
                        End If
                    Else
                        rsGriglia("ImportoUnitarioImballo").Value = PrezzoImballo
                    End If
                Else
                    If fnNotNullN(rsGriglia("ImportoUnitarioImballo").Value) = 0 Then
                        rsGriglia("ImportoUnitarioImballo") = Me.txtImportoUnitarioImballo.Value
                    End If
                End If
            Else
                If Me.txtImportoUnitarioImballo.Value > 0 Then
                    rsGriglia("ImportoUnitarioImballo") = Me.txtImportoUnitarioImballo.Value
                End If
            End If
            
            If Me.txtSconto1.Value > 0 Then
                rsGriglia("Sconto1") = Me.txtSconto1.Value
            End If
            If Me.txtSconto2.Value > 0 Then
                rsGriglia("Sconto2") = Me.txtSconto2.Value
            End If
            If Me.chkMerceInclusoImballo.Value = 1 Then
                rsGriglia("MerceInclusoImballo") = Me.chkMerceInclusoImballo.Value
            End If
            
            rsGriglia.UpdateBatch
            Screen.MousePointer = 0
            DoEvents
        
        rsGriglia.MoveNext
        Wend
        
        Me.GrigliaCorpo.UpdatePosition = True
        
    End If
    Me.GrigliaCorpo.Refresh
    GET_TOTALE_DOCUMENTO FrmMain.txtIDOrdine.Value

Exit Sub
ERR_cmdAggiorna_Click:
    Screen.MousePointer = 0
    DoEvents
    MsgBox Err.Description, vbCritical, "cmdAggiorna_Click"
End Sub

Private Sub cmdApriFrameOrdine_Click()
    If Me.fraPedaneOrdine.Width = 3015 Then
        Me.fraPedaneOrdine.Width = 8895
        Me.LVPedana.Width = 8655
    Else
        Me.fraPedaneOrdine.Width = 3015
        Me.LVPedana.Width = 2775
    End If
End Sub

Private Sub Command2_Click()
On Error GoTo ERR_cmdAggiorna_Click
Dim PrezzoImballo As Double

If Not (rsGriglia.BOF And rsGriglia.EOF) Then
    rsGriglia.MoveFirst
    
    Me.GrigliaCorpo.UpdatePosition = False
    
    
    While Not rsGriglia.EOF
        Screen.MousePointer = 11
        DoEvents
        
        If GET_CONTROLLO_MERCE_IN_ORDINE(FrmMain.txtIDOrdinePadre.Value) = True Then
            If PREZZI_ARTICOLI_DA_ORDINE = 1 Then
                If (GET_CONFIGURAZIONE_PREZZO_DA_ORDINE(fnNotNullN(rsGriglia!IDArticolo), fnNotNullN(rsGriglia!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rsGriglia)) = False Then
                    GET_CONFIGURAZIONE_IMPORTI_ARTICOLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rsGriglia!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGriglia!Qta_UM), rsGriglia
                    GET_CONFIGURAZIONE_IMPORTI_IMBALLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rsGriglia!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGriglia!Qta_UM), rsGriglia
                    rsGriglia!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rsGriglia!IDImballoVendita), FrmMain.cdCliente.KeyFieldID)
                Else
                    If RETURN_SEL_PREZZO_IMB_DA_ORD = 0 Then
                        GET_CONFIGURAZIONE_IMPORTI_IMBALLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rsGriglia!IDImballoVendita), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rsGriglia!Qta_UM), rsGriglia
                        rsGriglia!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rsGriglia!IDImballoVendita), FrmMain.cdCliente.KeyFieldID)
                    End If
                End If
            End If
        End If

        rsGriglia.UpdateBatch
        Screen.MousePointer = 0
        DoEvents
    rsGriglia.MoveNext
    Wend
    
    Me.GrigliaCorpo.UpdatePosition = True
    
End If

Me.GrigliaCorpo.Refresh

GET_TOTALE_DOCUMENTO FrmMain.txtIDOrdine.Value

Exit Sub
ERR_cmdAggiorna_Click:
    Screen.MousePointer = 0
    DoEvents
    MsgBox Err.Description, vbCritical, "cmdAggiorna_Click"

End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

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

    ParametroImballo
    
    INIT_CONTROLLI
    
    fnGrigliaPedane
    
    fnGrigliaAssegnazione
    
    Me.GrigliaCorpo.LoadUserSettings
    
    Me.cboListino.WriteOn GET_LISTINO_DEFAULT(FrmMain.cdCliente.KeyFieldID)
    
    If Me.cboListino.CurrentID > 0 Then
        Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked
    End If
    
    Me.chkAggiornaPrezziAZero.Value = vbChecked
    Me.chkMerceInclusoImballo.Value = 0 'GET_MERCE_INCLUSO_IMBALLO_CLIENTE(FrmMain.cdCliente.KeyFieldID)
    Me.FraDettaglioArticoli.Caption = "ORDINE NUMERO " & FrmMain.txtNumeroOrdine.Value & " DEL " & FrmMain.txtDataOrdine.Text & " DEL CLIENTE " & UCase(FrmMain.cdCliente.Description)
    
    GET_CONFIGURAZIONE_DOCUMENTO
    
    
End Sub
Private Sub INIT_CONTROLLI()
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


    'Imballo
    With Me.CDImballo
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
    
     With Me.CDSocio
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Socio\Fornitore"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Socio\Fornitore"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
 
 
    'Pedane dell'ordine
    With Me.cboPedana
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POPedana"
        .DisplayField = "CodicePedana"
        .Sql = "SELECT IDRV_POPedana, CodicePedana "
        .Sql = .Sql & "FROM RV_POAssegnazioneMerce "
        .Sql = .Sql & "WHERE IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
        .Sql = .Sql & " GROUP BY IDRV_POPedana, CodicePedana"
        .Fill
    End With

    'Listino
    With Me.cboListino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .Sql = "SELECT IDListino, Listino "
        .Sql = .Sql & "FROM Listino "
        .Sql = .Sql & "WHERE IDAzienda=" & TheApp.IDFirm
        .Sql = .Sql & " AND TipoListino=0"
        .Fill
    End With
    
    'Listino merce
    With Me.cboListinoMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .Sql = "SELECT IDListino, Listino "
        .Sql = .Sql & "FROM Listino "
        .Sql = .Sql & "WHERE IDAzienda=" & TheApp.IDFirm
        .Sql = .Sql & " AND TipoListino=0"
        .Fill
    End With
    
    'Inizializza la ListView contenente la ricerca
    With Me.LVPedana
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "Seleziona", 320
        .ColumnHeaders.Add , , "Pedana", 2500
        .ColumnHeaders.Add , , "Colli", 1000, 1
        .ColumnHeaders.Add , , "Q.tà Mov", 1000, 1
        .ColumnHeaders.Add , , "Peso lordo", 1000, 1
        .ColumnHeaders.Add , , "Peso netto", 1000, 1
        .ColumnHeaders.Add , , "Pezzi", 1000, 1

    End With
    
    

End Sub
Private Sub ParametroImballo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoImballo = rs!IDTipoImballo
Else
    Link_TipoImballo = 0
End If

rs.CloseResultset
Set rs = Nothing
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
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If
End Sub

Private Sub GrigliaCorpo_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
    rsGriglia.UpdateBatch
    GET_TOTALE_DOCUMENTO FrmMain.txtIDOrdine.Value
End Sub
Private Sub LVPedana_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim sSQL As String
Dim IPedana As Long
Dim sSQL_WHERE As String
Dim SQL_Pedana As String
sSQL = ""

    sSQL = sSQL & "(IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDImballoVendita=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If

    If Me.LVPedana.ListItems.Count > 0 Then
        IDPedana = 0
        For I = 1 To Me.LVPedana.ListItems.Count
            If Me.LVPedana.ListItems(I).Checked = True Then
                SQL_Pedana = ""
                
                If IDPedana = 0 Then
                    IDPedana = 1
                Else
                    SQL_Pedana = SQL_Pedana & " OR "
                End If
                                              
                SQL_Pedana = SQL_Pedana & sSQL & " AND IDRV_POPedana=" & fnNotNullN(Me.LVPedana.ListItems(I).Text) & ")"
                sSQL_WHERE = sSQL_WHERE & SQL_Pedana
            End If
        Next
        
    End If
    

    If IDPedana = 0 Then
        rsGriglia.Filter = Mid(sSQL, 2, Len(sSQL))
   
    Else
        rsGriglia.Filter = sSQL_WHERE
    End If
    
    'sSQL
    
    rsGriglia.Sort = "CodicePedana, CodiceArticolo"
    rsGriglia.Requery
    If Not (rsGriglia.EOF And rsGriglia.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings
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
Private Sub fnGrigliaAssegnazione()
On Error GoTo ERR_fnGrigliaAssegnazione
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    Dim I As Integer
    Dim Start_Pedana As Integer
    Dim SQL_Pedana As String
    
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
    
    If Me.LVPedana.ListItems.Count > 0 Then
        SQL_Pedana = ""
        Start_Pedana = 0
        For I = 1 To Me.LVPedana.ListItems.Count
            If Me.LVPedana.ListItems(I).Checked = True Then
                
                If Start_Pedana = 1 Then
                    sSQL = sSQL & " OR "
                End If
                
                If SQL_Pedana = "" Then
                    sSQL = sSQL & " AND ("
                    Start_Pedana = 1
                    SQL_Pedana = sSQL
                End If
                
                sSQL = sSQL & "(IDRV_POPedana=" & fnNotNullN(Me.LVPedana.ListItems(I).Text) & ")"
                
            End If
        Next
        
        If Start_Pedana = 1 Then
            sSQL = sSQL & ")"
        End If
        
    End If
    'If Me.cboPedana.CurrentID > 0 Then
    '    sSQL = sSQL & " AND IDRV_POPedana=" & Me.cboPedana.CurrentID
    'End If
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDImballoVendita=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
    End If
    
    
    sSQL = sSQL & " ORDER BY CodicePedana, CodiceArticolo"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
    If Not (rsGriglia Is Nothing) Then
        'rsGriglia.Close
        Set rsGriglia = Nothing
    End If
        
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
    
    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodicePedana", "Pedana", dgchar, True, 1100, dgAlignRight
                .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Articolo", "Articolo", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "Qta_UM", "Quantità", dgDouble, True, 1100, dgAlignRight
                Set cl = .ColumnsHeader.Add("ImportoUnitarioArticolo", "Importo Art.", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("Sconto1", "% Sc. 1", dgDouble, True, 1000, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("Sconto2", "% Sc. 2", dgDouble, True, 1000, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "Colli", "Colli", dgDouble, True, 1100, dgAlignRight
                .ColumnsHeader.Add "IDImballoVendita", "IDImballo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imb.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft
                Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Importo Imb.", dgDouble, True, 900, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("MerceInclusoImballo", "Imp. incl. Imb.", dgBoolean, True, 1300, dgAligncenter)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, True, 1200, dgAlignleft
                .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1000, dgAlignleft
                .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, False, 1000, dgAlignleft
                Set cl = .ColumnsHeader.Add("NotaRigaOrdRaggr", "Raggr. ord.", dgchar, False, 2000, dgAlignleft)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                                    
                
        Set .Recordset = rsGriglia
        .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor
    
    
    GET_TOTALE_DOCUMENTO FrmMain.txtIDOrdine.Value
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub
Private Function GET_LISTINO_DEFAULT(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset


GET_LISTINO_DEFAULT = 0

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


If GET_LISTINO_DEFAULT > 0 Then Exit Function

sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
End If

rs.CloseResultset
Set rs = Nothing

If GET_LISTINO_DEFAULT > 0 Then Exit Function

GET_LISTINO_DEFAULT = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA



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
Private Function GET_PREZZO_IMBALLO(IDImballo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDListinoDefault As Long

sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(IDListino=" & IDListino & ") "
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
Private Sub GET_TOTALE_DOCUMENTO(IDOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TOTALE_MERCE As Double
Dim TOTALE_IMBALLO As Double
Dim TOTALE_DOCUMENTO As Double
Dim IMPORTO_UNITARIO_ARTICOLO As Double

TOTALE_MERCE = 0
TOTALE_IMBALLO = 0
TOTALE_DOCUMENTO = 0


sSQL = "SELECT Qta_UM, Colli, ImportoUnitarioArticolo, ImportoUnitarioImballo, "
sSQL = sSQL & "MerceInclusoImballo, Sconto1, Sconto2 "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOrdine

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    IMPORTO_UNITARIO_ARTICOLO = fnNotNullN(rs!ImportoUnitarioArticolo) - ((fnNotNullN(rs!ImportoUnitarioArticolo) / 100)) * fnNotNullN(rs!Sconto1)
    IMPORTO_UNITARIO_ARTICOLO = IMPORTO_UNITARIO_ARTICOLO - ((fnNotNullN(rs!ImportoUnitarioArticolo) / 100)) * fnNotNullN(rs!Sconto2)
    
    TOTALE_MERCE = TOTALE_MERCE + (IMPORTO_UNITARIO_ARTICOLO * fnNotNullN(rs!Qta_UM))
    If fnNotNullN(rs!MerceInclusoImballo) = 0 Then
        TOTALE_IMBALLO = TOTALE_IMBALLO + (fnNotNullN(rs!ImportoUnitarioImballo) * fnNotNullN(rs!Colli))
    End If
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Me.txtTotaleMerce.Value = TOTALE_MERCE
Me.txtTotaleImballo.Value = TOTALE_IMBALLO
Me.txtTotaleDocumento.Value = TOTALE_MERCE + TOTALE_IMBALLO
End Sub
Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("MerceInclusoImballo").Value))
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean)
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
                
        rsGriglia.Fields("MerceInclusoImballo").Value = Abs(CLng(Selected))
        'sbCheckSelected
                
        rsGriglia.UpdateBatch
                
        Me.GrigliaCorpo.Refresh
                
        GET_TOTALE_DOCUMENTO FrmMain.txtIDOrdine.Value

    End If
End Sub
Private Function GET_MERCE_INCLUSO_IMBALLO_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoInclusoImballo "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_MERCE_INCLUSO_IMBALLO_CLIENTE = 0
Else
    GET_MERCE_INCLUSO_IMBALLO_CLIENTE = Abs(fnNotNullN(rs!PrezzoInclusoImballo))
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Sub fnGrigliaPedane()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem
Dim CodicePedanaSel As String




Me.LVPedana.ListItems.Clear

'sSQL = "SELECT IDRV_POPedana, CodicePedana, SUM(Colli) AS ColliTotali, SUM(PesoLordo) AS PesoLordoTotale, SUM(PesoNetto) "
'sSQL = sSQL & "AS PesoNettoTotale, SUM(Tara) AS TaraTotale, SUM(Pezzi) AS PezziTotali, SUM(Qta_UM) AS Qta_UM_Totale, "
'sSQL = sSQL & "IDImballoVendita, CodiceImballoVendita, ImballoVendita "
'sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
'sSQL = sSQL & "WHERE IDOggettoOrdine=" & FrmMain.txtIDOrdine.Value
'sSQL = sSQL & " GROUP BY IDRV_POPedana, CodicePedana, "
'sSQL = sSQL & " IDImballoVendita, CodiceImballoVendita, ImballoVendita"

sSQL = "SELECT RV_POAssegnazioneMerce.IDRV_POPedana, RV_POAssegnazioneMerce.CodicePedana, SUM(RV_POAssegnazioneMerce.Colli) AS ColliTotali, SUM(RV_POAssegnazioneMerce.PesoLordo) "
sSQL = sSQL & "AS PesoLordoTotale, SUM(RV_POAssegnazioneMerce.PesoNetto) AS PesoNettoTotale, SUM(RV_POAssegnazioneMerce.Tara) AS TaraTotale, SUM(RV_POAssegnazioneMerce.Pezzi) AS PezziTotali,"
sSQL = sSQL & "SUM(RV_POAssegnazioneMerce.Qta_UM) AS Qta_UM_Totale, RV_POAssegnazioneMerce.IDImballoVendita, RV_POAssegnazioneMerce.CodiceImballoVendita,"
sSQL = sSQL & "RV_POAssegnazioneMerce.ImballoVendita , RV_POTipoPedana.CodiceTipoPedana, RV_POTipoPedana.TipoPedana, RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POTipoPedana.IDRV_POTipoPedana = RV_POPedana.IDRV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON RV_POPedana.IDRV_POPedana = RV_POAssegnazioneMerce.IDRV_POPedana "
sSQL = sSQL & "Where (dbo.RV_POAssegnazioneMerce.IDOggettoOrdine = " & FrmMain.txtIDOrdine.Value & ") "
sSQL = sSQL & "GROUP BY RV_POAssegnazioneMerce.IDRV_POPedana, RV_POAssegnazioneMerce.CodicePedana, RV_POAssegnazioneMerce.IDImballoVendita, RV_POAssegnazioneMerce.CodiceImballoVendita,"
sSQL = sSQL & "RV_POAssegnazioneMerce.ImballoVendita , RV_POTipoPedana.CodiceTipoPedana, RV_POTipoPedana.TipoPedana, RV_POTipoPedana.IDRV_POTipoPedana"


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    
    CodicePedanaSel = fnNotNullN(rs!CodicePedana)
    If (fnNotNullN(rs!IDRV_POTipoPedana) > 0) Then
        CodicePedanaSel = CodicePedanaSel & " (" & fnNotNull(rs!CodiceTipoPedana) & ")"
    End If

    Set oItem = Me.LVPedana.ListItems.Add
    
    'Popola l 'item della listview
    oItem.Text = fnNotNullN(rs!IDRV_POPedana)
    oItem.SubItems(1) = CodicePedanaSel
    'oItem.SubItems(2) = fnNotNull(rs!CodiceArticolo)
    'oItem.SubItems(3) = fnNotNull(rs!CodiceImballoVendita)
    oItem.SubItems(2) = fnNotNullN(rs!ColliTotali)
    oItem.SubItems(3) = fnNotNullN(rs!Qta_UM_Totale)
    oItem.SubItems(4) = fnNotNullN(rs!PesoLordoTotale)
    oItem.SubItems(5) = fnNotNullN(rs!PesoNettoTotale)
    oItem.SubItems(6) = fnNotNullN(rs!PezziTotali)

rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_IMPORTO_DA_LISTINO(IDArticolo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDListino=" & IDListino

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_DA_LISTINO = 0
Else
    GET_IMPORTO_DA_LISTINO = fnNotNullN(rs!PrezzoNettoIva)
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
            If Screen.MousePointer = 11 Then
                Screen.MousePointer = 0
                AttivaCursore = True
            End If
            
            LINK_ARTICOLO_ORDINE = rstmp!IDArticolo
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

GET_INTESTAZIONE_DOCUMENTO FrmMain.cdCliente.KeyFieldID, LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA

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
