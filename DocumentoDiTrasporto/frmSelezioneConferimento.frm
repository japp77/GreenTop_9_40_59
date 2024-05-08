VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form frmSelezioneConferimento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELEZIONA CONFERIMENTO"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17340
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelezioneConferimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   17340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ulteiori filtri"
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
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   17295
      Begin DMTDataCmb.DMTCombo cboCatMerc 
         Height          =   315
         Left            =   5520
         TabIndex        =   32
         Top             =   450
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.TextBox txtLottoEntrata 
         Height          =   315
         Left            =   13920
         TabIndex        =   29
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txtDescrizioneArticolo 
         Height          =   315
         Left            =   10560
         TabIndex        =   27
         Top             =   450
         Width           =   3255
      End
      Begin VB.TextBox txtCodiceArticolo 
         Height          =   315
         Left            =   7920
         TabIndex        =   25
         Top             =   450
         Width           =   2535
      End
      Begin DmtSearchAccount.DmtSearchACS ACSSocio 
         Height          =   585
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   1032
         WidthDescription=   3000
         WidthSecondDescription=   1200
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SearchType      =   2
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
         CaptionDescription=   "Socio/Fornitore"
         CaptionCode     =   "Codice"
         OnlyAccounts    =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria merceologica"
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   33
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Lotto di entrata"
         Height          =   255
         Index           =   2
         Left            =   13920
         TabIndex        =   30
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Descrizione articolo"
         Height          =   255
         Index           =   1
         Left            =   10560
         TabIndex        =   28
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Codice articolo"
         Height          =   255
         Index           =   0
         Left            =   7920
         TabIndex        =   26
         Top             =   240
         Width           =   2415
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   12726
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
   Begin VB.Frame FraConferimento 
      Caption         =   "Riepilogo conferimento"
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
      Height          =   7215
      Left            =   14640
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtUnitaDiMisura 
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   555
         Width           =   2415
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaQuadrata 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaLavorata 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaConferita 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1155
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaVenduta 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3075
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaDiffLav 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   4395
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaDiffVend 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   5040
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaProcesso 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2475
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin DMTEDITNUMLib.dmtNumber txtNumeroColliConferiti 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   5835
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
         Left            =   120
         TabIndex        =   20
         Top             =   6315
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   8454143
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
      Begin DMTEDITNUMLib.dmtNumber txtNumeroColliVenduti 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   6795
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
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
      Begin VB.Label Label3 
         Caption         =   "Numero colli venduti"
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
         TabIndex        =   23
         Top             =   6600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Numero colli lavorati"
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
         TabIndex        =   21
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Numero colli conferiti"
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
         TabIndex        =   19
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   2520
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità IV GAMMA"
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
         TabIndex        =   17
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   2520
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label2 
         Caption         =   "Differenza Conf./Vend."
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
         TabIndex        =   15
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Differenza Conf./Lav."
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
         Index           =   16
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità venduta"
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
         TabIndex        =   11
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Unità di misura"
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Quantità conferita"
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
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità lavorata"
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
         Index           =   17
         Left            =   120
         TabIndex        =   6
         Top             =   1605
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità Quadrata"
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
         Index           =   18
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmSelezioneConferimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Load_Griglia As Boolean

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader

sSQL = "SELECT * FROM RV_POIEConferimentoDaVend "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)
sSQL = sSQL & " AND PreConferimento=0"

If frmMain.ACSSocio.IDAnagrafica > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & fnNotNullN(frmMain.ACSSocio.IDAnagrafica)
End If

If frmMain.txtDataConferimento.Value > 0 Then
    sSQL = sSQL & " AND DataDocumento=" & fnNormDate(frmMain.txtDataConferimento.Text)
End If
If frmMain.cboTipoDocumentoCoop.CurrentID > 0 Then
    sSQL = sSQL & " AND IDTipoDocumentoCoop=" & frmMain.cboTipoDocumentoCoop.CurrentID
End If

sSQL = sSQL & " ORDER BY DataDocumento DESC,  Anagrafica, CodiceArticolo"

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient
rsGriglia.Open sSQL, Cn.InternalConnection
    
With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDFiliale", "IDFiliale", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Chiuso", "Chiuso", dgBoolean, True, 800, dgAligncenter
        .ColumnsHeader.Add "IDTipoDocumentoCoop", "IDTipoDocumentoCoop", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoDocumentoCoop", "Tipo documento", dgchar, True, 1200, dgAlignleft
        .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceSocio", "Codice", dgchar, True, 1200, dgAlignleft
        .ColumnsHeader.Add "Anagrafica", "Socio\Fornitore", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "Nome", "Nome Socio\Fornitore", dgchar, False, 1000, dgAlignleft
        .ColumnsHeader.Add "DataDocumento", "Data conf.", dgDate, True, 1800, dgAligncenter
        .ColumnsHeader.Add "NumeroDocumento", "N° conf.", dgInteger, True, 1800, dgAlignRight
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisuraCoop", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisuraCoop", "U.M. Coop.", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDUnitaDiMisuraDiamante", "IDUnitaDiMisuraDiamante", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisuraDiamante", "U.M.", dgchar, True, 1500, dgAlignleft
        Set cl = .ColumnsHeader.Add("Qta_UM", "Q.tà Conf.", dgDouble, True, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 3
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoLordo", "Peso Lordo", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 3
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 3
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoNetto", "Peso Netto", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 3
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "IDCategoriaMerceologica", "IDCategoriaMerceologica", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CategoriaMerceologica", "Categoria merceologica", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "CodiceLotto", "Codice di entrata", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "IDImballo", "IDImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceImballo", "Codice imballo", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneImballo", "Imballo", dgchar, False, 2500, dgAlignleft
        Set cl = .ColumnsHeader.Add("TaraUnitaria", "Tara Unitaria", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "IDArticoloPedana", "IDArticoloPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceArticoloPedana", "Codice art. pedana", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "ArticoloPedana", "Pedana", dgchar, False, 2500, dgAlignleft
        Set cl = .ColumnsHeader.Add("QuantitaPedana", "Q.tà Pedana", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("TaraPedana", "Tara Pedana", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub


Private Sub ACSSocio_ChangedElement()
If Load_Griglia = False Then GET_SQL
    
End Sub

Private Sub cboCatMerc_Click()
    GET_SQL
End Sub

Private Sub Form_Activate()
    DISEGNA_FORM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
         GrigliaCorpo_DblClick
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE32)
    SELEZIONE_RIF_CONF = 0
    Load_Griglia = True
    INIT_CONTROLLI
    
    If frmMain.ACSSocio.IDAnagrafica > 0 Then
        Me.ACSSocio.sbLoadCFByIDAnagrafica 7, frmMain.ACSSocio.IDAnagrafica
        Me.ACSSocio.Enabled = False
    End If
    
    GET_GRIGLIA
    Load_Griglia = False
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub INIT_CONTROLLI()
On Error GoTo ERR_INIT_CONTROLLI
    
    Set Me.ACSSocio.Connection = TheApp.Database.Connection
    ACSSocio.ApplicationName = App.Title
    ACSSocio.Client = App.EXEName
    ACSSocio.IDFirm = TheApp.IDFirm
    ACSSocio.IDUser = TheApp.IDUser
    ACSSocio.UserName = TheApp.User
    ACSSocio.SearchType = DmtSearchSuppliers
    ACSSocio.HwndContainer = Me.hwnd
    
    With Me.cboCatMerc
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaMerceologica"
        .DisplayField = "CategoriaMerceologica"
        .SQL = "SELECT IDCategoriaMerceologica, CategoriaMerceologica FROM CategoriaMerceologica"
        .SQL = .SQL & " ORDER BY CategoriaMerceologica"
    End With

Exit Sub
ERR_INIT_CONTROLLI:
    MsgBox Err.Description, vbCritical, "INIT_CONTROLLI"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Not (rsGriglia Is Nothing) Then
        If rsGriglia.State > 0 Then
            rsGriglia.Close
        End If
        Set rsGriglia = Nothing
    End If
End Sub

Private Sub GrigliaCorpo_DblClick()
On Error GoTo ERR_GrigliaCorpo_DblClick
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        With frmMain
            .ACSSocio.IDAnagrafica = 0
            .ACSSocio.Code = ""
            .ACSSocio.Description = ""
            .ACSSocio.SecondDescription = ""
                    
            .txtDataConferimento.Text = fnNotNull(rsGriglia!DataDocumento)
            .cboTipoDocumentoCoop.WriteOn fnNotNullN(rsGriglia!IDTipoDocumentoCoop)
            .ACSSocio.sbLoadCFByIDAnagrafica 7, fnNotNullN(rsGriglia!IDAnagrafica)
            .CDSocioFatt.Load fnNotNullN(rsGriglia!IDAnagraficaFatturazione)
            Link_RigaConferimento = fnNotNullN(rsGriglia!IDRV_POCaricoMerceRighe)
            .txtIDConferimento.Value = Link_RigaConferimento
            .txtLottoCampagna.Text = fnNotNull(rsGriglia!LottoDiConferimento)
'            .txtQtaConferita.Value = fnNotNullN(rsGriglia!Qta_UM)
'            .txtColliConferiti.Value = fnNotNullN(rsGriglia!Colli)
            LINK_ARTICOLO_CONFERITO_SEL = fnNotNullN(rsGriglia!IDArticolo)
            
            SELEZIONE_RIF_CONF = 1
            Unload Me
        End With
    End If

Exit Sub
ERR_GrigliaCorpo_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_DblClick"
End Sub
Private Sub DISEGNA_FORM()
If VISUALIZZA_RIEP_CONF = 1 Then
    Me.GrigliaCorpo.Width = 14535
    Me.FraConferimento.Visible = True
End If
End Sub
Private Function GET_QUADRATURA(IDRigaConferimento As Long) As Double
On Error GoTo ERR_GET_QUADRATURA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

GET_QUADRATURA = 0


sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POLavorazioneL")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF

    Select Case GET_TIPO_PRODOTTO(rs!IDArticolo)
    
        Case Link_TipoCaloPeso
            GET_QUADRATURA = GET_QUADRATURA + fnNotNullN(rs!RV_POQuantitaMovimentata)
        Case Link_TipoScarto
            GET_QUADRATURA = GET_QUADRATURA + fnNotNullN(rs!RV_POQuantitaMovimentata)
        Case Link_TipoAumentoPeso
            GET_QUADRATURA = GET_QUADRATURA + fnNotNullN(-rs!RV_POQuantitaMovimentata)
        Case Else
            GET_QUADRATURA = GET_QUADRATURA + fnNotNullN(rs!RV_POQuantitaMovimentata)
    End Select
    
rs.MoveNext
Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Function

ERR_GET_QUADRATURA:
    GET_QUADRATURA = GET_QUADRATURA
    
End Function
Private Function GET_LAVORAZIONE(IDRigaConferimento As Long) As Double
On Error GoTo ERR_GET_LAVORAZIONE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

GET_LAVORAZIONE = 0
Me.txtNumeroColliLavorati.Value = 0

sSQL = "SELECT SUM(RV_POQuantitaMovimentata) AS QuantitaMovimentata,  "
sSQL = sSQL & "SUM(RV_PONumeroColli) AS NumeroColli "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)
If Not rs.EOF Then
    GET_LAVORAZIONE = GET_LAVORAZIONE + fnNotNullN(rs!QuantitaMovimentata)
    Me.txtNumeroColliLavorati.Value = Me.txtNumeroColliLavorati.Value + fnNotNullN(rs!NumeroColli)
    DoEvents
End If
    
rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_LAVORAZIONE:
    GET_LAVORAZIONE = GET_LAVORAZIONE
End Function

Private Function GET_PROCESSO(IDRigaConferimento As Long) As Double
On Error GoTo ERR_GET_PROCESSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

GET_PROCESSO = 0

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT SUM(RV_POQuantitaMovimentata) as QuantitaMovimentata "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POIVGamma")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_PROCESSO = GET_PROCESSO + fnNotNullN(rs!QuantitaMovimentata)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_PROCESSO:
    GET_PROCESSO = GET_PROCESSO
End Function
Private Function GET_VENDITA(IDRigaConferimento As Long) As Double
On Error GoTo ERR_GET_VENDITA

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_VENDITA = 0
Me.txtNumeroColliVenduti.Value = 0
''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT SUM(RV_POQuantitaMovimentata) AS QuantitaMovimentata, "
sSQL = sSQL & "SUM(RV_PONumeroColli) AS NumeroColli "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_VENDITA = GET_VENDITA + fnNotNullN(rs!QuantitaMovimentata)
    Me.txtNumeroColliVenduti.Value = Me.txtNumeroColliVenduti.Value + fnNotNullN(rs!NumeroColli)
    DoEvents
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_VENDITA:
    GET_VENDITA = GET_VENDITA
End Function
Private Function GET_TIPO_PRODOTTO(IDArticolo As Long) As Long
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PRODOTTO = 0
Else
    GET_TIPO_PRODOTTO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GrigliaCorpo_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
If VISUALIZZA_RIEP_CONF = 1 Then
    Me.txtUnitaDiMisura.Text = fnNotNull(Me.GrigliaCorpo.AllColumns("UnitaDiMisuraCoop").Value)
    Me.txtQtaConferita.Value = fnNotNullN(Me.GrigliaCorpo.AllColumns("Qta_UM").Value)
    Me.txtNumeroColliConferiti.Value = fnNotNullN(Me.GrigliaCorpo.AllColumns("Colli").Value)
    
    Me.txtQtaLavorata.Value = GET_LAVORAZIONE(fnNotNullN(Me.GrigliaCorpo.AllColumns("IDRV_POCaricoMerceRighe").Value))
    Me.txtQtaVenduta.Value = GET_VENDITA(fnNotNullN(Me.GrigliaCorpo.AllColumns("IDRV_POCaricoMerceRighe").Value))
    Me.txtQtaProcesso.Value = GET_PROCESSO(fnNotNullN(Me.GrigliaCorpo.AllColumns("IDRV_POCaricoMerceRighe").Value))
    Me.txtQtaQuadrata.Value = GET_QUADRATURA(fnNotNullN(Me.GrigliaCorpo.AllColumns("IDRV_POCaricoMerceRighe").Value))
    
    Me.txtQtaDiffLav.Value = Me.txtQtaConferita.Value - Me.txtQtaLavorata - Me.txtQtaQuadrata.Value - Me.txtQtaProcesso.Value
    Me.txtQtaDiffVend.Value = Me.txtQtaConferita.Value - Me.txtQtaVenduta.Value - Me.txtQtaQuadrata.Value


End If
End Sub

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
On Error Resume Next
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
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub txtCodiceArticolo_Change()
    GET_SQL
End Sub
Private Sub GET_SQL()
On Error GoTo ERR_GET_SQL
Dim Filtro As String

Filtro = ""

If Len(Trim(Me.txtCodiceArticolo.Text)) > 0 Then
    If Len(Filtro) > 0 Then
        Filtro = Filtro & " AND "
    End If
    
    Filtro = Filtro & "CodiceArticolo LIKE " & fnNormString("%" & Me.txtCodiceArticolo.Text & "%")
End If
If Len(Trim(Me.txtDescrizioneArticolo.Text)) > 0 Then
    If Len(Filtro) > 0 Then
        Filtro = Filtro & " AND "
    End If
    
    Filtro = Filtro & "Articolo LIKE " & fnNormString("%" & Me.txtDescrizioneArticolo.Text & "%")
End If
If Len(Trim(Me.txtLottoEntrata.Text)) > 0 Then
    If Len(Filtro) > 0 Then
        Filtro = Filtro & " AND "
    End If
    
    Filtro = Filtro & "CodiceLotto LIKE " & fnNormString("%" & Me.txtLottoEntrata.Text & "%")
End If
If Me.cboCatMerc.CurrentID > 0 Then
    If Len(Filtro) > 0 Then
        Filtro = Filtro & " AND "
    End If
    
    Filtro = Filtro & "IDCategoriaMerceologica=" & Me.cboCatMerc.CurrentID
    
End If
If Me.ACSSocio.IDAnagrafica > 0 Then
    If Len(Filtro) > 0 Then
        Filtro = Filtro & " AND "
    End If
    
    Filtro = Filtro & "IDAnagrafica=" & Me.ACSSocio.IDAnagrafica
    
End If
rsGriglia.Filter = Filtro

Me.GrigliaCorpo.Refresh
Exit Sub
ERR_GET_SQL:
    MsgBox Err.Description, vbCritical, "GET_SQL"
End Sub

Private Sub txtDescrizioneArticolo_Change()
    GET_SQL
End Sub

Private Sub txtLottoEntrata_Change()
    GET_SQL
End Sub
