VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmImbPrim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFEZIONI"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImbPrim.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gestione delle confezioni"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdScaricaImbPrim 
         Height          =   315
         Left            =   2880
         Picture         =   "frmImbPrim.frx":4781A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Altri articoli che compongono il kit"
         Top             =   1680
         Width           =   375
      End
      Begin DMTEDITNUMLib.dmtNumber txtCostoConfez 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DmtCodDescCtl.DmtCodDesc CDImballoPrimario 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1085
         PropCodice      =   $"frmImbPrim.frx":47DA4
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmImbPrim.frx":47DF3
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmImbPrim.frx":47E5B
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
      Begin DMTEDITNUMLib.dmtNumber txtCostoKit 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtCostoConfezLiq 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtCostoKitLiq 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
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
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerCollo 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1035
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   0
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoPerCollo 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   1035
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtMoltiplicatore 
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Top             =   1035
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Moltiplicatore"
         Height          =   255
         Index           =   17
         Left            =   2880
         TabIndex        =   17
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Peso per collo"
         Height          =   255
         Index           =   18
         Left            =   1680
         TabIndex        =   16
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà pezzi per collo"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Costo Kit liq."
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Costo confez. liq."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Costo Kit"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Costo confez."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmImbPrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub INIT_CONTROLLI()
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
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
End Sub

Private Sub cmdConferma_Click()

    CONFERMA
    
End Sub

Private Sub cmdScaricaImbPrim_Click()
    frmKIT.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    CONFERMA_IMBALLO_PRIM = 0
    
    INIT_CONTROLLI
    
    Me.CDImballoPrimario.Load IDImballoPrimario
    Me.txtCostoConfez.Value = CostoConfezione
    Me.txtCostoConfezLiq.Value = CostoConfezioneLiq
    Me.txtCostoKit.Value = GET_TOTALE_COSTO_KIT(Link_RigaAssegnazioneMerce)
    Me.txtCostoKitLiq.Value = CostoKitLiq
    Me.txtQuantitaPerCollo.Value = QUANTITA_PER_COLLO
    Me.txtPesoPerCollo.Value = PESO_LORDO
    Me.txtMoltiplicatore.Value = Moltiplicatore
    
End Sub

Private Function GET_TOTALE_COSTO_KIT(IDLavorazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(CostoTotaleRiga) as TotaleCosto "
sSQL = sSQL & "FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_COSTO_KIT = 0
Else
    GET_TOTALE_COSTO_KIT = fnNotNullN(rs!TotaleCosto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CONFERMA()

    IDImballoPrimario = Me.CDImballoPrimario.KeyFieldID
    CodiceImballoPrimario = Me.CDImballoPrimario.Code
    DescrizioneImballoPrimario = Me.CDImballoPrimario.Description
    CostoConfezione = Me.txtCostoConfez.Value
    CostoConfezioneLiq = Me.txtCostoConfezLiq.Value
    CostoKitLiq = Me.txtCostoKitLiq.Value
    QUANTITA_PER_COLLO = Me.txtQuantitaPerCollo.Value
    PESO_LORDO = Me.txtPesoPerCollo.Value
    Moltiplicatore = Me.txtMoltiplicatore.Value
    
    CONFERMA_IMBALLO_PRIM = 1
    
    Unload Me

End Sub


