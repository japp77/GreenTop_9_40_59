VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Begin VB.Form frmSpezzaturaVeloce 
   Caption         =   "Spezzatura veloce"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpezzaturaVeloce.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   20280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRaggrOrdine 
      Height          =   315
      Left            =   9480
      TabIndex        =   29
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtRicercaCodArt 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame fraElaborazione 
      Caption         =   "ELABORAZIONE"
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
      Left            =   3600
      TabIndex        =   21
      Top             =   6720
      Width           =   16575
      Begin VB.CommandButton cmdConferma 
         Caption         =   "CONFERMA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   13320
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   10095
      End
   End
   Begin VB.Frame fraPedana 
      Caption         =   "Pedana"
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
      Height          =   975
      Left            =   3600
      TabIndex        =   15
      Top             =   5760
      Width           =   16575
      Begin VB.TextBox txtSubLotto 
         Height          =   315
         Left            =   13800
         TabIndex        =   31
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdSelezionaPedana 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Picture         =   "frmSpezzaturaVeloce.frx":4781A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Trova pedana"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNuovaPedana 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         Picture         =   "frmSpezzaturaVeloce.frx":47DA4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Nuova pedana"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox txtCodicePedana 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   490
         Width           =   1935
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoPedana 
         Height          =   315
         Left            =   7200
         TabIndex        =   16
         Top             =   480
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTLblLinkCtl.LabelLink LabelLink2 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "Pedana"
         Name            =   "LabelLink"
      End
      Begin DmtCodDescCtl.DmtCodDesc CDTipoPedana 
         Height          =   615
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         PropCodice      =   $"frmSpezzaturaVeloce.frx":4832E
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmSpezzaturaVeloce.frx":4837D
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmSpezzaturaVeloce.frx":483D3
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
      Begin DMTEDITNUMLib.dmtNumber txtIDPedana 
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticoloScarto 
         Height          =   615
         Left            =   8280
         TabIndex        =   27
         Top             =   230
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1085
         PropCodice      =   $"frmSpezzaturaVeloce.frx":4842D
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmSpezzaturaVeloce.frx":4847C
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmSpezzaturaVeloce.frx":484EA
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
      Begin VB.Label Label3 
         Caption         =   "Sub lotto"
         Height          =   255
         Index           =   3
         Left            =   13800
         TabIndex        =   32
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Peso"
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtRicercaArticolo 
      Height          =   315
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtRicercaPedana 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame fraTotali 
      Caption         =   "Totale pedana"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   3375
      Begin DMTEDITNUMLib.dmtNumber txtTotaleColliSel 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoSel 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
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
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoReale 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotalePezziSel 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Totale pezzi selezionati"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Peso lordo riscontrato"
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
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Totale peso lordo selezionato"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Totale colli selezionati"
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
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   20055
      _ExtentX        =   35375
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
   Begin VB.Label Label3 
      Caption         =   "Sub lotto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   30
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Ricerca codice articolo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   25
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Ricerca articolo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Ricerca pedana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmSpezzaturaVeloce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLav As ADODB.Recordset
Private rsLavSomma As ADODB.Recordset
Private rsLavMov As ADODB.Recordset
Private rsQuad As ADODB.Recordset

Private Link_Pedana As Long

Private Link_TipoImballo As Long
Private Link_TipoScarto As Long
Private Link_TipoAumentoPeso As Long
Private Link_TipoCaloPeso As Long
Private Link_TipoProdotto_Q As Long

Private Link_Tipo_Oggetto_Lav As Long
Private Link_Tipo_Oggetto_Quad As Long
Private Link_Funzione_Carico As Long
Private Link_Funzione_Scarico As Long
Private Link_Funzione_Carico_IVGamma As Long
Private Link_Funzione_Scarico_IVGamma As Long

Private Link_Articolo_Neg As Long
Private Link_Articolo_Pos As Long

Private mov As DmtMovim.cMovimentazione

Public rsLottoImballoPrim As ADODB.Recordset
Public rsLottoImballo As ADODB.Recordset
Public rsKIT As ADODB.Recordset

Private MAX_QUANTITA_COLLI As Double
Private MAX_QUANTITA_CONFEZ As Double

Private Sub CDTipoPedana_ChangeElement()
    Me.txtPesoPedana.Value = GET_PESO_PEDANA(Me.txtIDPedana.Value, Me.CDTipoPedana.KeyFieldID)
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim Testo As String

If Me.txtTotaleColliSel.Value = 0 Then Exit Sub

If Me.txtIDPedana.Value = 0 Then
    MsgBox "Inserire la pedana", vbInformation, "Validazione dati"
    Exit Sub
End If

If Me.txtTotalePesoLordoReale.Value = 0 Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Il peso reale della pedana è a zero " & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Validazione dati") = vbNo Then Exit Sub
End If

If Me.txtTotalePesoLordoReale.Value <> Me.txtTotalePesoLordoSel.Value Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Il peso reale della pedana è diverso del peso dei colli selezionati " & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Validazione dati") = vbNo Then Exit Sub
End If

ELABORAZIONE

If PREZZI_ARTICOLI_DA_ORDINE = 1 Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "È stato impostato di prelevare gli importi dall'ordine, pertanto prima di procedere alla conferma dell'ordine eseguire la prezzatura veloce dal comando 'PREZZATURA DA ORDINE'"
    
    MsgBox Testo, vbInformation, "Prezzatura da ordine"
End If


Unload Me

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Private Sub cmdNuovaPedana_Click()
On Error GoTo ERR_cmdNuovaPedana_Click
If Me.txtIDPedana.Value > 0 Then
    If MsgBox("Questa riga ha già assegnata una pedana." & vbCrLf & "Vuoi crearla una nuova?", vbYesNo + vbQuestion, "Creazione nuova pedana") = vbNo Then Exit Sub
End If

Me.txtCodicePedana.Text = GetNumeroPedana(DatePart("yyyy", Date))
Me.txtIDPedana.Value = Link_Pedana
Me.CDTipoPedana.Load GET_TIPO_PEDANA(fnNotNullN(Link_Pedana))

Exit Sub
ERR_cmdNuovaPedana_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovaPedana_Click"
    
End Sub

Private Sub cmdSelezionaPedana_Click()
On Error GoTo ERR_cmdSelezionaPedana_Click
    WHERE_TROVA_PEDANA = 3
    frmTrovaPedana.Show vbModal
    Me.CDTipoPedana.Load fnNotNullN(Me.txtIDPedana.Value)

Exit Sub
ERR_cmdSelezionaPedana_Click:
MsgBox Err.Description, vbCritical, "cmdSelezionaPedana_Click"
End Sub

Private Sub Form_Load()
    
    INIT_CONTROLLI
    
    CREA_RECORDSET
    
    GET_GRIGLIA

End Sub
Private Sub CREA_RECORDSET()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient

Set rsLavSomma = New ADODB.Recordset
rsLavSomma.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=0 "

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsLav.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            rsLavSomma.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            
        Case adInteger
            rsLav.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsLavSomma.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsLav.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsLavSomma.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsLav.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
            rsLavSomma.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
        
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsLav.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
            rsLavSomma.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    
    End Select
Next

rsLav.Fields.Append "ColliSelezionati", adInteger, , adFldIsNullable
rsLav.Fields.Append "PesoLordoSelezionati", adDouble, , adFldIsNullable
rsLav.Fields.Append "PezziSelezionati", adDouble, , adFldIsNullable
rsLav.Fields.Append "Registra", adBoolean, , adFldIsNullable


rsLavSomma.Fields.Append "ColliSelezionati", adInteger, , adFldIsNullable
rsLavSomma.Fields.Append "PesoLordoSelezionati", adDouble, , adFldIsNullable
rsLavSomma.Fields.Append "PezziSelezionati", adDouble, , adFldIsNullable
rsLavSomma.Fields.Append "Registra", adBoolean, , adFldIsNullable

rs.Close
Set rs = Nothing

rsLav.Open , , adOpenKeyset, adLockBatchOptimistic
rsLavSomma.Open , , adOpenKeyset, adLockBatchOptimistic

'INSERIMENTO DATI
sSQL = "SELECT * FROM RV_POIEOrdineSmistamento "
sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND Qta_UM>0"

If FrmMain.CDAltroCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDCliente=" & FrmMain.CDAltroCliente.KeyFieldID
End If

If FrmMain.txtNumeroSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroOrdinePadre=" & FrmMain.txtNumeroSmistamento.Value
End If

If FrmMain.txtDataSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PODataOrdinePadre=" & fnNormDate(FrmMain.txtDataSmistamento.Text)
End If
If FrmMain.txtNListaSmistamento.Value > 0 Then
    sSQL = sSQL & " AND RV_PONumeroListaPrelievo=" & FrmMain.txtNListaSmistamento.Value
End If

If FrmMain.CDArticolo.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDArticolo=" & FrmMain.CDArticolo.KeyFieldID
End If

If Len(FrmMain.txtCodicePedana.Text) > 0 Then
    sSQL = sSQL & " AND CodicePedana LIKE " & fnNormString(FrmMain.txtCodicePedana.Text)
End If

If FrmMain.CDSocio.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & FrmMain.CDSocio.KeyFieldID
End If

If Len(Trim(FrmMain.txtLottoVendita.Text)) > 0 Then
    sSQL = sSQL & " AND CodiceLottoVendita LIKE " & fnNormString(FrmMain.txtLottoVendita.Text)
End If

If Len(Trim(Me.txtRaggrOrdine.Text)) > 0 Then
    sSQL = sSQL & " AND NotaRigaOrdRaggr LIKE " & fnNormString("%" & Me.txtRaggrOrdine.Text & "%")
End If

sSQL = sSQL & " ORDER BY CodicePedana, CodiceArticolo"

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    rsLav.AddNew
    rsLavSomma.AddNew
    
    For I = 0 To rs.Fields.Count - 1
        rsLav.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
        rsLavSomma.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
    Next
    rsLav!Registra = False
    rsLavSomma!Registra = False
    
    rsLav.Update
    rsLavSomma.Update
rs.MoveNext
Wend
rs.Close
Set rs = Nothing

'''''''''''LAVORAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not (rsLavMov Is Nothing) Then
    If rsLavMov.State > 0 Then
        rsLavMov.Close
    End If
    Set rsLavMov = Nothing
End If
Set rsLavMov = New ADODB.Recordset
rsLavMov.CursorLocation = adUseClient

rsLavMov.Fields.Append "IDRV_POAssegnazioneMerce", adInteger, , adFldIsNullable
rsLavMov.Fields.Append "IDRV_POCaricoMerceRighe", adInteger, , adFldIsNullable
rsLavMov.Fields.Append "NumeroColliMax", adDouble, , adFldIsNullable
rsLavMov.Fields.Append "NumeroConfezioniMax", adDouble, , adFldIsNullable
rsLavMov.Fields.Append "IDRV_POAssegnazioneMerceNew", adInteger, , adFldIsNullable

rsLavMov.Open , , adOpenKeyset, adLockBatchOptimistic

'''''''''QUADRATURA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not (rsQuad Is Nothing) Then
    If rsQuad.State > 0 Then
        rsQuad.Close
    End If
    Set rsQuad = Nothing
End If
Set rsQuad = New ADODB.Recordset
rsQuad.CursorLocation = adUseClient

rsQuad.Fields.Append "IDRV_POLavorazione", adInteger, , adFldIsNullable
rsQuad.Fields.Append "IDRV_POCaricoMerceRighe", adInteger, , adFldIsNullable

rsQuad.Open , , adOpenKeyset, adLockBatchOptimistic


End Sub

Private Sub GET_GRIGLIA()
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
        
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        
        Set cl = .ColumnsHeader.Add("Registra", "Sel.", dgBoolean, True, 1000, dgAligncenter, True, True, False)
            cl.Editable = True
            
        .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "ID", dgInteger, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodicePedana", "CodicePedana", dgchar, True, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 3000, dgAlignleft, True, True, False
        Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 900, dgAlignRight, True, True, False)
            'cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 1
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("ColliSelezionati", "Colli sel.", dgDouble, True, 900, dgAlignRight, True, True, False)
            cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 1
            cl.FormatOptions.FormatNumericThousandSep = "."
        
        Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, True, 1100, dgAlignRight, True, True, False)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
        Set cl = .ColumnsHeader.Add("PesoLordoSelezionati", "Peso lordo sel.", dgDouble, True, 1100, dgAlignRight, True, True, False)
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            
        Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgDouble, True, 1100, dgAlignRight, True, True, False)
            cl.BackColor = vbYellow
        Set cl = .ColumnsHeader.Add("PezziSelezionati", "Pezzi", dgDouble, True, 1100, dgAlignRight, True, True, False)
            cl.BackColor = vbGreen


        Set cl = .ColumnsHeader.Add("Qta_UM", "Quantità", dgDouble, False, 900, dgAlignRight, True, True, False)
            'cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 3
            cl.FormatOptions.FormatNumericThousandSep = "."
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

        Set cl = .ColumnsHeader.Add("NotaRigaOrdRaggr", "Raggr. ordine", dgchar, True, 1800, dgAlignleft)
            cl.Editable = True


        .ColumnsHeader.Add "IDImballoVendita", "IDImballo", dgNumeric, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imb.", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, False, 1000, dgAlignRight, True, True, False
        .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, False, 1000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, False, 1000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDCliente", "IDCliente", dgInteger, False, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "RV_PODataOrdinePadre", "Data ord.", dgDate, False, 1500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "RV_PONumeroOrdinePadre", "N° ord.", dgNumeric, False, 1500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "RV_PONumeroListaPrelievo", "N° lista", dgNumeric, False, 1500, dgAlignleft, True, True, False
        
        
        .ColumnsHeader.Add "Link_Doc_sezionale", "IDSezionale", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, False, 2500, dgAlignleft
        .ColumnsHeader.Add "Doc_data", "Data ord. ori.", dgDate, False, 1500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Doc_numero", "N° ord. ori.", dgNumeric, False, 1500, dgAlignleft, True, True, False
        
        
        
        .ColumnsHeader.Add "IDArticolo_Conferito", "IDArticolo_Conferito", dgInteger, False, 2000, dgAlignleft, True, True, False
        .ColumnsHeader.Add "CodiceArticolo_conferito", "Cod. Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Articolo_conferito", "Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDTipoLavorazione", "IDTipoLavorazione", dgNumeric, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDRV_POCalibro", "IDRV_POCalibro", dgNumeric, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Calibro", "Calibro", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDRV_POTipoCategoria", "IDRV_POTipoCategoria", dgNumeric, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, False, 1700, dgAlignleft, True, True, False
                                                                        
                                                                        
        .ColumnsHeader.Add "CodiceTipoPedana", "Cod. tipo pedana", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, True, 1700, dgAlignleft, False, True, False
        .ColumnsHeader.Add "CodiceArticoloPedana", "Codice articolo pedana", dgchar, True, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "ArticoloPedana", "Articolo pedana", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDRV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", dgInteger, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "AnnoProcesso", "Anno processo IV gamma", dgchar, False, 1700, dgAlignleft, True, True, False
        .ColumnsHeader.Add "NumeroProcesso", "Numero processo IV gamma", dgchar, False, 1700, dgAlignleft, True, True, False
                                                                        
                                                                        
        Set .Recordset = rsLav
        .LoadUserSettings
        .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor

End Sub

Private Sub Form_Resize()
On Error Resume Next

Me.Griglia.Width = Me.Width - 400
Me.fraPedana.Width = Me.Width - Me.fraPedana.Left - 280
Me.fraElaborazione.Width = Me.Width - Me.fraElaborazione.Left - 280


Me.cmdConferma.Left = Me.fraElaborazione.Width - Me.cmdConferma.Width - 120
Me.lblInfo.Width = Me.cmdConferma.Left - 240

End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
    
    If (Column.FieldName <> "ColliSelezionati") Then Exit Sub
    
    rsLav!PesoLordoSelezionati = (rsLav!PesoLordo / rsLav!Colli) * rsLav!ColliSelezionati
    rsLav!PezziSelezionati = (rsLav!Pezzi / rsLav!Colli) * rsLav!ColliSelezionati
    
    If (fnNotNullN(rsLav!ColliSelezionati) > 0) Then
        rsLav!Registra = True
    Else
        rsLav!Registra = False
    End If
    
    rsLav.Update
    
    Me.Griglia.Refresh
    
    
    rsLavSomma.Filter = "IDRV_POAssegnazioneMerce=" & fnNotNullN(rsLav!IDRV_POAssegnazioneMerce)
    
    If Not rsLavSomma.EOF Then
        
        rsLavSomma!Registra = rsLav!Registra
        
        rsLavSomma!ColliSelezionati = rsLav!ColliSelezionati
        
        rsLavSomma!PesoLordoSelezionati = rsLav!PesoLordoSelezionati
        
        rsLavSomma!PezziSelezionati = rsLav!PezziSelezionati
        
        rsLavSomma.Update
    End If
        
    rsLavSomma.Filter = vbNullString
    
    
    GET_SOMMA_SEL
End Sub

Private Sub GET_SOMMA_SEL()
Dim TotaleColliSel As Double
Dim TotalePesoSel As Double
Dim TotalePezziSel As Double

    TotaleColliSel = 0
    TotalePesoSel = 0
    TotalePezziSel = 0
    rsLavSomma.Filter = "ColliSelezionati>0"
    
    While Not rsLavSomma.EOF
        TotaleColliSel = TotaleColliSel + fnNotNullN(rsLavSomma!ColliSelezionati)
        TotalePesoSel = TotalePesoSel + fnNotNullN(rsLavSomma!PesoLordoSelezionati)
        TotalePezziSel = TotalePezziSel + fnNotNullN(rsLavSomma!PezziSelezionati)
    rsLavSomma.MoveNext
    Wend
    
    rsLavSomma.Filter = vbNullString

    Me.txtTotaleColliSel.Value = TotaleColliSel
    Me.txtTotalePesoLordoSel.Value = TotalePesoSel
    Me.txtTotalePezziSel.Value = TotalePezziSel
    Me.txtTotalePesoLordoReale.Value = TotalePesoSel
    
    
End Sub
Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsLav.Fields("Registra").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsLav.EOF And Not rsLav.BOF Then
        rsLav.Fields("Registra").Value = Abs(CLng(Selected))
        
        If rsLav!Registra = True Then
            If fnNotNullN(rsLav!ColliSelezionati) = 0 Then
                rsLav!ColliSelezionati = rsLav!Colli
                rsLav!PesoLordoSelezionati = (rsLav!PesoLordo / rsLav!Colli) * rsLav!ColliSelezionati
                rsLav!PezziSelezionati = (rsLav!Pezzi / rsLav!Colli) * rsLav!ColliSelezionati
            End If
        Else
            
            rsLav!ColliSelezionati = 0
            rsLav!PesoLordoSelezionati = 0
            rsLav!PezziSelezionati = 0
            
        End If
        rsLav.Update
        Me.Griglia.Refresh

        rsLavSomma.Filter = "IDRV_POAssegnazioneMerce=" & fnNotNullN(rsLav!IDRV_POAssegnazioneMerce)
        
    
        If Not rsLavSomma.EOF Then
            
            rsLavSomma!Registra = rsLav!Registra
            
            rsLavSomma!ColliSelezionati = rsLav!ColliSelezionati
            
            rsLavSomma!PesoLordoSelezionati = rsLav!PesoLordoSelezionati
            
            rsLavSomma!PezziSelezionati = rsLav!PezziSelezionati
            
            rsLavSomma.Update
        End If
        
        rsLavSomma.Filter = vbNullString
    End If
    
    GET_SOMMA_SEL
End Sub

Private Sub txtRaggrOrdine_Change()
    GET_FILTRO
End Sub

Private Sub txtRicercaArticolo_Change()
    GET_FILTRO
End Sub

Private Sub txtRicercaCodArt_Change()
    GET_FILTRO
End Sub

Private Sub txtRicercaPedana_Change()
    GET_FILTRO
End Sub
Private Sub GET_FILTRO()
    rsLav.Filter = vbNullString
    
    If Len(Trim(Me.txtRicercaPedana.Text)) > 0 Then
        If rsLav.Filter = 0 Then
            rsLav.Filter = "CodicePedana LIKE " & fnNormString("%" & txtRicercaPedana.Text & "%")
        Else
            rsLav.Filter = rsLav.Filter & " AND CodicePedana LIKE " & fnNormString("%" & txtRicercaPedana.Text & "%")
        End If
    End If

    If Len(Trim(Me.txtRicercaArticolo.Text)) > 0 Then
        If rsLav.Filter = 0 Then
            rsLav.Filter = "Articolo LIKE " & fnNormString(txtRicercaArticolo.Text & "%")
        Else
            rsLav.Filter = rsLav.Filter & " AND Articolo LIKE " & fnNormString(txtRicercaArticolo.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtRicercaCodArt.Text)) > 0 Then
        If rsLav.Filter = 0 Then
            rsLav.Filter = "CodiceArticolo LIKE " & fnNormString(txtRicercaCodArt.Text & "%")
        Else
            rsLav.Filter = rsLav.Filter & " AND CodiceArticolo LIKE " & fnNormString(txtRicercaCodArt.Text & "%")
        End If
    End If
    If Len(Trim(Me.txtRaggrOrdine.Text)) > 0 Then
        If rsLav.Filter = 0 Then
            rsLav.Filter = "NotaRigaOrdRaggr LIKE " & fnNormString("%" & txtRaggrOrdine.Text & "%")
        Else
            rsLav.Filter = rsLav.Filter & " AND NotaRigaOrdRaggr LIKE " & fnNormString("%" & txtRaggrOrdine.Text & "%")
        End If
    End If
    GET_GRIGLIA
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

Private Sub INIT_CONTROLLI()
    ParametroImballo
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroTipoScarto
    ParametroAggiornaTipoLavorazioneDaConf
    
    Link_Tipo_Oggetto_Lav = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    Link_Tipo_Oggetto_Quad = fnGetTipoOggetto("RV_POLavorazioneL")
    Link_Funzione_Carico = GET_FUNZIONE_MAGAZZINO(10, 1)
    Link_Funzione_Scarico = GET_FUNZIONE_MAGAZZINO(10, 2)
    Link_Funzione_Carico_IVGamma = GET_FUNZIONE_MAGAZZINO(2, 1)
    Link_Funzione_Scarico_IVGamma = GET_FUNZIONE_MAGAZZINO(2, 2)
     
     'Tipo di Pedana
    With Me.CDTipoPedana
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceTipoPedana"
        .DescriptionField = "TipoPedana"
        .KeyField = "IDRV_POTipoPedana"
        .TableName = "RV_POIETipoPedana"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .CodeIsNumeric = False
    End With
End Sub
Private Function GET_PESO_PEDANA(IDPedana As Long, IDTipoPedana As Long) As Double
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

If GET_PESO_PEDANA > 0 Then Exit Function

sSQL = "SELECT * FROM RV_PORepTipoPedana "
sSQL = sSQL & "WHERE IDRV_POTipoPedana=" & IDTipoPedana
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_PEDANA = 0
Else
    GET_PESO_PEDANA = fnNotNullN(rs!Tara)
    If fnNotNullN(rs!NumeroPallet) > 1 Then
        GET_PESO_PEDANA = GET_PESO_PEDANA * fnNotNullN(rs!NumeroPallet)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub ELABORAZIONE()
On Error GoTo ERR_ELABORAZIONE
Dim IDLavorazione As Long
Dim IDLavorazioneNew As Long

Me.lblInfo.Caption = "ELABORAZIONE IN CORSO..."
DoEvents

Screen.MousePointer = 11

If Me.txtTotalePesoLordoSel.Value = 0 Then Me.txtTotalePesoLordoSel.Value = 1

rsLav.Filter = "Registra=" & fnNormBoolean(1)

AGGIORNA_TIPO_PEDANA Link_Pedana, Me.CDTipoPedana.KeyFieldID, Me.txtPesoPedana.Value

While Not rsLav.EOF
    
    IDLavorazione = fnNotNullN(rsLav!IDRV_POAssegnazioneMerce)
    IDLavorazioneNew = 0
    
    MAX_QUANTITA_COLLI = fnNotNullN(rsLav!Colli)
    MAX_QUANTITA_CONFEZ = fnNotNullN(rsLav!Colli) * fnNotNullN(rsLav!NumeroConfezioniPerImballo)
    
    If fnNotNullN(rsLav!Colli) = fnNotNullN(rsLav!ColliSelezionati) Then
        AGGIORNA_LAVORAZIONE IDLavorazione, rsLav
    Else
        IDLavorazioneNew = AGGIORNA_LAVORAZIONE_SPEZZATURA(IDLavorazione, rsLav)
    End If

    rsLavMov.AddNew
        rsLavMov!IDRV_POAssegnazioneMerce = fnNotNullN(rsLav!IDRV_POAssegnazioneMerce)
        rsLavMov!IDRV_POCaricoMerceRighe = fnNotNullN(rsLav!IDRV_POCaricoMerceRighe)
        rsLavMov!NumeroColliMax = MAX_QUANTITA_COLLI
        rsLavMov!NumeroConfezioniMax = MAX_QUANTITA_CONFEZ
        If IDLavorazioneNew > 0 Then
            rsLavMov!IDRV_POAssegnazioneMerceNew = IDLavorazioneNew
        End If
    rsLavMov.Update
    
'    If IDLavorazioneNew > 0 Then
'        rsLavMov.AddNew
'            rsLavMov!IDRV_POAssegnazioneMerce = IDLavorazioneNew
'            rsLavMov!IDRV_POCaricoMerceRighe = fnNotNullN(rsLav!IDRV_POCaricoMerceRighe)
'            rsLavMov!NumeroColliMax = MAX_QUANTITA_COLLI
'            rsLavMov!NumeroConfezioniMax = MAX_QUANTITA_CONFEZ
'        rsLavMov.Update
'    End If
    
    DoEvents
    
rsLav.MoveNext
Wend
Screen.MousePointer = 0

Me.lblInfo.Caption = "MOVIMENTAZIONE IN CORSO..."
DoEvents
Screen.MousePointer = 11
AVVIA_MOVIMENTAZIONE_LAVORAZIONE
Screen.MousePointer = 0

Exit Sub
ERR_ELABORAZIONE:
    MsgBox Err.Description, vbCritical, "ELABORAZIONE"
End Sub
Private Sub AGGIORNA_LAVORAZIONE(IDLavorazione As Long, rstmp As ADODB.Recordset)
On Error GoTo ERR_AGGIORNA_LAVORAZIONE
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    rs!IDOggettoOrdine = FrmMain.txtIDOrdine.Value
    rs!NumeroOrdine = FrmMain.txtNumeroOrdine.Value
    rs!DataOrdine = FrmMain.txtDataOrdine.Value
    rs!IDCliente = FrmMain.cdCliente.KeyFieldID
    rs!NumeroListaPrelievo = FrmMain.txtNListaPrelievo.Value
    rs!IDOggettoOrdinePadre = FrmMain.txtIDOrdinePadre.Value
    
    rs!IDRV_POPedana = Me.txtIDPedana.Value
    rs!CodicePedana = Me.txtCodicePedana.Text
    rs!Colli = rstmp!ColliSelezionati
    rs!PesoLordo = (Me.txtTotalePesoLordoReale.Value / Me.txtTotalePesoLordoSel.Value) * fnNotNullN(rstmp!PesoLordoSelezionati)
    rs!Tara = fnNotNullN(rstmp!TaraUnitaria) * fnNotNullN(rs!Colli)
    rs!PesoNetto = fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara)
    rs!Pezzi = fnNotNullN(rstmp!PezziSelezionati)
    
    Select Case fnNotNullN(rstmp!IDUnitaDiMisuraCoop)
        Case 1
            rs!Qta_UM = rs!Colli
        Case 2
            rs!Qta_UM = rs!PesoLordo
        Case 3
            rs!Qta_UM = rs!PesoNetto
        Case 4
            rs!Qta_UM = rs!Tara
        Case 5
            rs!Qta_UM = rs!Pezzi
        Case Else
            rs!Qta_UM = rs!PesoNetto
    End Select
    If Len(Trim(fnNotNull(txtSubLotto.Text))) = 0 Then
        rs!NotaRigaOrdRaggr = fnNotNull(rstmp!NotaRigaOrdRaggr)
    Else
        rs!NotaRigaOrdRaggr = Me.txtSubLotto.Text
    End If
    GET_CONFIGURAZIONE_IMPORTI_ARTICOLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_ARTICOLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
    
    GET_CONFIGURAZIONE_IMPORTI_IMBALLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_ARTICOLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
    
    rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDArticolo), FrmMain.txtIDOrdinePadre.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.cdCliente.KeyFieldID, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria))
    
    rs.Update
End If


rs.Close
Set rs = Nothing
Exit Sub
ERR_AGGIORNA_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "AGGIORNA_LAVORAZIONE"
End Sub
Private Sub GET_CONFIGURAZIONE_IMPORTI_ARTICOLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset, DataOrdine As String, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String
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


Set ObjDoc = New DmtDocs.cDocument
Set ObjDoc.Connection = TheApp.Database.Connection
ObjDoc.SetTipoOggetto 2
ObjDoc.IDFunzione = 105
ObjDoc.TablesNames ObjDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
ObjDoc.IDAzienda = TheApp.IDFirm
ObjDoc.IDFiliale = TheApp.Branch
ObjDoc.IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
ObjDoc.IDTipoAnagrafica = 2
ObjDoc.IDUtente = TheApp.IDUser
ObjDoc.DataEmissione = Date

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal


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


Set ObjDoc = Nothing
End Sub

Private Sub GET_CONFIGURAZIONE_IMPORTI_IMBALLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset, DataOrdine As String, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String
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


Set ObjDoc = New DmtDocs.cDocument
Set ObjDoc.Connection = TheApp.Database.Connection
ObjDoc.SetTipoOggetto 2
ObjDoc.IDFunzione = 105
ObjDoc.TablesNames ObjDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
ObjDoc.IDAzienda = TheApp.IDFirm
ObjDoc.IDFiliale = TheApp.Branch
ObjDoc.IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
ObjDoc.IDTipoAnagrafica = 2
ObjDoc.IDUtente = TheApp.IDUser
ObjDoc.DataEmissione = Date

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.ReadDataFromArticle IDImballo, sTabellaDettaglioLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal


ObjDoc.ReadDataFromPriceList IDListino
ObjDoc.ReadDataFromDiscountsList

If Quantita = 0 Then
    ObjDoc.Field "Art_quantita_totale", "0,01", sTabellaDettaglioLocal
Else
    ObjDoc.Field "Art_quantita_totale", Quantita, sTabellaDettaglioLocal
End If

rstmp!ImportoUnitarioImballo = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))



Set ObjDoc = Nothing
End Sub


Private Function GET_LINK_ATTIVITA_AZIENDA(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda "
sSQL = sSQL & "FROM Filiale "
sSQL = sSQL & "WHERE Filiale.IDFiliale = " & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

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

Private Function AGGIORNA_LAVORAZIONE_SPEZZATURA(IDLavorazione As Long, rstmp As ADODB.Recordset) As Long
On Error GoTo ERR_AGGIORNA_LAVORAZIONE_SPEZZATURA
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim IDLavorazioneNew As Long

AGGIORNA_LAVORAZIONE_SPEZZATURA = 0

Dim ColliNew As Long
Dim PesoLordoNew As Double
Dim PesoNettoNew As Double
Dim TaraNew As Double
Dim PezziNew As Double

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    rs.AddNew
        IDLavorazioneNew = fnGetNewKey("RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce")
        rs.Fields("IDRV_POAssegnazioneMerce").Value = IDLavorazioneNew
        rs.Fields("IDRV_POCaricoMerceRighe").Value = rstmp!IDRV_POCaricoMerceRighe
        rs.Fields("DataDocumento").Value = rstmp!DataDocumento
        rs.Fields("IDArticolo").Value = rstmp!IDArticolo
        rs.Fields("CodiceArticolo").Value = rstmp!CodiceArticolo
        rs.Fields("Articolo").Value = rstmp!Articolo
        rs.Fields("IDUnitaDiMisuraCoop").Value = rstmp!IDUnitaDiMisuraCoop
        rs.Fields("IDUnitaDiMisura").Value = rstmp!IDUnitaDiMisura
        rs.Fields("Colli").Value = rstmp!ColliSelezionati
        rs.Fields("PesoLordo").Value = (Me.txtTotalePesoLordoReale / Me.txtTotalePesoLordoSel) * fnNotNullN(rstmp!PesoLordoSelezionati)
        rs.Fields("Tara").Value = rs!Colli * rstmp!TaraUnitaria
        rs.Fields("PesoNetto").Value = rs.Fields("PesoLordo").Value - rs.Fields("Tara").Value
        rs.Fields("Pezzi").Value = rstmp!PezziSelezionati
        
        Select Case rs.Fields("IDUnitaDiMisuraCoop").Value
            Case 1
                rs.Fields("Qta_UM").Value = rs.Fields("Colli").Value
            Case 2
                rs.Fields("Qta_UM").Value = rs.Fields("PesoLordo").Value
            Case 3
                rs.Fields("Qta_UM").Value = rs.Fields("PesoNetto").Value
            Case 4
                rs.Fields("Qta_UM").Value = rs.Fields("Tara").Value
            Case 5
                rs.Fields("Qta_UM").Value = rs.Fields("Pezzi").Value
        End Select
        
        
        rs.Fields("IDImballoVendita").Value = rstmp!IDImballoVendita
        rs.Fields("CodiceImballoVendita").Value = rstmp!CodiceImballoVendita
        rs.Fields("ImballoVendita").Value = rstmp!ImballoVendita
        rs.Fields("TaraUnitaria").Value = rstmp!TaraUnitaria
        
        rs.Fields("IDTipoLavorazione").Value = rstmp!IDTipoLavorazione
        rs.Fields("IDRV_POCalibro").Value = rstmp!IDRV_POCalibro
        rs.Fields("IDRV_POTipoCategoria").Value = rstmp!IDRV_POTipoCategoria
        
        rs.Fields("IDCliente").Value = FrmMain.cdCliente.KeyFieldID
        rs.Fields("IDOggettoOrdine").Value = FrmMain.txtIDOrdine.Value
        rs.Fields("NumeroOrdine").Value = FrmMain.txtNumeroOrdine.Value
        rs.Fields("DataOrdine").Value = FrmMain.txtDataOrdine.Value
        rs.Fields("NumeroListaPrelievo").Value = FrmMain.txtNListaPrelievo.Value
        rs.Fields("IDOggettoOrdinePadre").Value = FrmMain.txtIDOrdinePadre.Value
        
        rs.Fields("IDRV_POPedana").Value = Me.txtIDPedana.Value
        rs.Fields("CodicePedana").Value = Me.txtCodicePedana.Text
        rs.Fields("CodiceLottoVendita").Value = rstmp!CodiceLottoVendita
        
        rs.Fields("IDAnagraficaSocio").Value = rstmp!IDAnagraficaSocio
        rs.Fields("CodiceSocio").Value = rstmp!CodiceSocio
        rs.Fields("AnagraficaSocio").Value = rstmp!AnagraficaSocio
        rs.Fields("NomeSocio").Value = rstmp!NomeSocio
        
        rs.Fields("DataConferimento").Value = rstmp!DataConferimento
        rs.Fields("NumeroConferimento").Value = rstmp!NumeroConferimento
        
        rs.Fields("LottoCliente").Value = rstmp!LottoCliente
        rs.Fields("AltreAnnotazioniPerCliente").Value = rstmp!AltreAnnotazioniPerCliente
        rs.Fields("IDUtente").Value = rstmp!IDUtente
        rs.Fields("CodiceUtente").Value = rstmp!CodiceUtente
        rs.Fields("NumeroPedanaCliente").Value = rstmp!NumeroPedanaCliente
        rs.Fields("MerceInclusoImballo").Value = rstmp!MerceInclusoImballo
        rs.Fields("UtentePC").Value = rstmp!UtentePC
        rs.Fields("NomePC").Value = rstmp!NomePC
        rs.Fields("IDLinguaPredefinita").Value = rstmp!IDLinguaPredefinita
        rs.Fields("IDLinguaCliente").Value = rstmp!IDLinguaCliente
        
        rs.Fields("CodicePedanaCliente").Value = rstmp!CodicePedanaCliente
        rs.Fields("CodiceGSI").Value = rstmp!CodiceGSI
        rs.Fields("CodiceAssociatoPressoCliente").Value = rstmp!CodiceAssociatoPressoCliente
        rs.Fields("CodiceABarreArticoloCliente").Value = rstmp!CodiceABarreArticoloCliente
        rs.Fields("DescrizioneCodiceABarreArticoloCliente").Value = rstmp!DescrizioneCodiceABarreArticoloCliente
        rs.Fields("CodiceABarreImballoCliente").Value = rstmp!CodiceABarreImballoCliente
        rs.Fields("DescrizioneCodiceABarreImballoCliente").Value = rstmp!DescrizioneCodiceABarreImballoCliente
        rs.Fields("DescrizioneArticoloInLinguaPred").Value = rstmp!DescrizioneArticoloInLinguaPred
        rs.Fields("DescrizioneCalibroInLinguaPred").Value = rstmp!DescrizioneCalibroInLinguaPred
        rs.Fields("DescrizioneCategoriaInLinguaPred").Value = rstmp!DescrizioneCategoriaInLinguaPred
        rs.Fields("DescrizioneArticoloInLinguaCliente").Value = rstmp!DescrizioneArticoloInLinguaCliente
        rs.Fields("DescrizioneCalibroInLinguaCliente").Value = rstmp!DescrizioneCalibroInLinguaCliente
        rs.Fields("DescrizioneCategoriaInLinguaCliente").Value = rstmp!DescrizioneCategoriaInLinguaCliente
        rs.Fields("CodiceABarreArticoloPred").Value = rstmp!CodiceABarreArticoloPred
        rs.Fields("DescrizioneCodiceABarreArticoloPred").Value = rstmp!DescrizioneCodiceABarreArticoloPred
        rs.Fields("CodiceABarreImballoPred").Value = rstmp!CodiceABarreImballoPred
        rs.Fields("DescrizioneCodiceABarreImballoPred").Value = rstmp!DescrizioneCodiceABarreImballoPred
        rs.Fields("LinguaPredefinita").Value = rstmp!LinguaPredefinita
        rs.Fields("LinguaCliente").Value = rstmp!LinguaCliente
        rs.Fields("Link_Ordinamento").Value = GET_LINK_ORDINAMENTO("RV_POAssegnazioneMerce", "IDRV_POCaricoMerceRighe", fnNotNullN(rs.Fields("IDRV_POCaricoMerceRighe").Value))
        rs!IDRV_POProcessoIVGamma = rstmp!IDRV_POProcessoIVGamma
        rs!OraLavorazione = rstmp!OraLavorazione
        rs!NotaRigaOrdRaggr = rstmp!NotaRigaOrdRaggr
        rs!IDArticoloImballoPrimario = rstmp!IDArticoloImballoPrimario
        rs!NumeroConfezioniPerImballo = rstmp!NumeroConfezioniPerImballo
        rs!TaraConfezioneImballo = rstmp!TaraConfezioneImballo
        rs!CostoConfezioneImballo = rstmp!CostoConfezioneImballo
        rs!TracciaImballoPrim = rstmp!TracciaImballoPrim
        rs!ConfermaDaUtentePrim = rstmp!ConfermaDaUtentePrim
        rs!TracciaImballo = rstmp!TracciaImballo
        rs!ConfermaDaUtente = rstmp!ConfermaDaUtente
        rs!QuantitaPerCollo = rstmp!QuantitaPerCollo
        rs!PesoPerCollo = rstmp!PesoPerCollo
        rs!MoltiplicatorePerCollo = rstmp!MoltiplicatorePerCollo
        rs!IDRV_POProcessoIVGamma = rstmp!IDRV_POProcessoIVGamma
        rs!IDOggettoOrdinePrec = rstmp!IDOggettoOrdinePrec
        rs!IDRV_POProcessoLavorazione = rstmp!IDRV_POProcessoLavorazione
        rs!IDRV_POProcessoLavorazioneRighe = rstmp!IDRV_POProcessoLavorazioneRighe
        rs!IDRV_POLineaProduzione = rstmp!IDRV_POLineaProduzione
        rs!IDRV_POCaricoMerceRighePrelievi = rstmp!IDRV_POCaricoMerceRighePrelievi
        rs!IDRV_POTipoUtilizzoLinea = rstmp!IDRV_POTipoUtilizzoLinea
        rs!IDRV_PO01_LottoCampagna = rstmp!IDRV_PO01_LottoCampagna
        rs!PreConferimento = rstmp!PreConferimento
        
        If Len(Trim(fnNotNull(txtSubLotto.Text))) = 0 Then
            rs!NotaRigaOrdRaggr = fnNotNull(rstmp!NotaRigaOrdRaggr)
        Else
            rs!NotaRigaOrdRaggr = Me.txtSubLotto.Text
        End If
        
        GET_CONFIGURAZIONE_IMPORTI_ARTICOLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_ARTICOLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
        GET_CONFIGURAZIONE_IMPORTI_IMBALLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_IMBALLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
        rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDArticolo), FrmMain.txtIDOrdinePadre.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.cdCliente.KeyFieldID, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria))
    
    rs.Update
    
    PesoLordoNew = rs!PesoLordo
    PesoNettoNew = rs!PesoNetto
    TaraNew = rs!Tara
        
    CREA_RECORDSET_KIT fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!IDArticoloImballoPrimario), IDLavorazioneNew, IDLavorazione
    
    SALVA_KIT IDLavorazioneNew
        
rs.Close
Set rs = Nothing

AGGIORNA_LAVORAZIONE_SPEZZATURA = IDLavorazioneNew

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione
Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    rs!Colli = rs!Colli - rstmp!ColliSelezionati
    rs!PesoLordo = rs!PesoLordo - PesoLordoNew
    rs!Tara = rs!TaraUnitaria * rs!Colli
    rs!PesoNetto = rs!PesoLordo - rs!Tara
    rs!Pezzi = rs!Pezzi - rstmp!PezziSelezionati
    
    Select Case fnNotNullN(rs!IDUnitaDiMisuraCoop)
        Case 1
            rs!Qta_UM = rs!Colli
        Case 2
            rs!Qta_UM = rs!PesoLordo
        Case 3
            rs!Qta_UM = rs!PesoNetto
        Case 4
            rs!Qta_UM = rs!Tara
        Case 5
            rs!Qta_UM = rs!Pezzi
    End Select
    
    rs.Update
End If
    
CREA_RECORDSET_KIT fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!IDArticoloImballoPrimario), IDLavorazione, IDLavorazione

SALVA_KIT IDLavorazione
    
rs.Close
Set rs = Nothing

Exit Function
ERR_AGGIORNA_LAVORAZIONE_SPEZZATURA:
    MsgBox Err.Description, vbCritical, "AGGIORNA_LAVORAZIONE_SPEZZATURA"
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
Private Function AVVIA_MOVIMENTAZIONE_LAVORAZIONE()
On Error GoTo ERR_AVVIA_MOVIMENTAZIONE_LAVORAZIONE

If ((rsLavMov.EOF) And (rsLavMov.BOF)) Then Exit Function

rsLavMov.MoveFirst

While Not rsLavMov.EOF

    CREA_RECORDSET_LOTTI_IMBALLI fnGetTipoOggetto("RV_POAssegnazioneMerce"), fnNotNullN(rsLavMov!IDRV_POAssegnazioneMerce)
    
    MOVIMENTAZIONE_RIGA_LAVORAZIONE fnNotNullN(rsLavMov!IDRV_POCaricoMerceRighe), fnNotNullN(rsLavMov!IDRV_POAssegnazioneMerce), fnNotNullN(rsLavMov!NumeroColliMax), fnNotNullN(rsLavMov!NumeroConfezioniMax)
    
    If fnNotNullN(rsLavMov!IDRV_POAssegnazioneMerceNew) > 0 Then
        MOVIMENTAZIONE_RIGA_LAVORAZIONE fnNotNullN(rsLavMov!IDRV_POCaricoMerceRighe), fnNotNullN(rsLavMov!IDRV_POAssegnazioneMerceNew), fnNotNullN(rsLavMov!NumeroColliMax), fnNotNullN(rsLavMov!NumeroConfezioniMax)
    End If

rsLavMov.MoveNext
Wend

rsLavMov.Close
Set rsLavMov = Nothing

Exit Function

ERR_AVVIA_MOVIMENTAZIONE_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "AVVIA_MOVIMENTAZIONE_LAVORAZIONE"
End Function
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
End Sub
Private Sub ParametroTipoScarto()
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
End Sub

Private Function fncTrovaIDFunzione(Gestore As String, Optional Funzione As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE Gestore.Gestore = " & fnNormString(Gestore)
sSQL = sSQL & " AND Funzione = " & fnNormString(Funzione)
sSQL = sSQL & " AND Funzione.IDFunzione >= 10000"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
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
Public Function GeneraMovimentoCaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcesso As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoVendita, DataLavorazione As String) As Boolean
On Error GoTo ERR_GeneraMovimentoCaricoImballo

mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
'mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
'mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

If (IDProcesso) = 0 Then
    mov.IDFunzione = IDFunzione
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
Else
    mov.IDFunzione = Link_Funzione_Carico_IVGamma
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso)
End If

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 2
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0

mov.Field "RV_PODataLavorazione", ""
mov.Field "RV_POIDTipoLavorazione", 0
mov.Field "RV_POIDCalibro", 0
mov.Field "RV_POIDTipoCategoria", 0
mov.Field "RV_POIDTipoLavorazioneConf", 0
mov.Field "RV_POPrezzoMedioConf", 0

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

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

'CnDMT.BeginTrans
    GeneraMovimentoCaricoImballo = mov.Insert
'CnDMT.CommitTrans

Exit Function
ERR_GeneraMovimentoCaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function

Public Function GeneraMovimentoScaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcessoIVGamma As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoVendita, DataLavorazione As String, ColliMax As Double) As Boolean
On Error GoTo ERR_GeneraMovimentoScaricoImballo
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String
    
rsLottoImballo.Filter = "IDArticoloImballo=" & IDArticolo

If Not ((rsLottoImballo.EOF) And (rsLottoImballo.BOF)) Then

    rsLottoImballo.MoveFirst
    
    While Not rsLottoImballo.EOF
    
        mov.DataMovimento = DataLavorazione
        mov.FattoreDiConversione = Null
        
        mov.GestioneMatricole = False
        mov.IDEsercizio = IDEsercizio
        mov.IDTipoOggetto = IDTipoOggetto
        mov.IDOggetto = IDRigaConferimento
        mov.IDUtente = TheApp.IDUser
        mov.IDMagazzinoUscita = IDMagazzino
        mov.Cessione = 0
        mov.Field "IDAzienda", TheApp.IDFirm
        mov.Field "IDAnagrafica", IDAnagraficaSocio
        mov.Field "IDTipoAnagrafica", 3
        mov.Field "IDArticolo", IDArticolo
        mov.Field "IDUnitaDiMisura", IDUMDiamante
        mov.Field "IDcambio", Null
        mov.Field "DescrizioneArticolo", Articolo
        mov.Field "QuantitaTotale", (Qta_UM / ColliMax) * fnNotNullN(rsLottoImballo!QuantitaMovimentata)
        mov.Field "Importo", 0
        mov.Field "DataDocumento", DataLavorazione
        
        mov.Field "IDTipoMovimento", 1
        If (IDProcessoIVGamma = 0) Then
            mov.IDFunzione = IDFunzione
            mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
        Else
            mov.IDFunzione = Link_Funzione_Scarico_IVGamma
            mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcessoIVGamma)
        End If
        
        'DATI DI CONFERIMENTO
        mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
        mov.Field "RV_POTipoRiga", 2
        mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
        mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
        mov.Field "RV_POIDProcessoIVGamma", IDProcessoIVGamma
        mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
        mov.Field "RV_PODataConferimento", DataConferimento
        mov.Field "RV_PONumeroConferimento", NumeroConferimento
        mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
        mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
        mov.Field "RV_POCodiceLottoVendita", "" 'CodiceLottoVendita
        mov.Field "RV_POQuantitaLiquidazione", 0
        mov.Field "RV_POImportoInclusoImballo", 0
        mov.Field "RV_POImportoLiquidazione", 0
        mov.Field "RV_POQuantitaMovimentata", 0
        mov.Field "RV_PONumeroColli", 0
        mov.Field "RV_POPesoLordo", 0
        mov.Field "RV_POPesoNetto", 0
        mov.Field "RV_POTara", 0
        mov.Field "RV_POQuantitaPezzi", 0
        
        mov.Field "RV_PODataLavorazione", DataLavorazione
        mov.Field "RV_POIDTipoLavorazione", 0
        mov.Field "RV_POIDCalibro", 0
        mov.Field "RV_POIDTipoCategoria", 0
        mov.Field "RV_POIDTipoLavorazioneConf", 0
        mov.Field "RV_POPrezzoMedioConf", 0
        
        mov.Field "RV_POIDPedana", 0
        mov.Field "RV_POIDTipoPedana", 0
        mov.Field "RV_POCodicePedana", ""
        mov.Field "RV_POPesoPedana", 0
        
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
        
        GeneraMovimentoScaricoImballo = mov.Insert
    
    rsLottoImballo.MoveNext
    Wend

Else
    mov.DataMovimento = DataLavorazione
    mov.FattoreDiConversione = Null
    
    mov.GestioneMatricole = False
    mov.IDEsercizio = IDEsercizio
    mov.IDTipoOggetto = IDTipoOggetto
    mov.IDOggetto = IDRigaConferimento
    'mov.IDFunzione = IDFunzione
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = IDMagazzino
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", IDAnagraficaSocio
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", IDArticolo
    mov.Field "IDUnitaDiMisura", IDUMDiamante
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", Articolo
    mov.Field "QuantitaTotale", Qta_UM
    mov.Field "Importo", 0
    mov.Field "DataDocumento", DataLavorazione
    'mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
    mov.Field "IDTipoMovimento", 1
    If (IDProcessoIVGamma = 0) Then
        mov.IDFunzione = IDFunzione
        mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
    Else
        mov.IDFunzione = Link_Funzione_Scarico_IVGamma
        mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcessoIVGamma)
    End If
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    mov.Field "RV_POTipoRiga", 2
    mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
    mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
    mov.Field "RV_POIDProcessoIVGamma", 0
    mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
    mov.Field "RV_PODataConferimento", DataConferimento
    mov.Field "RV_PONumeroConferimento", NumeroConferimento
    mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
    mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
    mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
    mov.Field "RV_POQuantitaLiquidazione", 0
    mov.Field "RV_POImportoInclusoImballo", 0
    mov.Field "RV_POImportoLiquidazione", 0
    mov.Field "RV_POQuantitaMovimentata", 0
    mov.Field "RV_PONumeroColli", 0
    mov.Field "RV_POPesoLordo", 0
    mov.Field "RV_POPesoNetto", 0
    mov.Field "RV_POTara", 0
    mov.Field "RV_POQuantitaPezzi", 0
    
    mov.Field "RV_PODataLavorazione", ""
    mov.Field "RV_POIDTipoLavorazione", 0
    mov.Field "RV_POIDCalibro", 0
    mov.Field "RV_POIDTipoCategoria", 0
    mov.Field "RV_POIDTipoLavorazioneConf", 0
    mov.Field "RV_POPrezzoMedioConf", 0
    
    mov.Field "RV_POIDPedana", 0
    mov.Field "RV_POIDTipoPedana", 0
    mov.Field "RV_POCodicePedana", ""
    mov.Field "RV_POPesoPedana", 0
    
    
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
    
    'CnDMT.BeginTrans
        GeneraMovimentoScaricoImballo = mov.Insert
    'CnDMT.CommitTrans
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

Private Sub MOVIMENTAZIONE_RIGA_LAVORAZIONE(IDRigaConferimento As Long, IDAssegnazioneMerce As Long, ColliMax As Double, ConfezioniMax As Double)
On Error GoTo ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE
Dim OLD_Cursor As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsMov As DmtOleDbLib.adoResultset
Dim IDEsercizio As Long
Dim Movimentato As Long

OLD_Cursor = CnDMT.CursorLocation
CnDMT.CursorLocation = adUseClient
    
Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection

'''''''''''''''''''''''ELIMINAZIONE MOVIMENTI DELLA RIGA DI LAVORAZIONE'''''''''''''''''''''
sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
sSQL = sSQL & " AND IDOggetto=" & IDRigaConferimento
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Movimentato = 1

sSQL = "SELECT * FROM RV_POIEMovimentazioneLavorazioni "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
        
    IDEsercizio = fncEsercizio(fnNotNull(rs!DataDocumento))
        
    If GeneraMovimentoDiCarico(fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNull(rs!CodiceLottoVendita), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDUnitaDiMisura), _
    fnNotNull(rs!Articolo), fnNotNullN(rs!Qta_UM), fnNotNull(rs!DataDocumento), _
    fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
    fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), fnNotNullN(rs!IDRV_POPedana), _
    fnNotNullN(rs!IDRV_POTipoPedana), fnNotNull(rs!CodicePedana), fnNotNullN(rs!PesoPedana), fnNotNullN(rs!IDUnitaDiMisuraConf), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), _
    fnNotNullN(rs!NumeroConferimento), fnNotNull(rs!CodiceLottoConf), fnNotNullN(rs!IDMagazzinoVendita), IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Carico, _
    fnNotNullN(rs!IDArticoloImballoPrimario), fnNotNull(rs!CodiceArticoloPrimario), fnNotNull(rs!ArticoloPrimario), fnNotNullN(rs!NumeroConfezioniPerImballo), fnNotNullN(rs!TaraConfezioneImballo), 0) = False Then Movimentato = 0
    
    If GeneraMovimentoDiScarico(fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNull(rs!DataDocumento), fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNullN(IDAnagraficaSocio), _
    fnNotNullN(rs!IDArticoloConf), fnNotNull(rs!ArticoloConf), fnNotNullN(rs!IDUnitaDiMisuraDiamanteConf), fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), _
    fnNotNullN(rs!Pezzi), fnNotNullN(rs!IDUnitaDiMisuraConf), Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), IDEsercizio) = False Then Movimentato = 0
    
    If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballoVendita), GET_LINK_UM_ARTICOLO(fnNotNullN(rs!IDImballoVendita)), fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli), _
    IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Carico, fnNotNullN(rs!IDMagazzinoVendita), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento)) = False Then Movimentato = 0
    
    If GeneraMovimentoScaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballoVendita), GET_LINK_UM_ARTICOLO(fnNotNullN(rs!IDImballoVendita)), fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli), _
    IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento), ColliMax) = False Then Movimentato = 0
                
    If ((fnNotNullN(rs!IDArticoloImballoPrimario) > 0) And (fnNotNullN(rs!NumeroConfezioniPerImballo))) > 0 Then
        If GeneraMovimentoScaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticoloImballoPrimario), GET_LINK_UM_ARTICOLO(fnNotNullN(rs!IDArticoloImballoPrimario)), fnNotNull(rs!ArticoloPrimario), fnNotNullN(rs!Colli) * fnNotNullN(rs!NumeroConfezioniPerImballo), _
        IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento), ConfezioniMax) = False Then Movimentato = 0
    End If
                
    GeneraMovimentoScaricoKit fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POProcessoIVGamma), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), _
        IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento)
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

Private Function GeneraMovimentoDiCarico(IDAssegnazione As Long, IDRigaConferimento As Long, IDProcesso As Long, CodiceLottoVendita As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, DataLavorazione As String, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double, IDUMConferimentoCoop As Long, _
IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoEntrata As String, _
IDMagazzino As Long, IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDImballoPrim As Long, CodiceImballoPrim As String, DescrizioneImballoPrim As String, _
NumeroConfezioni As Long, TaraUnitariaConfezione As Double, CostoImballoConfezione As Double) As Boolean

On Error GoTo ERR_GeneraMovimentoDiCarico

mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento

mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione

mov.Field "IDTipoMovimento", 1
If (IDProcesso) = 0 Then
    mov.IDFunzione = IDFunzione
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
Else
    mov.IDFunzione = Link_Funzione_Carico_IVGamma
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso)
End If

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 1
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", IDProcesso
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_PONumeroColli", Colli
mov.Field "RV_POPesoLordo", PesoLordo
mov.Field "RV_POPesoNetto", PesoNetto
mov.Field "RV_POTara", Tara
mov.Field "RV_POQuantitaPezzi", Pezzi

Select Case IDUMConferimentoCoop
    Case 1
        mov.Field "RV_POQuantitaMovimentata", Colli
    Case 2
        mov.Field "RV_POQuantitaMovimentata", PesoLordo
    Case 3
        mov.Field "RV_POQuantitaMovimentata", PesoNetto
    Case 4
        mov.Field "RV_POQuantitaMovimentata", Tara
    Case 5
        mov.Field "RV_POQuantitaMovimentata", Pezzi
End Select

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

mov.Field "RV_POIDLottoImballo", 0
mov.Field "LottoImballo", ""

mov.Field "TipoRiga", trcNessuno

'CnDMT.BeginTrans
GeneraMovimentoDiCarico = mov.Insert
'CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoDiCarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function



Private Function GeneraMovimentoDiScarico(IDAssegnazione As Long, DataLavorazione As String, IDRigaConferimento As Long, IDProcesso As Long, _
IDAnagraficaSocio As Long, IDArticoloConferito As Long, ArticoloConferito As String, IDUnitaDiMisura As Long, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, IDUnitaDiMosuraConfCoop As Long, _
IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDEsercizio As Long) As Boolean

On Error GoTo ERR_GeneraMovimentoDiScarico


mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
'mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.IDMagazzinoUscita = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticoloConferito
mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", ArticoloConferito

Select Case IDUnitaDiMosuraConfCoop
    Case 1
        mov.Field "QuantitaTotale", Colli
    Case 2
        mov.Field "QuantitaTotale", PesoLordo
    Case 3
        mov.Field "QuantitaTotale", PesoNetto
    Case 4
        mov.Field "QuantitaTotale", Tara
    Case 5
        mov.Field "QuantitaTotale", Pezzi
End Select

mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
'mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

If (IDProcesso) = 0 Then
    mov.IDFunzione = IDFunzione
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
Else
    mov.IDFunzione = Link_Funzione_Scarico_IVGamma
    mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcesso)
End If

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


mov.Field "RV_PODataLavorazione", ""
mov.Field "RV_POIDTipoLavorazione", 0
mov.Field "RV_POIDCalibro", 0
mov.Field "RV_POIDTipoCategoria", 0
mov.Field "RV_POIDTipoLavorazioneConf", 0
mov.Field "RV_POPrezzoMedioConf", 0

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

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
'CnDMT.BeginTrans
GeneraMovimentoDiScarico = mov.Insert
'CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoDiScarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans

End Function

Private Sub ParametroAggiornaTipoLavorazioneDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaTipoLavDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_TIPO_LAVORAZIONE = fnNotNullN(rs!AggiornaTipoLavDaConf)
Else
    AGGIORNA_TIPO_LAVORAZIONE = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

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
Dim DescrizioneErrore As String

'VALORI DALLA LAVORAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    DescrizioneErrore = "VALORI DELLA LAVORAZIONE"
    
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
    
    DescrizioneErrore = "CREA RECORDSET KIT"
    
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
    
    DescrizioneErrore = "INSERIMENTO ARTICOLI DA SCARICARE SEMPRE"
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
    
    DescrizioneErrore = "INSERIMENTO ARTICOLI DA SCARICARE KIT"
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
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_KIT (" & DescrizioneErrore & ")"
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
Public Function GeneraMovimentoScaricoKit(IDRigaConferimento As Long, IDAssegnazione As Long, IDProcessoIVGamma As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, _
IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoVendita, DataLavorazione As String) As Boolean
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
            
    mov.DataMovimento = DataLavorazione
    mov.FattoreDiConversione = Null
    
    mov.GestioneMatricole = False
    mov.IDEsercizio = IDEsercizio
    mov.IDTipoOggetto = IDTipoOggetto
    mov.IDOggetto = IDRigaConferimento
    mov.IDUtente = TheApp.IDUser
    mov.IDMagazzinoUscita = IDMagazzino
    mov.Cessione = 0
    mov.Field "IDAzienda", TheApp.IDFirm
    mov.Field "IDAnagrafica", IDAnagraficaSocio
    mov.Field "IDTipoAnagrafica", 3
    mov.Field "IDArticolo", fnNotNullN(rs!IDArticolo) 'IDArticolo
    mov.Field "IDUnitaDiMisura", fnNotNullN(rs!IDUnitaDiMisura)
    mov.Field "IDcambio", Null
    mov.Field "DescrizioneArticolo", fnNotNull(rs!Articolo)
    mov.Field "QuantitaTotale", QuantitaRimasta 'Me.txtColli.Value
    mov.Field "Importo", 0
    mov.Field "DataDocumento", DataLavorazione
    mov.Field "IDTipoMovimento", 1
    If IDProcessoIVGamma = 0 Then
        mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
    Else
        mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione & " da processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcessoIVGamma) ' & Me.txtAnnoProcesso.Value & "-" & Me.txtNumeroProcesso.Value
        mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(2, 2)
    End If
    mov.Field "IDTipoMovimento", 1
    
    'DATI DI CONFERIMENTO
    mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    mov.Field "RV_POTipoRiga", 2
    mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
    mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
    mov.Field "RV_POIDProcessoIVGamma", IDProcessoIVGamma
    mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
    mov.Field "RV_PODataConferimento", DataConferimento
    mov.Field "RV_PONumeroConferimento", NumeroConferimento
    mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
    mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
    mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
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
    'CnDMT.BeginTrans
    GeneraMovimentoScaricoKit = mov.Insert
    'CnDMT.CommitTrans
            
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

