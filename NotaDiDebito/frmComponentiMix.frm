VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmComponentiMix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Componenti MIX"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19740
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComponentiMix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   19740
   StartUpPosition =   3  'Windows Default
   Begin DmtGridCtl.DmtGrid GrigliaCorpoOrdine 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   7646
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
End
Attribute VB_Name = "frmComponentiMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET

GET_GRIGLIA

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * FROM RV_POIEComponentiMixInDocVend "
sSQL = sSQL & "WHERE IDRV_POProcessoIVGamma=" & Link_RigaProcessoIVGamma
Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection


With Me.GrigliaCorpoOrdine
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        
    .ColumnsHeader.Add "IDRV_POProcessoIVGammaRighe", "IDRV_POProcessoIVGammaRighe", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", dgNumeric, False, 500, dgAlignRight
    
    .ColumnsHeader.Add "IDArticoloLavorato", "IDArticoloLavorato", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceArticoloLavorato", "Codice articolo", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "DescrizioneArticoloLavorato", "Descrizione articolo", dgchar, True, 4500, dgAlignleft
    .ColumnsHeader.Add "IDRV_POGruppoArticoloPerEvasioneMix", "IDRV_POGruppoArticoloPerEvasioneMix", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "GruppoArticoloPerEvasioneMix", "Gruppo per evasione MIX", dgchar, False, 2500, dgAlignleft
    
    .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 4500, dgAlignleft
    .ColumnsHeader.Add "CodiceLottoProduzione", "Lotto di produzione", dgchar, False, 4500, dgAlignleft
    .ColumnsHeader.Add "CodiceLottoEntrata", "Lotto di entrata", dgchar, False, 4500, dgAlignleft
    
    .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "DescrizioneUnitaDiMisura", "U.M.", dgchar, False, 4500, dgAlignleft
    .ColumnsHeader.Add "IDUnitaDiMisuraCoop", "IDUnitaDiMisuraCoop", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "DescrizioneUnitaDiMisuraCoop", "U.M. coop", dgchar, False, 4500, dgAlignleft
        
    Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    .ColumnsHeader.Add "IDArticoloConferito", "IDArticoloConferito", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceArticoloConferito", "Codice articolo conf.", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "DescrizioneArticoloConferito", "Descrizione articolo conf.", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "Anagrafica", "Anagrafica socio", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "Nome", "Nome socio", dgchar, False, 2000, dgAlignleft
    .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, False, 2000, dgAlignleft
    .ColumnsHeader.Add "IDTipoDocumentoCoop", "IDTipoDocumentoCoop", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "Tipo documento", "DescrizioneTipoConferimento", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "Numero", "Numero conf.", dgInteger, False, 2500, dgAlignRight
    .ColumnsHeader.Add "DataDocumento", "Data documento conf.", dgDate, False, 2500, dgAlignleft
    .ColumnsHeader.Add "DataConferimento", "Data consegna conf.", dgDate, False, 2500, dgAlignleft
    .ColumnsHeader.Add "IDTipoDocumentoAcq", "IDTipoDocumentoAcq", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "DescrizioneTipoDocumentoAcq", "Tipo documento Acq.", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "NumeroDocumentoAcq", "Numero doc. acq", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "DataDocumentoAcq", "Data documento acq.", dgDate, False, 2500, dgAlignleft
    
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CREA_RECORDSET
End Sub

