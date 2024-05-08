VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmAnalisiComm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Analisi commissioni per documento"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalisiComm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAnalisiDett 
      Caption         =   "Analisi dettagliata"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   14415
      Begin DmtGridCtl.DmtGrid GrigliaDett 
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   14175
         _ExtentX        =   25003
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
         UpdatePosition  =   0   'False
         ColumnsHeaderHeight=   20
      End
   End
   Begin VB.Frame fraTotaliComm 
      Caption         =   "Totali commissoni nel documento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin DmtGridCtl.DmtGrid GrigliaCorpo 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   3413
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
         ColumnsHeaderHeight=   20
      End
   End
End
Attribute VB_Name = "frmAnalisiComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsGrigliaDett As ADODB.Recordset


Private Sub GET_GRIGLIA()
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader


sSQL = "SELECT * FROM RV_POIECommissioniPerDocRigheRaggr "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

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
            .ColumnsHeader.Add "IDRV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POTipoRiga", "RV_POTipoRiga", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POTipoPedana", "IDRV_POTipoPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoCommissione", "Commissione", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "CodiceTipoPedanaComm", "Codice tipo pedana", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "TipoPedanaComm", "Descr. tipo pedana", dgchar, True, 2000, dgAlignleft
            Set cl = .ColumnsHeader.Add("NumeroPedana", "N° pedane", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 0
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("TotaleCommissioni", "Totali commissioni", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.BackColor = vbYellow
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
  
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor

End Sub

Private Sub Form_Load()
    GET_GRIGLIA
    
End Sub
Private Sub GET_GRIGLIA_DETT(IDCommissioneDoc As Long)
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader

sSQL = "SELECT * FROM RV_POIECommissioniPerDocRighe "
sSQL = sSQL & " WHERE IDRV_POCommissioniPerDoc=" & IDCommissioneDoc
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
Set rsGrigliaDett = New ADODB.Recordset
rsGrigliaDett.CursorLocation = adUseClient
rsGrigliaDett.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaDett
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDMovimento", "IDMovimento", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc", dgNumeric, False, 500, dgAlignleft
            
            .ColumnsHeader.Add "IDOggettoDoc", "IDOggettoDoc", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggettoDoc", "IDTipoOggettoDoc", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDValoriOggettoDettaglioDoc", "IDValoriOggettoDettaglioDoc", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POTipoRiga", "RV_POTipoRiga", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POTipoCommissione", "IDRV_POTipoCommissione", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoCommissione", "Commissione", dgchar, False, 2000, dgAlignleft
            
            
            .ColumnsHeader.Add "IDArticoloVenduto", "IDArticoloVenduto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticoloVenduto", "Codice articolo", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DescrizioneArticoloVenduto", "Articolo", dgchar, True, 2000, dgAlignleft
            
            Set cl = .ColumnsHeader.Add("Importo", "Imp. comm. uni.", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "UnitaDiMisuraVenduto", "U.M. vend.", dgchar, True, 2000, dgAlignleft
             Set cl = .ColumnsHeader.Add("QuantitaTotale", "Q.tà vend.", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "IDUnitaDiMisuraLiquidazione", "IDUnitaDiMisuraLiquidazione", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "DescrizioneUnitaDiMisuraLiquidazione", "U.M. Liq", dgchar, True, 2000, dgAlignleft
            
             Set cl = .ColumnsHeader.Add("Quantita", "Q.tà liq.", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .ColumnsHeader.Add("TotaleCommissione", "Totale riga", dgDouble, True, 1800, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            .ColumnsHeader.Add "RV_POIDPedana", "RV_POIDPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POCodicePedana", "Codice pedana", dgchar, True, 2000, dgAlignleft
            
            .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, False, 2000, dgAlignleft

  
        Set .Recordset = rsGrigliaDett
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor

End Sub

Private Sub GrigliaCorpo_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    GET_GRIGLIA_DETT fnNotNullN(GrigliaCorpo.AllColumns("IDRV_POCommissioniPerDoc").Value)
End Sub
