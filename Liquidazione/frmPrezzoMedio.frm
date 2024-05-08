VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmPrezzoMedio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCOLO DEL PREZZO MEDIO"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13350
   Icon            =   "frmPrezzoMedio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   13350
   StartUpPosition =   1  'CenterOwner
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11668
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
Attribute VB_Name = "frmPrezzoMedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    fncGriglia LINK_RIGA_PREZZO_MEDIO
End Sub
Private Sub fncGriglia(IDRigaPrezzoMedio As Long)
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
Dim I As Integer
    
    sSQL = "SELECT * FROM RV_POIEPrezzoMedioLiq "
    sSQL = sSQL & "WHERE IDRV_POLiquidazioneRigheElaPM=" & IDRigaPrezzoMedio
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.Open sSQL, Cn.InternalConnection
        
        With Me.GrigliaCorpo
            
            .ColumnsHeader.Clear
            
            
            .ColumnsHeader.Add "IDRV_POLiquidazioneRigheElaPMRighe", "IDRV_POLiquidazioneRigheElaPMRighe", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POLiquidazioneRigheElaPM", "IDRV_POLiquidazioneRigheElaPM", dgNumeric, False, 500, dgAlignRight
            
            .ColumnsHeader.Add "IDRV_POLiquidazionePeriodo", "IDRV_POLiquidazionePeriodo", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgNumeric, False, 500, dgAlignRight
            
            .ColumnsHeader.Add "TipoDocumento", "Tipo documento", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "IDSezionale", "IDSezionale", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, False, 1200, dgAlignleft
            .ColumnsHeader.Add "Prefisso", "Prefisso", dgchar, False, 1000, dgAlignleft
            .ColumnsHeader.Add "DataDocumentoVendita", "Data doc.", dgDate, True, 1200, dgAlignleft
            .ColumnsHeader.Add "NumeroDocumentoVendita", "N° doc.", dgInteger, True, 800, dgAlignRight
            
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, False, 1800, dgAlignleft
            .ColumnsHeader.Add "IDCategoriaMerceologica", "IDCategoriaMerceologica", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "CategoriaMerceologica", "Categoria merc.", dgchar, True, 1800, dgAlignleft
            
            Set cl = .ColumnsHeader.Add("Quantita", "Q.ta P.M.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            Set cl = .ColumnsHeader.Add("QuantitaDocumento", "Q.ta doc.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            Set cl = .ColumnsHeader.Add("ImportoLiquidazione", "Imp. Liq.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
             Set cl = .ColumnsHeader.Add("ImportoSconti", "Sconti", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
             Set cl = .ColumnsHeader.Add("ImportoVariazionePrezzoImballo", "Var. Prz. Imb.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
             Set cl = .ColumnsHeader.Add("ImportoCommissioni", "Imp. Comm.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
             Set cl = .ColumnsHeader.Add("ImportoNettoVendita", "Imp. Netto Vend.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
                
            .ColumnsHeader.Add "DataVendita", "Data vend.", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DataLavorazione", "Data lav.", dgDate, True, 1800, dgAlignleft
          
            Set .Recordset = rsGriglia
            .LoadUserSettings
            .Refresh
        End With
        
        Cn.CursorLocation = OLDCursor



End Sub


