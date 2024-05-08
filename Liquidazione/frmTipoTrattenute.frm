VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmTipoTrattenute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo trattenute applicate"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipoTrattenute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4683
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
Attribute VB_Name = "frmTipoTrattenute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsGriglia As DmtOleDbLib.adoResultset

Private Sub Form_Load()
    fncGriglia
End Sub
Private Sub fncGriglia()
    Dim sSQL As String
    Dim sSQL_WHERE As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    Dim I As Integer
    
    sSQL = "SELECT * FROM RV_POIELiquidazioneRigheTratt "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & LINK_LIQUIDAZIONE
    sSQL = sSQL & " AND IDArticoloVendita=" & frmMain.GrigliaTrattenuteEla.AllColumns("IDArticolo").Value
    sSQL = sSQL & " AND IDTipoOggetto=" & frmMain.GrigliaTrattenuteEla.AllColumns("IDTipoOggetto").Value
    sSQL = sSQL & " AND IDOggetto=" & frmMain.GrigliaTrattenuteEla.AllColumns("IDOggetto").Value
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & frmMain.GrigliaTrattenuteEla.AllColumns("IDValoriOggettoDettaglioArticolo").Value
    sSQL = sSQL & " ORDER BY TipoTrattenuta"
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        Set rsGriglia = Cn.OpenResultset(sSQL)
            Set rsevent = rsGriglia.Data
        
        With Me.GrigliaCorpo
            .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POLiquidazioneRigheTratt", "IDRV_POLiquidazioneRigheTratt", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POLiquidazione", "IDRV_POLiquidazione", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POLiquidazionePeriodo", "IDRV_POLiquidazionePeriodo", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDArticoloVendita", "IDArticoloVendita", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POTrattenutaPerLiquidazione", "IDRV_POTrattenutaPerLiquidazione", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POTipoTrattenuta", "IDRV_POTipoTrattenuta", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "TipoTrattenuta", "Tipo trattenuta", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Tipo1", "Tipo1", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "Tipo2", "Tipo2", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "Tipo3", "Tipo3", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "Tipo4", "Tipo4", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDAnagraficaSocio", "IDAnagraficaSocio", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2500, dgAlignleft
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta1", "Val. Tratt. 1", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta1_Oggi", "Val. Tratt. 1 Oggi", dgDouble, False, 1300, dgAlignRight)
                cl.BackColor = vbRed
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
    
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta2", "Val. Tratt. 2", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta2_Oggi", "Val. Tratt. 2 Oggi", dgDouble, False, 1300, dgAlignRight)
                cl.BackColor = vbRed
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PercTrattenuta1", "% 1", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PercTrattenuta1_Oggi", "% 1 Oggi", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbRed
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PercTrattenuta2", "% 2", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PercTrattenuta2_Oggi", "% 2 Oggi", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbRed
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta1Conf", "Val. Tratt. Conf. 1", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbBlue
                cl.ForeColor = vbWhite
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta1Conf_Oggi", "Val. Tratt. Conf. 1 Oggi", dgDouble, False, 1300, dgAlignRight)
                cl.BackColor = vbRed
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta2Conf", "Val. Tratt. Conf. 2", dgDouble, True, 1300, dgAlignRight)
                cl.BackColor = vbBlue
                cl.ForeColor = vbWhite
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ValoreTrattenuta2Conf_Oggi", "Val. Tratt. Conf. 2 Oggi", dgDouble, False, 1300, dgAlignRight)
                cl.BackColor = vbRed

                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "SoloRigaConferimento", "Riga conf.", dgBoolean, False, 1000, dgAligncenter
            
            Set .Recordset = rsGriglia.Data
            .LoadUserSettings
            .Refresh
        End With
        
        Cn.CursorLocation = OLDCursor



End Sub
