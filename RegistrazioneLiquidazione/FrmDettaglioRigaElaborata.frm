VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form FrmDettaglioRigaElaborata 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DETTAGLIO LIQUIDAZIONE "
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid DmtGrid1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   10186
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
Attribute VB_Name = "FrmDettaglioRigaElaborata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
fncGriglia
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form Load"
End Sub
Private Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
sSQL = "SELECT * FROM RV_POTMPLiquidazioneRigheEla "
sSQL = sSQL & "WHERE RV_POTMPLiquidazioneRigheEla.IDRV_POTMPLiquidazione=" & LINK_DOCUMENTO_TMP_LIQ

If TIPO_LIQUIDAZIONE = 1 Then
    sSQL = sSQL & " ORDER BY DataConferimento , NumeroDocumento "
Else
    sSQL = sSQL & " ORDER BY DataDocumentoVendita"
End If

    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            
    
        
    
        With Me.DmtGrid1
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            
                    .ColumnsHeader.Add "IDRV_POLavorazione", "ID", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "NumeroDocumento", "N° conf.", dgInteger, True, 1000, dgAlignRight
                    .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceArticolo", "Cod. Prod.", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "Articolo", "Prodotto", dgchar, True, 2000, dgAlignleft
                    .ColumnsHeader.Add "IDLottoArticolo", "IDLotto", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceLottoArticolo", "Codice lotto", dgchar, False, 2000, dgAlignleft
                    .ColumnsHeader.Add "LottoArticolo", "Codice lotto", dgchar, False, 2000, dgAlignleft
                    Set cl = .ColumnsHeader.Add("QuantitaConferita", "Q.tà Conf.", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QuantitaLavorata", "Q.tà netta lav.", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QuantitaQuadrata", "Q.tà quad. lav.", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QuantitaTotaleLavorata", "Q.tà ToT. lav.", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QuantitaVenduta", "Q.tà Vend", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QuantitaQuadrataDiVendita", "Q.tà quad. Vend.", dgDouble, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("ImponibileDaReg", "Imp. riga", dgCurrency, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("TrattenuteTotali", "Tot. Tratt.", dgCurrency, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
End Sub
