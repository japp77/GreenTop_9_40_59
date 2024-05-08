VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmCorpoOrdine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORPO ORDINE COLLEGATO"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCorpoOrdine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   19110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Non visualizzare le righe ordine completate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpoOrdine 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   9128
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
   Begin VB.Frame Frame1 
      Caption         =   "Note lavorazione da ordine"
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
      Height          =   5175
      Left            =   15480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtNoteLavDaOrd 
         Height          =   4815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmCorpoOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private VIS_NOTE_ORD_ELENCO As Long
Private bLoading As Boolean

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

'RECUPERO CAMPI
sSQL = "SELECT * FROM RV_POIEOrdineSelCorpo "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
        Case adInteger
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    End Select
Next

rsGriglia.Fields.Append "ImportoUnitarioImballo", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "ColliLavorati", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "DifferenzaColliLavorati", adDouble, , adFldIsNullable
rs.Close
Set rs = Nothing

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEOrdineSelCorpo "
sSQL = sSQL & "WHERE IDOggetto=" & LINK_ORDINE_PER_PREZZO
sSQL = sSQL & " AND RV_POTipoRiga=1 "
If MODALITA_RECUPERO_RIGA_ORD = 0 Then
    sSQL = sSQL & " AND Link_art_articolo=" & LINK_ARTICOLO_ORDINE
End If

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsGriglia.AddNew
        For I = 0 To rs.Fields.Count - 1
            rsGriglia.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
        Next
        
        rsGriglia!ImportoUnitarioImballo = GET_PREZZO_IMBALLO_DA_ORDINE(fnNotNullN(rsGriglia!RV_POIDImballo), fnNotNullN(rsGriglia!RV_POLinkRiga), fnNotNullN(rsGriglia!IDOggetto))
        rsGriglia!ColliLavorati = GET_COLLI_LAVORATI(fnNotNullN(rs!IDValoriOggettoDettaglio))
        rsGriglia!DifferenzaColliLavorati = fnNotNullN(rsGriglia!Art_numero_colli) - fnNotNullN(rsGriglia!ColliLavorati)
        
    rsGriglia.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

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

rsGriglia.Filter = vbNullString

If Me.Check1.Value = vbChecked Then
    rsGriglia.Filter = "DifferenzaColliLavorati>0"
End If

With Me.GrigliaCorpoOrdine
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
           
        .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Codice tipo pedana", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione codice pedana", dgchar, False, 3000, dgAlignleft
        
        .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Sub lotto", dgchar, True, 2500, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedana", "Q.tà pedana ord.", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedanaEffettiva", "Q.tà pedana", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_numero_colli", "Colli", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("ColliLavorati", "Colli lav.", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("DifferenzaColliLavorati", "Diff.", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        .ColumnsHeader.Add "RV_POIDTipoUMOrdine", "RV_POIDTipoUMOrdine", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POTipoUMOrdine", "Tipo U.M. riga ordine", dgchar, True, 2500, dgAlignleft

        Set cl = .ColumnsHeader.Add("Art_peso", "Peso lordo", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_tara", "Tara", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Imp. uni. merce", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_1", "% Sc1", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

         Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_2", "% Sc2", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, True, 3000, dgAlignleft
         Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Imp. uni. imb.", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POImportoImballoInArticolo", "Incluso Imballo", dgBoolean, False, 1300, dgAligncenter)
            cl.BackColor = vbYellow
        .ColumnsHeader.Add "RV_POIDImballoPrimario", "RV_POIDImballoPrimario", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceImballoPrimario", "Codice imballo primario", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballoPrimario", "Descrizione imballo primario", dgchar, True, 3000, dgAlignleft
        Set cl = .ColumnsHeader.Add("RV_POTaraImballoPrimario", "Tara imballo primario", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_PONumeroConfezioniPerImballo", "Numero confezioni", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

        .ColumnsHeader.Add "RV_POIDTipoCategoria", "RV_POIDTipoCategoria", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDCalibro", "RV_POIDCalibro", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Calibro", "Calibro", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDTipoLavorazione", "RV_POIDTipoLavorazione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, True, 1500, dgAlignleft
            
        .ColumnsHeader.Add "RV_POIDConfigurazioneEtichettaLavorazione", "RV_POIDConfigurazioneEtichettaLavorazione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "ConfigurazioneEtichettaLavorazione", "Etichette prodotto", dgchar, True, 1100, dgAlignleft

        .ColumnsHeader.Add "RV_POIDConfigurazioneEtichettaPedana", "RV_POIDConfigurazioneEtichettaPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "ConfigurazioneEtichettaPedana", "Etichette pedana", dgchar, True, 1100, dgAlignleft

        .ColumnsHeader.Add "RV_POAnnotazioniRigaOrdine", "Note riga ord.", dgchar, False, 3000, dgAlignleft
        
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
    
End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Function GET_PREZZO_IMBALLO_DA_ORDINE(IDImballo As Long, linkRiga As Long, IDOggettoOrdine As Long) As Double
On Error GoTo ERR_GET_PREZZO_IMBALLO_DA_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_PREZZO_IMBALLO_DA_ORDINE = 0

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=2 "
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
   GET_PREZZO_IMBALLO_DA_ORDINE = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PREZZO_IMBALLO_DA_ORDINE:
    GET_PREZZO_IMBALLO_DA_ORDINE = 0
End Function

Private Sub Check1_Click()
    If bLoading = False Then GET_GRIGLIA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        GrigliaCorpoOrdine_DblClick
    End If
End Sub

Private Sub Form_Load()
    bLoading = True
    CONFERMA_SEL_PREZZO_DA_ORD = 0
    
    Me.Check1.Value = Abs(NON_VIS_RIGHE_ORD_COMPLETE)
    
    GET_PARAMETRI
    
    DISEGNA_FORM
    
    CREA_RECORDSET
    bLoading = False
End Sub
Private Sub GrigliaCorpoOrdine_DblClick()
On Error GoTo ERR_GrigliaCorpoOrdine_DblClick
    
    If MODALITA_RECUPERO_RIGA_ORD = 0 Then
        frmMain.txtImportoUnitarioArticolo.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_prezzo_unitario_neutro").Value)
        frmMain.txtSconto1.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_sco_in_percentuale_1").Value)
        frmMain.txtSconto2.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_sco_in_percentuale_2").Value)
        frmMain.chkPrezzoInclusoImballo.Value = Abs(fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("RV_POImportoImballoInArticolo").Value))
        frmMain.txtRaggrOrd.Text = fnNotNull(Me.GrigliaCorpoOrdine.AllColumns("RV_PONotaRigaOrdRaggr").Value)
        
        If frmMain.CDCodiceImballo.KeyFieldID = fnNotNullN(rsGriglia!RV_POIDImballo) Then
            frmMain.txtImportoUnitarioImballo.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("ImportoUnitarioImballo").Value)
            RETURN_SEL_PREZZO_IMB_DA_ORD = 1
        End If
        
        RECUPERO_PREZZI_DA_ORD = True
        frmMain.txtIDRigaOrdine.Value = fnNotNullN(rsGriglia!IDValoriOggettoDettaglio)
        RECUPERO_PREZZI_DA_ORD = False
        
        CONFERMA_SEL_PREZZO_DA_ORD = 1
        
    End If
    
    If MODALITA_RECUPERO_RIGA_ORD = 1 Then
        frmMain.txtIDRigaOrdine.Value = fnNotNullN(rsGriglia!IDValoriOggettoDettaglio)
    End If
    
    RIPORTA_RIGA_DA_ORDINE = 1
    
    Unload Me
    
Exit Sub
ERR_GrigliaCorpoOrdine_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpoOrdine_DblClick"
    
End Sub

Private Sub GET_PARAMETRI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT VisualizzaNoteLavDaOrdineInElenco FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    VIS_NOTE_ORD_ELENCO = fnNotNullN(rs!VisualizzaNoteLavDaOrdineInElenco)
Else
    VIS_NOTE_ORD_ELENCO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub DISEGNA_FORM()
If VIS_NOTE_ORD_ELENCO = 1 Then
    Me.GrigliaCorpoOrdine.Width = 15375
    Me.Frame1.Visible = True
End If
End Sub

Private Sub GrigliaCorpoOrdine_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)

Me.txtNoteLavDaOrd.Text = fnNotNull(rsGriglia!RV_POAnnotazioniRigaLavorazione)

End Sub
Private Function GET_COLLI_LAVORATI(IDRigaOrdine As Long) As Double
On Error GoTo ERR_GET_COLLI_LAVORATI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_COLLI_LAVORATI = 0

sSQL = "SELECT SUM(Colli) AS NumeroColli "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglioRigaOrd=" & IDRigaOrdine

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_COLLI_LAVORATI = fnNotNullN(rs!NumeroColli)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_COLLI_LAVORATI:
    MsgBox Err.Description, vbCritical, "GET_COLLI_LAVORATI"
End Function
