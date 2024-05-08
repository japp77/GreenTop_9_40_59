VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmInserisciPedane 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INSERISCI PEDANE"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInserisciPedane.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   19110
   ShowInTaskbar   =   0   'False
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
      Left            =   16200
      TabIndex        =   1
      Top             =   4080
      Width           =   2895
   End
   Begin DmtGridCtl.DmtGrid GridPedana 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   7011
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
Attribute VB_Name = "frmInserisciPedane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
If ((rsGriglia.EOF) And (rsGriglia.BOF)) Then Exit Sub

rsGriglia.MoveFirst

While Not rsGriglia.EOF
    rsInserisciPedane.AddNew
        For I = 0 To rsGriglia.Fields.Count - 1
            rsInserisciPedane.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
        Next
    rsInserisciPedane.Update
rsGriglia.MoveNext
Wend

Unload Me

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Private Sub Form_Load()
    
    CREA_RECORDSET
    
    GET_GRIGLIA
    
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
With Me.GridPedana
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDTipoPedana", "IDTipoPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodicePedana", "Tipo pedana", dgchar, True, 2000, dgAlignleft
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 2000, dgAlignleft
        .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("QuantitaPedane", "Q.tà pedane", dgDouble, True, 2300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("NumeroPallet", "N° pallet per ogni pedana", dgDouble, True, 2300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("QuantitaEffettiva", "Q.tà effettiva", dgDouble, True, 2300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoNetto", "Peso", dgDouble, True, 2300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoTotale", "Peso totale", dgDouble, True, 2300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
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

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim rs As DmtOleDbLib.adoResultset


If Not (rsInserisciPedane Is Nothing) Then
    Set rsInserisciPedane = Nothing
End If

Set rsInserisciPedane = New ADODB.Recordset
rsInserisciPedane.CursorLocation = adUseClient

rsInserisciPedane.Fields.Append "IDTipoPedana", adInteger, , adFldIsNullable
rsInserisciPedane.Fields.Append "CodicePedana", adVarChar, 250, adFldIsNullable
rsInserisciPedane.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsInserisciPedane.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsInserisciPedane.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsInserisciPedane.Fields.Append "PesoNetto", adDouble, , adFldIsNullable
rsInserisciPedane.Fields.Append "QuantitaPedane", adDouble, , adFldIsNullable
rsInserisciPedane.Fields.Append "NumeroPallet", adDouble, , adFldIsNullable
rsInserisciPedane.Fields.Append "QuantitaEffettiva", adDouble, , adFldIsNullable
rsInserisciPedane.Fields.Append "PesoTotale", adDouble, , adFldIsNullable


rsInserisciPedane.Open , , adOpenKeyset, adLockBatchOptimistic




If Not (rsGriglia Is Nothing) Then
    Set rsGriglia = Nothing
End If

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "IDTipoPedana", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "CodicePedana", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "PesoNetto", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "QuantitaPedane", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "NumeroPallet", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "QuantitaEffettiva", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "PesoTotale", adDouble, , adFldIsNullable


rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT " & sTabellaDettaglio & ".RV_POIDTipoPedana, "
sSQL = sSQL & "Articolo.IDArticolo, Articolo.CodiceArticolo, Articolo.Articolo, Articolo.PesoNetto, "
sSQL = sSQL & "RV_POTipoPedana.NumeroPallet,  RV_POTipoPedana.CodiceTipoPedana "
sSQL = sSQL & "FROM " & sTabellaDettaglio & " INNER JOIN "
sSQL = sSQL & "RV_POTipoPedana ON " & sTabellaDettaglio & ".RV_POIDTipoPedana = dbo.RV_POTipoPedana.IDRV_POTipoPedana INNER JOIN "
sSQL = sSQL & "Articolo ON dbo.RV_POTipoPedana.IDArticoloImballo = dbo.Articolo.IDArticolo "
sSQL = sSQL & "WHERE " & sTabellaDettaglio & ".IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND " & sTabellaDettaglio & ".RV_POTipoRiga=1 "
sSQL = sSQL & "GROUP BY " & sTabellaDettaglio & ".RV_POIDTipoPedana, "
sSQL = sSQL & "Articolo.IDArticolo, Articolo.CodiceArticolo, Articolo.Articolo, Articolo.PesoNetto, "
sSQL = sSQL & "RV_POTipoPedana.NumeroPallet,  RV_POTipoPedana.CodiceTipoPedana "


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsGriglia.AddNew
        rsGriglia!IDTipoPedana = fnNotNullN(rs!RV_POIDTipoPedana)
        rsGriglia!CodicePedana = fnNotNull(rs!CodiceTipoPedana)
        rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsGriglia!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
        rsGriglia!Articolo = fnNotNull(rs!Articolo)
        rsGriglia!PesoNetto = fnNotNullN(rs!PesoNetto)
        rsGriglia!QuantitaPedane = GET_QUANTITA_PEDANE(oDoc.IDOggetto, fnNotNullN(rsGriglia!IDTipoPedana)) 'fnNotNullN(rs!Quantita)
        rsGriglia!NumeroPallet = IIf(fnNotNullN(rs!NumeroPallet) = 0, 1, fnNotNullN(rs!NumeroPallet))
        rsGriglia!QuantitaEffettiva = rsGriglia!QuantitaPedane * rsGriglia!NumeroPallet
        rsGriglia!PesoTotale = rsGriglia!QuantitaEffettiva * rsGriglia!PesoNetto
    rsGriglia.Update
    
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub

Private Sub GridPedana_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
On Error GoTo ERR_GridPedana_AfterChangeFieldValue
    rsGriglia!QuantitaEffettiva = fnNotNullN(rsGriglia!QuantitaPedane) * fnNotNullN(rsGriglia!NumeroPallet)
    rsGriglia!PesoTotale = fnNotNullN(rsGriglia!QuantitaEffettiva) * fnNotNullN(rsGriglia!PesoNetto)
    
    Me.GridPedana.Refresh

Exit Sub
ERR_GridPedana_AfterChangeFieldValue:
    MsgBox Err.Description, vbCritical, "GridPedana_AfterChangeFieldValue"
End Sub

Private Function GET_QUANTITA_PEDANE(IDOggetto As Long, IDTipoPedana As Long) As Long
On Error GoTo ERR_GET_QUANTITA_PEDANE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPed As ADODB.Recordset

GET_QUANTITA_PEDANE = 0

Set rsPed = New ADODB.Recordset
rsPed.Fields.Append "IDPedana", adInteger, , adFldIsNullable
rsPed.CursorLocation = adUseClient
rsPed.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT IDValoriOggettoDettaglio, RV_POIDPedana "
sSQL = sSQL & "FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsPed.Filter = "IDPedana=" & fnNotNullN(rs!RV_POIDPedana)
    
    If rsPed.EOF Then
        rsPed.AddNew
            rsPed!IDPedana = fnNotNullN(rs!RV_POIDPedana)
        rsPed.Update
        GET_QUANTITA_PEDANE = GET_QUANTITA_PEDANE + 1

    End If
    
    rsPed.Filter = vbNullString
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_QUANTITA_PEDANE:
    MsgBox Err.Description, vbCritical, "GET_QUANTITA_PEDANE"
End Function

