VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAnalisiLiq 
   Caption         =   "Analisi liquidazione per documento"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalisiLiq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   15055
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAnalisiLiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ELABORAZIONE_DATI()
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim NumeroColonna As Long
Dim NumeroColonnaAfterComm As Long
Dim NumeroColonnaAfterCommCol As Long
Dim riga As Long
Dim NumeroColonnaComm As Long

Dim I As Long

With Me.MSFlexGrid1
    .Top = 0
    .Left = 0
    .Cols = 12 + GET_NUMERO_COMMISSIONI
    .Rows = 1
    NumeroColonna = 0
    
    'PREPARAZIONE COLONNE
    .TextMatrix(0, 0) = "Articolo"
    .TextMatrix(0, 1) = "Unità di misura vendita"
    .TextMatrix(0, 2) = "Quantità venduta"
    .TextMatrix(0, 3) = "Unità di misura Liquidazione"
    .TextMatrix(0, 4) = "Quantità liquidazione"
    .TextMatrix(0, 5) = "Importo di vendita"
    .TextMatrix(0, 6) = "Importo merce"
    .TextMatrix(0, 7) = "Importo imballo"
    .TextMatrix(0, 8) = "Variazione manuale"
    
    NumeroColonna = 9
    
    GET_COMMISSIONI NumeroColonna
        
    NumeroColonnaAfterComm = NumeroColonna
    NumeroColonnaAfterCommCol = NumeroColonnaAfterComm
    
    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Importo di liquidazione"
    NumeroColonnaAfterCommCol = NumeroColonnaAfterCommCol + 1
    
    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Pedana"
    NumeroColonnaAfterCommCol = NumeroColonnaAfterCommCol + 1
    
    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Tipo pedana"
'    NumeroColonnaAfterCommCol = NumeroColonnaAfterCommCol + 1
'
'    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Volume unitario"
'    NumeroColonnaAfterCommCol = NumeroColonnaAfterCommCol + 1
'
'    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Vol. Tot. merce"
'    NumeroColonnaAfterCommCol = NumeroColonnaAfterCommCol + 1
'
'    .TextMatrix(0, NumeroColonnaAfterCommCol) = "Vol. Tot. pedana"

    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignCenterCenter
        .ColWidth(I) = 1500
    Next
    
    .ColWidth(0) = 4000
    .ColWidth(.Cols - 2) = 2500
    .ColWidth(.Cols - 1) = 3000
    

    .RowHeight(0) = 600
    
    riga = 1
    
    
    sSQL = "SELECT IDValoriOggettoDettaglio, IDArticoloVenduto, CodiceArticoloVenduto, DescrizioneArticoloVenduto, IDUnitaDiMisuraVenduto, UnitaDiMisuraVenduto, "
    sSQL = sSQL & "PrezzoScontato, RV_POQuantitaLiquidazione, RV_POImportoInclusoImballo, RV_POIDImballo, CodiceArticoloImballo, "
    sSQL = sSQL & "RV_POImportoUnitarioImballo, RV_POPrezzoMerceNetta, RV_POVariazionePrezzoImballo, RV_POImportoRigaCommissioni, "
    sSQL = sSQL & "RV_POIDPedana, RV_POCodicePedana, RV_POIDTipoPedana, CodiceTipoPedana, TipoPedana, RV_POImportoLiquidazione, QuantitaTotale, DescrizioneUnitaDiMisuraLiquidazione "
    sSQL = sSQL & "FROM RV_POIEControlloVendite "
    sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection

    While Not rs.EOF
        .Rows = riga + 1
        
        .Row = riga
        
        .RowHeight(.Row) = 500
        
        .RowData(riga) = fnNotNullN(rs!IDValoriOggettoDettaglio)
        
        .Col = 0
        .TextMatrix(riga, .Col) = fnNotNull(rs!DescrizioneArticoloVenduto)
        .CellAlignment = flexAlignLeftCenter
        
        .Col = 1
        .TextMatrix(riga, .Col) = fnNotNull(rs!UnitaDiMisuraVenduto)
        .CellAlignment = flexAlignLeftCenter
        
        .Col = 2
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!QuantitaTotale), 2)
        .CellAlignment = flexAlignRightCenter
        
        
        .Col = 3
        .TextMatrix(riga, .Col) = Trim(fnNotNull(rs!DescrizioneUnitaDiMisuraLiquidazione))
        If .TextMatrix(riga, .Col) = "" Then .TextMatrix(riga, .Col) = fnNotNull(rs!UnitaDiMisuraVenduto)
        .CellAlignment = flexAlignLeftCenter

        .Col = 4
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!RV_POQuantitaLiquidazione), 2)
        .CellAlignment = flexAlignRightCenter
        
        .Col = 5
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!PrezzoScontato) * fnNotNullN(rs!QuantitaTotale), 2)
        .CellAlignment = flexAlignRightCenter
        
        .Col = 6
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!RV_POPrezzoMerceNetta) * fnNotNullN(rs!QuantitaTotale), 5)
        .CellAlignment = flexAlignRightCenter
        
        .Col = 7
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!RV_POVariazionePrezzoImballo) * fnNotNullN(rs!QuantitaTotale), 5)
        .CellAlignment = flexAlignRightCenter
        
        .Col = 8
        .TextMatrix(riga, 8) = 0
        .CellAlignment = flexAlignRightCenter
        
        If (NumeroColonnaAfterComm > 9) Then
            For NumeroColonnaComm = 9 To NumeroColonnaAfterComm - 1
                GET_COMMISSIONE_PER_RIGA fnNotNullN(rs!IDValoriOggettoDettaglio), Me.MSFlexGrid1.ColData(NumeroColonnaComm), NumeroColonnaComm, riga, fnNotNullN(rs!RV_POQuantitaLiquidazione)
            Next
        End If
        
        .Col = NumeroColonnaAfterComm
        .TextMatrix(riga, .Col) = FormatNumber(fnNotNullN(rs!RV_POImportoLiquidazione) * fnNotNullN(rs!RV_POQuantitaLiquidazione), 5)
        .CellAlignment = flexAlignRightCenter
        
        NumeroColonnaAfterCommCol = NumeroColonnaAfterComm + 1
        
        .Col = NumeroColonnaAfterCommCol
        .TextMatrix(riga, .Col) = fnNotNull(rs!RV_POCodicePedana)
        .CellAlignment = flexAlignLeftCenter
        
        .Col = NumeroColonnaAfterCommCol + 1
        .TextMatrix(riga, .Col) = fnNotNull(rs!TipoPedana)
        .CellAlignment = flexAlignLeftCenter
        riga = riga + 1
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    'TOTALI
    .Rows = riga + 1
    
    .Row = riga
    .RowHeight(.Row) = 400
    .Col = 0
    .TextMatrix(riga, .Col) = "TOTALI"
    .CellAlignment = flexAlignLeftCenter
    .CellFontBold = True
    
    .Col = 4
    .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .CellBackColor = vbYellow
    
    .Col = 5
    .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .CellBackColor = vbYellow
    
    .Col = 6
    .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .CellBackColor = vbYellow
    
    .Col = 7
    .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .CellBackColor = vbYellow
    .CellFontSize = 10
    If (NumeroColonnaAfterComm > 9) Then
        For NumeroColonnaComm = 9 To NumeroColonnaAfterComm - 1
            .Col = NumeroColonnaComm
            .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
            .CellAlignment = flexAlignRightCenter
            .CellFontBold = True
            .CellBackColor = vbYellow
        Next
    End If
                
    

    .Col = NumeroColonnaAfterComm
    .TextMatrix(riga, .Col) = FormatNumber(GET_TOTALI_COLONNA(.Col), 2)
    .CellAlignment = flexAlignRightCenter
    .CellFontBold = True
    .CellBackColor = vbYellow
    
End With


End Sub

Private Sub Form_Load()
    ELABORAZIONE_DATI
End Sub

Private Sub Form_Resize()
    
    If ((Me.Width < 14000) Or (Me.Height < 9300)) Then Exit Sub

    Me.MSFlexGrid1.Width = Me.Width - 350
    Me.MSFlexGrid1.Height = Me.Height - 600

End Sub
Private Function GET_NUMERO_COMMISSIONI() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(IDRV_POCommissioniPerDoc) AS NumeroColonne "
sSQL = sSQL & "FROM RV_POIECommissioniPerDocRigheRaggr "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_COMMISSIONI = 0
Else
    GET_NUMERO_COMMISSIONI = fnNotNullN(rs!NumeroColonne)
End If

rs.CloseResultset
Set rs = Nothing
    
    
End Function
Private Function GET_COMMISSIONI(NumeroColonna As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroColonnaComm As Long
Dim TestoColonna As String




sSQL = "SELECT *  "
sSQL = sSQL & "FROM RV_POIECommissioniPerDocRigheRaggr "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
        
    With MSFlexGrid1
        TestoColonna = ""
        .ColData(NumeroColonna) = fnNotNullN(rs!IDRV_POCommissioniPerDoc)
        TestoColonna = fnNotNull(rs!TipoCommissione)
        If (fnNotNullN(rs!IDRV_POTipoPedana) > 0) Then
            TestoColonna = TestoColonna & vbCrLf & "(" & fnNotNull(rs!TipoPedanaComm) & ")"
        End If
        .TextMatrix(0, NumeroColonna) = TestoColonna

        
        
    End With
    
    NumeroColonna = NumeroColonna + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
    
    
End Function
Private Sub GET_COMMISSIONE_PER_RIGA(IDRiga As Long, IDTipoCommissione As Long, NumeroColonna As Long, NumeroRiga As Long, QuantitaLiquidazione As Double)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCommissioniPerDocRighe "
sSQL = sSQL & "WHERE IDRV_POCommissioniPerDoc = " & IDTipoCommissione
sSQL = sSQL & " AND IDValoriOggettoDettaglio = " & IDRiga
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

Me.MSFlexGrid1.Col = NumeroColonna
'Me.MSFlexGrid1.Row = NumeroRiga

If rs.EOF Then
    Me.MSFlexGrid1.TextMatrix(NumeroRiga, NumeroColonna) = FormatNumber(0, 2)
    Me.MSFlexGrid1.CellAlignment = flexAlignRightCenter
    Me.MSFlexGrid1.CellBackColor = vbGreen
Else

    Me.MSFlexGrid1.TextMatrix(NumeroRiga, NumeroColonna) = FormatNumber(fnNotNullN(rs!Importo) * QuantitaLiquidazione, 5)
    Me.MSFlexGrid1.CellAlignment = flexAlignRightCenter
    Me.MSFlexGrid1.CellBackColor = vbGreen
End If



rs.CloseResultset
Set rs = Nothing

End Sub

Private Function GET_TOTALI_COLONNA(NumeroColonna) As Double
Dim I As Long
Dim TotaleRiga As Double

TotaleRiga = 0
For I = 1 To Me.MSFlexGrid1.Rows - 1
    TotaleRiga = TotaleRiga + fnNotNullN(MSFlexGrid1.TextMatrix(I, NumeroColonna))
Next

GET_TOTALI_COLONNA = TotaleRiga

End Function
Private Function GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = ""

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
        
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
        Case 1
            GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = "Colli"
        Case 2
            GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = "Peso lordo"
        Case 3
            GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = "Peso netto"
        Case 4
            GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = "Tara"
        Case 5
            GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = "Pezzi"
    End Select
End If

rs.CloseResultset
Set rs = Nothing

End Function

