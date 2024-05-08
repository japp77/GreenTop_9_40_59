VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRettifica 
   Caption         =   "RETTIFICHE RIGA ORDINE SELEZIONATA"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21060
   Icon            =   "frmRettifica.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   21060
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7455
      Left            =   3960
      TabIndex        =   0
      Top             =   0
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   13150
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
   Begin DmtGridCtl.DmtGrid GrigliaRettifica2 
      Height          =   7455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   13150
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
Attribute VB_Name = "frmRettifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsGrigliaRettifica As ADODB.Recordset

Private Sub Form_Load()
    GET_GRIGLIA
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * FROM RV_POIECorpoOrdineRettifica "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT
sSQL = sSQL & " ORDER BY DataModifica DESC"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    .ColumnsHeader.Add "ID", "ID", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "RV_POLinkRiga", "RV_POLinkRiga", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "NumeroRettifica", "N° rett.", dgNumeric, True, 900, dgAlignRight
    .ColumnsHeader.Add "DataModifica", "Data modifica", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "UtenteModifica", "Utente", dgchar, False, 2500, dgAlignleft
    ''PEDANA
    .ColumnsHeader.Add "RV_POIDTipoPedana", "IDTipoPedana", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Codice tipo pedana", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione tipo pedana", dgchar, False, 3500, dgAlignleft
    Set cl = .ColumnsHeader.Add("RV_POQuantitaPedana", "Q.tà pedane", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("RV_POQuantitaPedanaEffettiva", "Q.tà pedane effettive", dgDouble, False, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    ''IMBALLO
    .ColumnsHeader.Add "RV_POIDImballo", "IDImballo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, False, 3500, dgAlignleft
    .ColumnsHeader.Add "RV_POImportoImballoInArticolo", "Importo incluso imballo", dgBoolean, False, 500, dgAligncenter
    Set cl = .ColumnsHeader.Add("RV_POColliPerPedana", "Colli per pedana", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("RV_POColliSfusi", "Colli sfusi", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    .ColumnsHeader.Add "Link_Art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, False, 3500, dgAlignleft
    Set cl = .ColumnsHeader.Add("Art_numero_colli", "Totale colli", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_peso", "Peso lordo", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_tara", "Tara", dgDouble, True, 1300, dgAlignRight)
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
    Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà U.M.", dgDouble, True, 1300, dgAlignleft)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Prezzo unitario", dgDouble, True, 1300, dgAlignleft)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_1", "% Sc.1", dgDouble, True, 1300, dgAlignleft)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_2", "% Sc.2", dgDouble, True, 1300, dgAlignleft)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_pre_uni_net_sco_net_IVA", "Prezzo unitario scontato", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    .ColumnsHeader.Add "RV_POIDTipoLavorazione", "IDTipoLavorazione", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POIDTipoCategoria", "IDTipoCategoria", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POIDCalibro", "IDCalibro", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Calibro", "Calibro", dgchar, False, 2500, dgAlignleft
    'CONFEZIONI
    .ColumnsHeader.Add "RV_POIDImballoPrimario", "IDImballoPrimario", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "RV_POCodiceImballoPrimario", "Codice imballo primario", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_PODescrizioneImballoPrimario", "Descrizione imballo primario", dgchar, False, 3500, dgAlignleft
    Set cl = .ColumnsHeader.Add("RV_PONumeroConfezioniPerImballo", "Numero confez.", dgDouble, False, 1300, dgAlignRight)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
        
    'NOTE
    .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Raggr. ordine", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POAnnotazioniRigaLavorazione", "Nota riga lavorazione", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POAnnotazioniRigaOrdine", "Nota riga ordine", dgchar, True, 2500, dgAlignleft
        
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Form_Resize()
    Me.Griglia.Width = Me.Width - Me.GrigliaRettifica2.Width - 240
    Me.Griglia.Height = Me.Height - 480
    Me.GrigliaRettifica2.Height = Me.Griglia.Height
End Sub
Private Function GET_NUMERO_RETTIFICHE() As Long
On Error GoTo ERR_GET_NUMERO_RETTIFICHE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_RETTIFICHE = 0

sSQL = "SELECT COUNT(ID) as Numero "
sSQL = sSQL & "FROM RV_POCorpoOrdineRettifica "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_NUMERO_RETTIFICHE = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_NUMERO_RETTIFICHE:
    GET_NUMERO_RETTIFICHE = -1
End Function
Private Sub GET_CONFRONTO(NumeroRettifica As Long)
On Error GoTo ERR_GET_CONFRONTO
Dim NumeroTotaliRettifiche As Long
Dim NumeroRettificaControllo As Long
Dim sSQL As String
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim I As Long

If Not (rsGrigliaRettifica Is Nothing) Then
    If (rsGrigliaRettifica.State > 0) Then
        rsGrigliaRettifica.Close
    End If
    Set rsGrigliaRettifica = Nothing
End If

Set rsGrigliaRettifica = New ADODB.Recordset
rsGrigliaRettifica.CursorLocation = adUseClient
rsGrigliaRettifica.Fields.Append "NomeCampo", adVarChar, 250, adFldIsNullable

rsGrigliaRettifica.Open , , adOpenKeyset, adLockBatchOptimistic

NumeroTotaliRettifiche = GET_NUMERO_RETTIFICHE

If (NumeroTotaliRettifiche > NumeroRettifica) Then
    NumeroRettificaControllo = NumeroRettifica + 1
        
    sSQL = "SELECT * FROM RV_POCorpoOrdineRettifica "
    sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT
    sSQL = sSQL & " AND NumeroRettifica=" & NumeroRettifica
    
    Set rs1 = New ADODB.Recordset
    rs1.Open sSQL, Cn.InternalConnection
    
    sSQL = "SELECT * FROM RV_POCorpoOrdineRettifica "
    sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT
    sSQL = sSQL & " AND NumeroRettifica=" & NumeroRettificaControllo
    
    Set rs2 = New ADODB.Recordset
    rs2.Open sSQL, Cn.InternalConnection
        
    For I = 0 To rs1.Fields.Count - 1
        Select Case rs1.Fields(I).Name
            Case "ID"
            
            Case "DataModifica"
            
            Case "UtenteModifica"
            
            Case "NumeroRettifica"
            
            Case "RV_POImportoImballoInArticolo"
                If Abs(fnNotNullN(rs1.Fields(I).Value)) <> Abs(fnNotNullN(rs2.Fields(rs1.Fields(I).Name).Value)) Then
                    rsGrigliaRettifica.AddNew
                        rsGrigliaRettifica!nomecampo = rs1.Fields(I).Name
                    rsGrigliaRettifica.Update
                End If
            Case Else
                If rs1.Fields(I).Value <> rs2.Fields(rs1.Fields(I).Name).Value Then
                    rsGrigliaRettifica.AddNew
                        rsGrigliaRettifica!nomecampo = rs1.Fields(I).Name
                    rsGrigliaRettifica.Update
                End If
        End Select
    Next
    
Else
    
    sSQL = "SELECT * FROM RV_POCorpoOrdineRettifica "
    sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT
    sSQL = sSQL & " AND NumeroRettifica=" & NumeroRettifica
    
    Set rs1 = New ADODB.Recordset
    rs1.Open sSQL, Cn.InternalConnection
    
    sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
    sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND RV_POLinkRiga=" & LINK_RIGA_SELEZIONATA_RETT
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs2 = New ADODB.Recordset
    rs2.Open sSQL, Cn.InternalConnection
        
    For I = 0 To rs1.Fields.Count - 1
        Select Case rs1.Fields(I).Name
            Case "ID"
            
            Case "DataModifica"
            
            Case "UtenteModifica"
            
            Case "NumeroRettifica"
            
            Case "RV_POImportoImballoInArticolo"
                If Abs(fnNotNullN(rs1.Fields(I).Value)) <> Abs(fnNotNullN(rs2.Fields(rs1.Fields(I).Name).Value)) Then
                    rsGrigliaRettifica.AddNew
                        rsGrigliaRettifica!nomecampo = rs1.Fields(I).Name
                    rsGrigliaRettifica.Update
                End If
            Case Else
                If rs1.Fields(I).Value <> rs2.Fields(rs1.Fields(I).Name).Value Then
                    rsGrigliaRettifica.AddNew
                        rsGrigliaRettifica!nomecampo = rs1.Fields(I).Name
                    rsGrigliaRettifica.Update
                End If
        End Select
    Next
    
End If

GET_GRIGLIA_CONFRONTO

Exit Sub
ERR_GET_CONFRONTO:
    MsgBox Err.Description, vbCritical, "GET_CONFRONTO"
    GET_GRIGLIA_CONFRONTO
End Sub


Private Sub GET_GRIGLIA_CONFRONTO()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

With Me.GrigliaRettifica2
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    .ColumnsHeader.Add "NomeCampo", "Campo", dgchar, True, 2500, dgAlignleft
        
    Set .Recordset = rsGrigliaRettifica
    .LoadUserSettings
    .Refresh
    
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_CONFRONTO"
End Sub

Private Sub Griglia_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    GET_CONFRONTO AllColumns("NumeroRettifica").Value
End Sub
