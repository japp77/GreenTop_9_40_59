VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.1#0"; "DmtGridCtl.ocx"
Begin VB.Form frmChiusuraConferimento 
   Caption         =   "Gestione chiusura conferimenti"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "Conferma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9551
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      UpdatePosition  =   0   'False
      UseUserSettings =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmChiusuraConferimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Sub cmdConferma_Click()
Dim I As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

    rsGriglia.UpdateBatch
    
    For I = 0 To 1000000
    
    Next
    
sSQL = "SELECT * FROM RV_POTMPGestioneChiusuraConf WHERE IDOggetto=" & Link_Oggetto
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "UPDATE RV_POCaricoMerceRighe SET "
    sSQL = sSQL & "Chiuso=" & fnNormBoolean(rs!Chiuso)
    sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe= " & fnNotNullN(rs!IDRV_POcaricoMerceRighe)
    
    Cn.Execute sSQL
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
GET_RIGHE_DOCUMENTO Link_Oggetto, 2
fncGriglia
End Sub
Public Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    

    sSQL = "SELECT * FROM RV_POTMPGestioneChiusuraConf WHERE IDOggetto=" & Link_Oggetto
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            'Set rsEvent = rsGriglia2.Data
    
        
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            
                    Set cl = .ColumnsHeader.Add("Chiuso", "Chiuso", dgBoolean, True, 700, dgAlignleft)
                        cl.Editable = True
                    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDConferimentoRiga", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "DataConferimento", "Data Conf.", dgDate, True, 1100, dgAlignleft
                    .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, True, 1000, dgAlignleft
                    .ColumnsHeader.Add "CodiceArticoloConferito", "Codice Articolo", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "ArticoloConferito", "Articolo", dgchar, False, 1500, dgAlignleft
                    Set cl = .ColumnsHeader.Add("QtaConferita", "Q.tà Conf.", dgDouble, True, 1200, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QtaQuadrata", "Q.tà Quad.", dgDouble, True, 1200, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QtaVenduta", "Q.tà Vend.", dgDouble, True, 1200, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QtaDifferenza", "Differenza", dgDouble, True, 1200, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    
                    .ColumnsHeader.Add "Socio", "Socio", dgchar, True, 2500, dgAlignleft
                    .ColumnsHeader.Add "NomeSocio", "Nome", dgchar, False, 1500, dgAlignleft
                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
End Sub
Private Sub GET_RIGHE_DOCUMENTO(IDOggetto As Long, IDTipoOggetto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QtaVenduta As Double
Dim QtaQuadrata As Double
Dim QtaDifferenza As Double

Cn.Execute "DELETE RV_POTMPGestioneChiusuraConf WHERE IDOggetto=" & IDOggetto

Select Case IDTipoOggetto
    Case 2 'Documento di trasporto
        sSQL = "SELECT RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM, ValoriOggettoDettaglio0004.IDOggetto, ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PODataConferimento, ValoriOggettoDettaglio0004.RV_POSocio, ValoriOggettoDettaglio0004.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento , RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo "
        sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 ON "
        sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe = ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0004.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1 "
        sSQL = sSQL & " GROUP BY ValoriOggettoDettaglio0004.IDOggetto, ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PODataConferimento, ValoriOggettoDettaglio0004.RV_POSocio, ValoriOggettoDettaglio0004.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, "
        sSQL = sSQL & "RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM "
    Case 114 'Fattura accompaganatoria
        sSQL = "SELECT RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM, ValoriOggettoDettaglio0001.IDOggetto, ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_PODataConferimento, ValoriOggettoDettaglio0001.RV_POSocio, ValoriOggettoDettaglio0001.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento , RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo "
        sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0001 ON "
        sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe = ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0001.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1 "
        sSQL = sSQL & " GROUP BY ValoriOggettoDettaglio0001.IDOggetto, ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_PODataConferimento, ValoriOggettoDettaglio0001.RV_POSocio, ValoriOggettoDettaglio0001.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, "
        sSQL = sSQL & "RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM "
   Case 8 'Scontrino non fiscale
        sSQL = "SELECT RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM, ValoriOggettoDettaglio0034.IDOggetto, ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_PODataConferimento, ValoriOggettoDettaglio0034.RV_POSocio, ValoriOggettoDettaglio0034.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento , RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo "
        sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0034 ON "
        sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe = ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0034.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1 "
        sSQL = sSQL & " GROUP BY ValoriOggettoDettaglio0034.IDOggetto, ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe, "
        sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_PODataConferimento, ValoriOggettoDettaglio0034.RV_POSocio, ValoriOggettoDettaglio0034.RV_PONomeSocio, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, "
        sSQL = sSQL & "RV_POCaricoMerceRighe.Chiuso, RV_POCaricoMerceRighe.Qta_UM "
    
End Select

Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF
    If fnNotNullN(rs!RV_POIDConferimentoRighe) > 0 Then
        
        QtaQuadrata = GET_RIEPILOGO_QUANTITA_LAVORAZIONE(fnNotNullN(rs!RV_POIDConferimentoRighe))
        QtaVenduta = GET_RIEPILOGO_QUANTITA_VENDUTO(fnNotNullN(rs!RV_POIDConferimentoRighe))
        QtaDifferenza = fnNotNullN(rs!Qta_UM) - (QtaQuadrata + QtaVenduta)
        
        sSQL = "INSERT INTO RV_POTMPGestioneChiusuraConf ("
        sSQL = sSQL & "IDOggetto, IDRV_POCaricoMerceRighe, CodiceArticoloConferito, ArticoloConferito, "
        sSQL = sSQL & "NumeroConferimento, DataConferimento, Chiuso, Socio, NomeSocio, "
        sSQL = sSQL & "QtaConferita, QtaQuadrata, QtaVenduta, QtaDifferenza) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDOggetto & ", "
        sSQL = sSQL & fnNotNullN(rs!RV_POIDConferimentoRighe) & ", "
        sSQL = sSQL & fnNormString(rs!CodiceArticolo) & ", "
        sSQL = sSQL & fnNormString(rs!Articolo) & ", "
        sSQL = sSQL & fnNormNumber(rs!NumeroDocumento) & ", "
        sSQL = sSQL & fnNormDate(rs!RV_PODataConferimento) & ", "
        sSQL = sSQL & fnNormBoolean(rs!Chiuso) & ", "
        sSQL = sSQL & fnNormString(rs!RV_POSocio) & ", "
        sSQL = sSQL & fnNormString(rs!RV_PONomeSocio) & ", "
        sSQL = sSQL & fnNormNumber(rs!Qta_UM) & ", "
        sSQL = sSQL & fnNormNumber(QtaQuadrata) & ", "
        sSQL = sSQL & fnNormNumber(QtaVenduta) & ", "
        sSQL = sSQL & fnNormNumber(QtaDifferenza) & ")"
        
        Cn.Execute sSQL
    End If
    
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_RIEPILOGO_QUANTITA_VENDUTO(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_RIEPILOGO_QUANTITA_VENDUTO = 0

'DOCUMENTO DI TRASPORTO
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'CORRISPETTIVI
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_RIEPILOGO_QUANTITA_LAVORAZIONE(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



sSQL = "SELECT Sum(Qta_UM) as QtaTotale "
sSQL = sSQL & "FROM RV_POLavorazione "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = 0
Else
    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing



End Function
Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If Griglia.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If Griglia.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("Chiuso").Value), 2
            End If
        End If
    End If
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Chiuso").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
        rsGriglia.Fields("Chiuso").Value = Abs(CLng(Selected))
        'sbCheckSelected
        Me.Griglia.Refresh
    End If
End Sub
