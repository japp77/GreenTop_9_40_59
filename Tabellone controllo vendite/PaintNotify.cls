VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaintNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements dmtgridctl.IPaintNotify

Private Sub IPaintNotify_PaintValue(ByVal ColumnsValue As dmtgridctl.dgColumns, ByVal CurrentColumn As dmtgridctl.dgColumnHeader, ByVal Row As Integer, Value As Variant)
On Error Resume Next
    
    Select Case CurrentColumn.FieldName
        Case "RiferimentoDocumento"
            Value = GET_STRINGA_DOCUMENTO(fnNotNullN(ColumnsValue("IDOggetto").Value), fnNotNull(ColumnsValue("RV_POCodiceLottoVendita").Value), fnNotNullN(ColumnsValue("IDTipoOggetto").Value))
        Case "QuantitaDiRiferimento"
            Value = GET_QUANTITA_DOCUMENTO(fnNotNullN(ColumnsValue("IDOggetto").Value), fnNotNull(ColumnsValue("RV_POCodiceLottoVendita").Value), fnNotNullN(ColumnsValue("IDTipoOggetto").Value))
        Case "ImportoDiRiferimento"
            Value = GET_IMPORTO_DOCUMENTO(fnNotNullN(ColumnsValue("IDOggetto").Value), fnNotNull(ColumnsValue("RV_POCodiceLottoVendita").Value), fnNotNullN(ColumnsValue("IDTipoOggetto").Value))
    End Select

End Sub
Private Function GET_STRINGA_DOCUMENTO(IDOggettoCollegato As Long, CodiceLotto As String, IDTipoOggettoCollegato As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'NOTA DI CREDITO
sSQL = "SELECT ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000B.IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDOggetto=" & IDOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDTipoOggetto=" & IDTipoOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POCodiceLotto=" & fnNormString(CodiceLotto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_STRINGA_DOCUMENTO = ""
Else
    GET_STRINGA_DOCUMENTO = "Nota di credito numero " & fnNotNull(rs!doc_numero) & " del " & fnNotNull(rs!doc_data)
End If
rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_QUANTITA_DOCUMENTO(IDOggettoCollegato As Long, CodiceLotto As String, IDTipoOggettoCollegato As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'NOTA DI CREDITO
sSQL = "SELECT ValoriOggettoDettaglio0016.Art_quantita_totale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000B.IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDOggetto=" & IDOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDTipoOggetto=" & IDTipoOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POCodiceLotto=" & fnNormString(CodiceLotto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_QUANTITA_DOCUMENTO = 0
Else
    GET_QUANTITA_DOCUMENTO = fnNotNullN(rs!art_quantita_totale)
End If
rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_IMPORTO_DOCUMENTO(IDOggettoCollegato As Long, CodiceLotto As String, IDTipoOggettoCollegato As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'NOTA DI CREDITO
sSQL = "SELECT ValoriOggettoDettaglio0016.Art_prezzo_unitario_netto_IVA "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000B.IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDOggetto=" & IDOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDTipoOggetto=" & IDTipoOggettoCollegato
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POCodiceLotto=" & fnNormString(CodiceLotto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_DOCUMENTO = 0
Else
    GET_IMPORTO_DOCUMENTO = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
End If
rs.CloseResultset
Set rs = Nothing

End Function
