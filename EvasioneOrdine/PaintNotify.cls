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

Implements DmtGridCtl.IPaintNotify

Private Sub IPaintNotify_PaintValue(ByVal ColumnsValue As DmtGridCtl.dgColumns, ByVal CurrentColumn As DmtGridCtl.dgColumnHeader, ByVal Row As Integer, Value As Variant)
On Error Resume Next
    
    Select Case CurrentColumn.FieldName
        Case "StatoOrdine"
            Value = GET_STATO_ORDINE(fnNotNullN(ColumnsValue("IDOggetto").Value))
        Case "TipoLavorazione"
            Value = GET_VALORE_CAMPO("RV_POTipoLavorazione", "TipoLavorazione", "IDRV_POTipoLavorazione", fnNotNullN(ColumnsValue("IDTipoLavorazione").Value))
        Case "TipoCategoria"
            Value = GET_VALORE_CAMPO("RV_POTipoCategoria", "TipoCategoria", "IDRV_POTipoCategoria", fnNotNullN(ColumnsValue("IDRV_POTipoCategoria").Value))
        Case "Calibro"
            Value = GET_VALORE_CAMPO("RV_POCalibro", "Calibro", "IDRV_POCalibro", fnNotNullN(ColumnsValue("IDRV_POCalibro").Value))
        
    End Select

End Sub
Private Function GET_STATO_ORDINE(IDOggetto As Long) As String
On Error GoTo ERR_GET_STATO_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_STATO_ORDINE = "APERTO"
Else
    GET_STATO_ORDINE = "CONFERMATO"
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_STATO_ORDINE:
    GET_STATO_ORDINE = ""
End Function
Private Function GET_VALORE_CAMPO(Tabella As String, CampoValoreRitorno As String, CampoWhere As String, ValoreCampoWhere As Long) As String
On Error GoTo ERR_GET_VALORE_CAMPO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & CampoValoreRitorno
sSQL = sSQL & " FROM " & Tabella
sSQL = sSQL & " WHERE " & CampoWhere & "=" & ValoreCampoWhere

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALORE_CAMPO = ""
Else
    GET_VALORE_CAMPO = fnNotNull(rs.adoColumns(CampoValoreRitorno).Value)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_VALORE_CAMPO:
    GET_VALORE_CAMPO = ""
End Function
