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
'/****************************************************************************
' * PROJECT     : Programma dimostrativo DemoDmtGrid (Test.vbp)
' *
' * NAME        : PaintNotify.cls
' *
' * DESCRIPTION : Programma dimostrativo sulle funzionalità del controllo DmtGridCtl.OCX
' *
' *
' * AUTHOR      : Diamante spa
' *
' * DATE        : 04/04/2000
' *
' * MODIFIED    : 25/03/2008
' *****************************************************************************/

Option Explicit

Implements DmtGridCtl.IPaintNotify

Private Sub IPaintNotify_PaintValue(ByVal ColumnsValue As DmtGridCtl.dgColumns, ByVal CurrentColumn As DmtGridCtl.dgColumnHeader, ByVal Row As Integer, Value As Variant)
    
    Select Case CurrentColumn.FieldName
    
        Case "TipoPedana"
            DoEvents
            Screen.MousePointer = 11
            Value = GET_TIPO_PEDANA(ColumnsValue("IDRV_POPedana").Value)
            DoEvents
            Screen.MousePointer = 0
        Case "StatoOrdine"
            Value = GET_STATO_ORDINE(fnNotNullN(ColumnsValue("IDOggetto").Value))
        Case "Serre"
            Value = GET_SERRE(fnNotNullN(ColumnsValue("IDRV_PO01_LottoCampagna").Value))
            
    End Select

End Sub
Private Function GET_TIPO_PEDANA(IDPedana As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POTipoPedana.TipoPedana "
sSQL = sSQL & "FROM RV_POPedana LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoPedana ON RV_POPedana.IDRV_POTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & "WHERE RV_POPedana.IDRV_POPedana=" & IDPedana

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PEDANA = ""
Else
    GET_TIPO_PEDANA = fnNotNull(rs!TipoPedana)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_STATO_ORDINE(IDOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_STATO_ORDINE = "APERTO"
Else
    GET_STATO_ORDINE = "CONFERMATO"
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SERRE(IDLotto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT Codice "
sSQL = sSQL & "FROM RV_PO01_SerraPerLotto LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_Serra ON RV_PO01_SerraPerLotto.IDRV_PO01_Serra = RV_PO01_Serra.IDRV_PO01_Serra "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLotto
    
Set rs = Cn.OpenResultset(sSQL)
    
    GET_SERRE = ""
    
    While Not rs.EOF
        GET_SERRE = GET_SERRE & Trim(fnNotNull(rs!Codice)) & "; "
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function
