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
    
            
        Case "ColliImpeganti"
            Value = GET_Colli_Impegnati(fnNotNullN(ColumnsValue("IDCodiceLotto_Vendita").Value), ColumnsValue("Colli").Value, fnNotNullN(ColumnsValue("QtaColliVenduti").Value))
        Case "ColliDisponibili"
            Value = GET_Colli_Disponibili(fnNotNullN(ColumnsValue("IDCodiceLotto_Vendita").Value), fnNotNullN(ColumnsValue("Colli").Value), fnNotNullN(ColumnsValue("QtaColliVenduti").Value))
        Case "DisponibilitaLotto"
            Value = DispLotto(fnNotNullN(ColumnsValue("IDCodiceLotto_Vendita").Value))
    End Select




End Sub
Private Function GET_Colli_Impegnati(IDLotto As Long, ColliCaricati As Double, ColliVenduti As Double) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT sum(QtaColliDoc) as QtaColliDoc, sum(QtaColliPrec) as QtaColliPrec FROM RV_POTMPQtaLottoDocumento WHERE IDLottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_Colli_Impegnati = 0
Else
    If fnNotNullN(rs!QtaColliPrec) = 0 Then
        GET_Colli_Impegnati = fnNotNullN(rs!QtaColliDoc)
    Else
        If (fnNotNullN(rs!QtaColliDoc) - fnNotNullN(rs!QtaColliPrec)) >= 0 Then
            GET_Colli_Impegnati = fnNotNullN(rs!QtaColliDoc) - fnNotNullN(rs!QtaColliPrec)
        Else
            GET_Colli_Impegnati = 0
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_Colli_Disponibili(IDLotto As Long, ColliCaricati As Double, ColliVenduti As Double) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ColliImpegnati As Double

sSQL = "SELECT sum(QtaColliDoc) as QtaColliDoc, sum(QtaColliPrec) as QtaColliPrec FROM RV_POTMPQtaLottoDocumento WHERE IDLottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_Colli_Disponibili = 0
Else
    If fnNotNullN(rs!QtaColliPrec) = 0 Then
        GET_Colli_Disponibili = ColliCaricati - ColliVenduti - fnNotNullN(rs!QtaColliDoc)
    Else
        If (fnNotNullN(rs!QtaColliDoc) - fnNotNullN(rs!QtaColliPrec)) >= 0 Then
            GET_Colli_Disponibili = ColliCaricati - ColliVenduti - (fnNotNullN(rs!QtaColliDoc) - fnNotNullN(rs!QtaColliPrec))
        Else
            GET_Colli_Disponibili = ColliCaricati - ColliVenduti + (fnNotNullN(rs!QtaColliPrec) - fnNotNullN(rs!QtaColliDoc))
        End If
    End If
End If
    

rs.CloseResultset
Set rs = Nothing



End Function

Private Function DispLotto(IDLotto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Giacenza FROM LottoArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDLottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    DispLotto = 0
Else
    If IsNull(rs!Giacenza) Then
        DispLotto = 0
    Else
        DispLotto = fnNotNullN(rs!Giacenza)
    End If
    
End If
End Function
