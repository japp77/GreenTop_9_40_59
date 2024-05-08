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
        Case "Serre"
            Value = GET_SERRE(fnNotNullN(ColumnsValue("IDRV_PO01_LottoCampagna").Value))
        Case "Settore"
            Value = GET_SETTORE(fnNotNullN(ColumnsValue("IDRV_PO01_Serra").Value), fnNotNullN(ColumnsValue("IDRV_PO01_Schema").Value))
        Case "Varieta"
            Value = GET_VARIETA(fnNotNullN(ColumnsValue("IDRV_PO01_Varieta").Value))
        Case "Famiglia"
            Value = GET_FAMIGLIA(fnNotNullN(ColumnsValue("IDRV_PO01_FamigliaProdotti").Value))
        Case "Stato"
            Value = GET_STATO(fnNotNullN(ColumnsValue("IDRV_PO01_StatoLotto").Value))
    End Select

End Sub
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
Private Function GET_VARIETA(IDVarieta As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT Varieta "
sSQL = sSQL & "FROM RV_PO01_Varieta "
sSQL = sSQL & "WHERE IDRV_PO01_Varieta=" & IDVarieta
    
Set rs = Cn.OpenResultset(sSQL)
    
    GET_VARIETA = ""
    
    While Not rs.EOF
        GET_VARIETA = GET_VARIETA & Trim(fnNotNull(rs!Varieta))
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_FAMIGLIA(IDFamiglia As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT FamigliaProdotti "
sSQL = sSQL & "FROM RV_PO01_FamigliaProdotti "
sSQL = sSQL & "WHERE IDRV_PO01_FamigliaProdotti=" & IDFamiglia
    
Set rs = Cn.OpenResultset(sSQL)
    
    GET_FAMIGLIA = ""
    
    While Not rs.EOF
        GET_FAMIGLIA = GET_FAMIGLIA & Trim(fnNotNull(rs!FamigliaProdotti))
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SETTORE(IDSerra As Long, IDSchema As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema, RV_PO01_Settore.Codice "
sSQL = sSQL & "FROM RV_PO01_SettoreSchema LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_Settore ON RV_PO01_SettoreSchema.IDRV_PO01_Settore = RV_PO01_Settore.IDRV_PO01_Settore LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_SettoreSerra ON RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema = RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema "
sSQL = sSQL & "WHERE (RV_PO01_SettoreSchema.IDRV_PO01_Schema = " & IDSchema & ") AND (RV_PO01_SettoreSerra.IDRV_PO01_Serra = " & IDSerra & ") "
    
Set rs = Cn.OpenResultset(sSQL)
    
    GET_SETTORE = ""
    
    While Not rs.EOF
        GET_SETTORE = GET_SETTORE & Trim(fnNotNull(rs!Codice)) & " - (" & GET_TERRENO_SETTORE(fnNotNullN(rs!IDRV_PO01_SettoreSchema)) & ")"
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_STATO(IDStato As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT * FROM RV_PO01_StatoLotto "
sSQL = sSQL & "WHERE IDRV_PO01_StatoLotto=" & IDStato
Set rs = Cn.OpenResultset(sSQL)
    
    GET_STATO = ""
    
    While Not rs.EOF
        GET_STATO = GET_STATO & Trim(fnNotNull(rs!StatoLotto))
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TERRENO_SETTORE(IDSettoreSchema As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT RV_PO01_SettoreTerreni.IDRV_PO01_SettoreTerreni, RV_PO01_SettoreTerreni.IDRV_PO01_SettoreSchema, "
sSQL = sSQL & "RV_PO01_SettoreTerreni.IDRV_PO01_Settore, RV_PO01_SettoreTerreni.IDRV_PO01_TerrenoTesta, RV_PO01_SettoreTerreni.DimensioneMQ,"
sSQL = sSQL & "RV_PO01_SettoreTerreni.DimensioneHA, RV_PO01_TerrenoTesta.Terreno, RV_PO01_TerrenoTesta.Indirizzo, RV_PO01_TerrenoTesta.IDComune,"
sSQL = sSQL & "RV_PO01_TerrenoTesta.Cap, Comune.Comune, RV_PO01_TerrenoRighe.ProprietaAffitto, RV_PO01_TerrenoRighe.FoglioCatastale,"
sSQL = sSQL & "RV_PO01_TerrenoRighe.MappaCatastale , RV_PO01_TerrenoRighe.ParticellaCatastale "
sSQL = sSQL & "FROM Comune RIGHT OUTER JOIN "
sSQL = sSQL & "RV_PO01_TerrenoTesta ON Comune.IDComune = RV_PO01_TerrenoTesta.IDComune LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_TerrenoRighe ON "
sSQL = sSQL & "RV_PO01_TerrenoTesta.IDRV_PO01_TerrenoTesta = RV_PO01_TerrenoRighe.IDRV_PO01_TerrenoRighe RIGHT OUTER JOIN "
sSQL = sSQL & "RV_PO01_SettoreTerreni ON RV_PO01_TerrenoTesta.IDRV_PO01_TerrenoTesta = RV_PO01_SettoreTerreni.IDRV_PO01_TerrenoTesta "
sSQL = sSQL & "WHERE (RV_PO01_SettoreTerreni.IDRV_PO01_SettoreSchema = " & IDSettoreSchema & ") "
    
Set rs = Cn.OpenResultset(sSQL)
    
    GET_TERRENO_SETTORE = ""
    
    While Not rs.EOF
        GET_TERRENO_SETTORE = GET_TERRENO_SETTORE & "Terreno: " & Trim(fnNotNull(rs!Terreno)) & "; "
        GET_TERRENO_SETTORE = GET_TERRENO_SETTORE & "Indirizzo: " & Trim(fnNotNull(rs!Indirizzo)) & " " & Trim(fnNotNull(rs!Cap)) & " " & Trim(fnNotNull(rs!Comune)) & "; "
        GET_TERRENO_SETTORE = GET_TERRENO_SETTORE & "F.C.: " & Trim(fnNotNull(rs!FoglioCatastale)) & "; "
        GET_TERRENO_SETTORE = GET_TERRENO_SETTORE & "M.C.: " & Trim(fnNotNull(rs!MappaCatastale)) & "; "
        GET_TERRENO_SETTORE = GET_TERRENO_SETTORE & "P.C.: " & Trim(fnNotNull(rs!ParticellaCatastale)) & "; "
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Function

