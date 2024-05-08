Attribute VB_Name = "ModAttivazione"
Public StringaLicenza As String

Public Const CodiceProgramma As String = "01"
Public Const IdentificativoProgramma As Long = 2
Public Const Programma As String = "Bìo"

Public TipoAttivazione As Integer
Public Const NumeroRecordDemo As Long = 10
Public NumeroRecordTabella As Long
Public NumeroSoci As Long
Public StringaLottoStd As String

'Variabili dalla base dati
Private NumeroPostiLavoroDB As Long
Private AziendaUnicaDB As Long
Private InstallazioneDemoDB As Long
Private TipoFilialeDB As String
Private CodiceDiamanteDB As String
Private CodiceAttivazioneDB As String
Private CodiceDiSbloccoDB As String
Private Link_TipoAttivazioneDB As Long

Private CodiceDiamanteOriginale As String

Public Function TipoAttivazioneLicenza() As Long
CodiceDiamanteOriginale = Get_CodiceDiamante

RecuperoDati

    If ControlloLicenza = True Then
        If InstallazioneDemoDB = 1 Then
            TipoAttivazioneLicenza = 2
        Else
            TipoAttivazioneLicenza = 1
        End If
    Else
        TipoAttivazioneLicenza = 0
    End If


End Function


Private Function ControlloLicenza() As Boolean

If InstallazioneDemo = 1 Then
    ControlloLicenza = True
    StringaLicenza = ""
    Exit Function
End If

StringaLicenza = ""
ControlloLicenza = True
If CodiceDiamanteDB <> CodiceDiamanteOriginale Then
    StringaLicenza = StringaLicenza & "La licenza del programma non è compatibile con quella di Diamante" & vbCrLf & vbCrLf
End If

If AziendaUnicaDB = 17 Then
    If TheApp.IDFirm <> CLng(TipoFilialeDB) Then
        StringaLicenza = StringaLicenza & "Licenza non attiva per questa azienda" & vbCrLf & vbCrLf
    End If
End If

If Len(StringaLicenza) > 0 Then
    ControlloLicenza = False
End If

End Function
Private Sub RecuperoDati()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsRighe As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POComponenteTesta WHERE IDRV_POProgramma=" & IdentificativoProgramma
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    NumeroPostiLavoroDB = fnNotNullN(rs!NumeroPostiLavoro)
    InstallazioneDemoDB = IIf((rs!Demo = True), 1, 0)
    TipoFilialeDB = fnNotNullN(rs!TipoFiliale)
    CodiceAttivazioneDB = fnNotNull(rs!CodiceAttivazione)
    CodiceDiSbloccoDB = fnNotNull(rs!CodiceSblocco)
    CodiceDiamanteDB = fnNotNull(rs!CodiceDiamante)
    AziendaUnicaDB = IIf((rs!AziendaUnica = True), 17, 12)
    Link_TipoAttivazioneDB = fnNotNullN(rs!IDRV_POTipoAttivazione)
Else
    NumeroPostiLavoroDB = 0
    InstallazioneDemoDB = 0
    TipoFilialeDB = 0
    CodiceAttivazioneDB = ""
    CodiceDiSbloccoDB = ""
    CodiceDiamanteDB = ""
    AziendaUnicaDB = 0
    Link_TipoAttivazioneDB = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function Get_CodiceDiamante() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT Descrizione FROM ComponenteSwAbilitata WHERE NomeCompSW=" & fnNormString("*IDSW___")

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Get_CodiceDiamante = ""
Else
    Get_CodiceDiamante = fnNotNull(rs!Descrizione)
End If

rs.CloseResultset
Set rs = Nothing


End Function
