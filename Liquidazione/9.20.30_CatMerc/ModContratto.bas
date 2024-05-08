Attribute VB_Name = "ModContratto"
Public Link_Contratto As Long
Public Link_StoriaContratto As Long
Public Mesi_Durata_Contratto As Long
Public Mesi_Rinnovo_Contratto As Long 'Indica ogni quanti mesi il contratto deve essere rinnovato
Public Numero_Rate As Long 'Indica il numero di rate del contratto
Public Pagamento_Anticipato_Periodo As Boolean 'Indica se la rata deve essere pagata ad inizio periodo
Public Mesi_Rate As Long 'Indica ogni quanti mesi devono essere create le rate
Public Sub SviluppoRateContratto(IDContratto As Long)
    Dim sSQL As String
    Dim rsDelete As DmtADOLib.adoResultset
    
    sSQL = "SELECT IDRV_PORateContratto FROM RV_PORateContratto WHERE IDRV_POContratto=" & IDContratto
    Set rsDelete = Cn.OpenResultset(sSQL)
    While Not rsDelete.EOF
        sSQL = "DELETE FROM RV_PORateContratto WHERE ("
        sSQL = sSQL & "(IDRV_PORateContratto=" & rsDelete!IDRV_PORateContratto & ") AND "
        sSQL = sSQL & "(Adeguamento=" & fnNormBoolean(0) & ") AND "
        sSQL = sSQL & "(Manuale=" & fnNormBoolean(0) & ") AND "
        sSQL = sSQL & "(ContrattoAttuale=" & fnNormBoolean(1) & "))"
        Cn.Execute sSQL
        
        sSQL = "DELETE FROM RV_POStoriaRateContratto WHERE IDRiferimentoRata=" & rsDelete!IDRV_PORateContratto
        Cn.Execute sSQL

    rsDelete.MoveNext
    Wend
    
    ElaborazioneRate IDContratto, frmMain.txtDataDecorrenza.Text, frmMain.txtImportoAttuale.Value, Mesi_Rate, Numero_Rate, Pagamento_Anticipato_Periodo, frmMain.cboPagamentoRate.CurrentID
    
End Sub
Private Sub ElaborazioneRate(IDContratto As Long, DataDecorrenza As String, ImportoContratto As Double, MesiRate As Long, NumeroRate As Long, PagamentoAnticipato As Boolean, IDPagamentoRata As Long)
    Dim ImportoRata As Double
    Dim UltimaRata As Double
    Dim ImportoRataProgressiva As Double
    Dim DataRataProgressiva As String
    Dim I As Long
    Dim sSQL As String
    Dim IDRata As Long
    Dim DataFinePeriodo As String
    Dim Periodo As String
    
    ImportoRata = 0
    ImportoRataProgressiva = 0
    
    If PagamentoAnticipato = True Then
        DataRataProgressiva = DataDecorrenza
        DataFinePeriodo = DateAdd("m", MesiRate, DataRataProgressiva) - 1
    Else
        DataFinePeriodo = DataDecorrenza
        DataRataProgressiva = DateAdd("m", MesiRate, DataDecorrenza) - 1
    End If
    ImportoRata = FormatNumber((ImportoContratto / NumeroRate), 2)
    
    For I = 1 To NumeroRate
        IDRata = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        sSQL = "INSERT INTO RV_PORateContratto ("
        sSQL = sSQL & "IDRV_PORateContratto, IDRV_POContratto, NumeroRata, DataRata, IDPagamentoRata, ImportoRata, Mese, Anno, Periodo, Adeguamento, Manuale, ContrattoAttuale) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDRata & ", "
        sSQL = sSQL & IDContratto & ", "
        sSQL = sSQL & I & ", "
        sSQL = sSQL & fnNormDate(DataRataProgressiva) & ", "
        sSQL = sSQL & IDPagamentoRata & ", "
        sSQL = sSQL & fnNormNumber(ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(DatePart("m", DataRataProgressiva)) & ", "
        sSQL = sSQL & fnNormNumber(DatePart("yyyy", DataRataProgressiva)) & ", "
        If PagamentoAnticipato = True Then
            Periodo = "Canone " & frmMain.cboTipoRateizzazione.Text & " relativo al contratto " & frmMain.cboTipoContratto.Text & vbCrLf & "Decorrenza dal " & DataRataProgressiva & " al " & DataFinePeriodo
            sSQL = sSQL & fnNormString(Periodo) & ", "
        Else
            Periodo = "Canone " & frmMain.cboTipoRateizzazione.Text & " relativo al contratto " & frmMain.cboTipoContratto.Text & vbCrLf & "Decorrenza dal " & DataFinePeriodo & " al " & DataRataProgressiva
            sSQL = sSQL & fnNormString(Periodo) & ", "
        End If
        sSQL = sSQL & fnNormBoolean(0) & ", "
        sSQL = sSQL & fnNormBoolean(0) & ", "
        sSQL = sSQL & fnNormBoolean(1) & ")"
        Cn.Execute sSQL
        
        AccodaStoriaRata Link_StoriaContratto, IDRata, I, DataRataProgressiva, ImportoRata, IDPagamentoRata, DatePart("m", DataRataProgressiva), DatePart("yyyy", DataRataProgressiva), Periodo, False, False, True
        
        DataRataProgressiva = DateAdd("m", MesiRate, DataRataProgressiva)
        If PagamentoAnticipato = True Then
            DataFinePeriodo = DateAdd("m", MesiRate, DataRataProgressiva) - 1
        Else
            DataFinePeriodo = DateAdd("m", MesiRate, DataRataProgressiva) '+ 1
        End If
        ImportoRataProgressiva = ImportoRataProgressiva + ImportoRata
        If I + 1 = NumeroRate Then
            ImportoRata = ImportoContratto - ImportoRataProgressiva
        End If
    Next
    
End Sub
Public Function ControlloRatePagate(IDContratto As Long) As Boolean
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT * FROM RV_PORateContratto WHERE ("
    sSQL = sSQL & "(IDRV_POContratto=" & IDContratto & ") AND "
    sSQL = sSQL & "(IDOggettoCollegato>0))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        ControlloRatePagate = False
    Else
        ControlloRatePagate = True
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub AccodaStoriaRata(IDStoriaContratto As Long, IDRiferimentoRata As Long, NumeroRata As Long, DataScadenza As String, ImportoRata As Double, IDPagamentoRata As Long, Mese As Long, Anno As Long, Periodo As String, Adeguamento As Boolean, Manuale As Boolean, ContrattoAttuale As Boolean)
    Dim sSQL As String
    
                 
        sSQL = "INSERT INTO RV_POStoriaRateContratto ("
        sSQL = sSQL & "IDRV_POStoriaRateContratto, IDRV_POStoriaContratto, IDRiferimentoRata, NumeroRata, DataRata, IDPagamentoRate, ImportoRata, Mese, Anno, Periodo, Adeguamento, Manuale, ContrattoAttuale) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POStoriaRateContratto", "IDRV_POStoriaRateContratto") & ", "
        sSQL = sSQL & Link_StoriaContratto & ", "
        sSQL = sSQL & IDRiferimentoRata & ", "
        sSQL = sSQL & NumeroRata & ", "
        sSQL = sSQL & fnNormDate(DataScadenza) & ", "
        sSQL = sSQL & IDPagamentoRata & ", "
        sSQL = sSQL & fnNormNumber(ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(Mese) & ", "
        sSQL = sSQL & fnNormNumber(Anno) & ", "
        sSQL = sSQL & fnNormString(Periodo) & ", "
        sSQL = sSQL & fnNormBoolean(Adeguamento) & ", "
        sSQL = sSQL & fnNormBoolean(Manuale) & ", "
        sSQL = sSQL & fnNormBoolean(ContrattoAttuale) & ")"
        
        
        Cn.Execute sSQL
    
End Sub
Public Sub CreaStoriaContratto(IDContratto As Long)
    Dim sSQL As String
    
    Link_StoriaContratto = fnGetNewKey("RV_POStoriaContratto", "IDRV_POStoriaContratto")
    
    sSQL = "INSERT INTO RV_POStoriaContratto ("
    sSQL = sSQL & "IDRV_POStoriaContratto, IDRV_POContratto, DataDecorrenza, DataScadenza, ImportoContratto, "
    sSQL = sSQL & "IDDurataContratto, IDTipoRinnovo, DataRinnovoContratto, IDRateizzazione, IDAdeguamentoIstat, ContrattoAttuale) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_StoriaContratto & ", "
    sSQL = sSQL & IDContratto & ", "
    sSQL = sSQL & fnNormDate(frmMain.txtDataDecorrenza.Value) & ", "
    sSQL = sSQL & fnNormDate(frmMain.txtDataScadenza.Value) & ", "
    sSQL = sSQL & fnNormNumber(frmMain.txtImportoAttuale.Value) & ", "
    sSQL = sSQL & frmMain.cboDurataContratto.CurrentID & ", "
    sSQL = sSQL & frmMain.cboTipoRinnovo.CurrentID & ", "
    sSQL = sSQL & fnNormDate(frmMain.txtDataScadenzaPerRinnovo.Value) & ", "
    sSQL = sSQL & frmMain.cboTipoRateizzazione.CurrentID & ", "
    sSQL = sSQL & frmMain.cboIstat.CurrentID & ", "
    sSQL = sSQL & fnNormBoolean(1) & ")"
    
    Cn.Execute sSQL
    
End Sub
Public Sub AggiornaStoriaContratto(IDStoriaContratto As Long)
    Dim sSQL As String
    sSQL = "UPDATE RV_POStoriaContratto SET "
    sSQL = sSQL & "DataDecorrenza=" & fnNormDate(frmMain.txtDataDecorrenza.Value) & ", "
    sSQL = sSQL & "DataScadenza=" & fnNormDate(frmMain.txtDataScadenza.Value) & ", "
    sSQL = sSQL & "ImportoContratto=" & fnNormNumber(frmMain.txtImportoAttuale.Value) & ", "
    sSQL = sSQL & "IDDurataContratto=" & frmMain.cboDurataContratto.CurrentID & ", "
    sSQL = sSQL & "IDTipoRinnovo=" & frmMain.cboTipoRinnovo.CurrentID & ", "
    sSQL = sSQL & "DataRinnovoContratto=" & fnNormDate(frmMain.txtDataScadenzaPerRinnovo.Value) & ", "
    sSQL = sSQL & "IDRateizzazione=" & frmMain.cboTipoRateizzazione.CurrentID & ", "
    sSQL = sSQL & "IDAdeguamentoIstat=" & frmMain.cboIstat.CurrentID & " "
    sSQL = sSQL & "WHERE IDRV_POStoriaContratto=" & Link_StoriaContratto
    
    Cn.Execute sSQL
End Sub
Public Function ContrattoAttualeStorico(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDRV_POStoriaContratto FROM RV_POStoriaContratto WHERE ("
sSQL = sSQL & "(IDRV_POContratto=" & IDContratto & ") AND "
sSQL = sSQL & "(ContrattoAttuale=1))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ContrattoAttualeStorico = rs!IDRV_POStoriaContratto
Else
    ContrattoAttualeStorico = 0
End If
End Function
Private Sub AggiornaStoriaRata(IDRiferimentoRata As Long)
    Dim sSQL As String
    
                 
        sSQL = "UPDATE  RV_POStoriaRateContratto SET "
        sSQL = sSQL & "DataRata=" & fnNormDate(DataScadenza) & ", "
        sSQL = sSQL & "PagamentoRate=" & IDPagamentoRata & ", "
        sSQL = sSQL & "ImportoRata=" & fnNormNumber(ImportoRata) & ")"
        'sSQL = sSQL & "Fatturata=" & fnNormBoolean(frmMain.chkRataFatturata)
        sSQL = sSQL & "IDRiferimentoRata=" & IDRiferimentoRata
        
        Cn.Execute sSQL
    
End Sub
Private Sub EliminaStoriaRata(IDRiferimentoRata As Long)
    Dim sSQL As String
    
                 
        sSQL = "DELETE FROM RV_POStoriaRateContratto WHERE IDRiferimentoRata=" & IDRiferimentoRata
        Cn.Execute sSQL
    
End Sub
Public Function NumeroRate(IDContratto As Long)
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT NumeroRata FROM RV_POStoriaRateContratto WHERE IDRV_POStoriaContratto=" & IDContratto
    sSQL = sSQL & " ORDER BY NumeroRata DESC"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        NumeroRate = 1
    Else
        NumeroRate = rs!NumeroRata + 1
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
