Attribute VB_Name = "ModTMPLiquidazione"
Private TotaleDocumento As Double
Private TotaleDocumentoRiga As Double
Private TotaleNettoRiga As Double

Private TotaleTrattenutaRiga As Double
Private Link_TestataLiquidazione As Long
Private QuantitaTotaleRiga As Double
Private QuantitaLavorata As Double
Private Quantita_Quadratura_Lavorazione As Double
Private Qta_Totale_Lavorata As Double
Private Quantita_Quadratura_Vendita As Double

Private TotaleIva As Double
Private TotaleIvaDocumento As Double
Private TotaleTrattenutaPerLavorazione As Double
Private TotaleTrattenutaGenerale As Double
Private TotaleTrattenutaPerSocio As Double
Private TotaleImportoLordo As Double

Private TotaleTrattenuta As Double
Private TrattenutaPerLavorazione As Double
Private TrattenuteGenerale As Double

Private rsTesta As DmtOleDbLib.adoResultset
Private rsRighe As DmtOleDbLib.adoResultset
Private rsLav As DmtOleDbLib.adoResultset
Private rsVend As DmtOleDbLib.adoResultset

Private Link_Tipo_CaloPeso As Long
Private Link_Tipo_Scarto As Long
Private Link_Tipo_AumentoPeso As Long




Public Sub ElborazioneTMPLiquidazione(IDLiquidazione As Long, IDPeriodoLiquidazione As Long)
Dim sSQL As String

    
    
Start_Liquidazione IDLiquidazione, IDPeriodoLiquidazione
End Sub

Private Sub Start_Liquidazione(IDLiquidazione As Long, IDPeriodoLiquidazione As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Unita_progresso As Double

sSQL = "SELECT IDAnagrafica "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneRigheConf "
sSQL = sSQL & "WHERE IDRV_POPeriodoLiquidazione =" & IDPeriodoLiquidazione & " "
sSQL = sSQL & "GROUP BY IDAnagrafica "
sSQL = sSQL & "ORDER BY IDAnagrafica"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset

While Not rs.EOF
    Link_TestataLiquidazione = IDLiquidazione
            
    GET_CONFERIMENTI_RIGHE fnNotNullN(rs!IDAnagrafica)
rs.MoveNext
Wend

rs.Close
Set rs = Nothing


End Sub


Private Sub GET_CONFERIMENTI_RIGHE(IDSocio As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsFatt As DmtOleDbLib.adoResultset
Dim rsScarti As DmtOleDbLib.adoResultset

Dim Cont As Integer
Dim TrattenuteTotali As Double
Dim TipoRiga As Integer

sSQL = "SELECT * FROM RV_POTMPLiquidazioneRigheConf "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDSocio

Set rs = Cn.OpenResultset(sSQL)

    GET_RIGA_VENDITA_SENZA_RIFERIMENTI IDSocio

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POTMPLiquidazioneVendita "
    sSQL = sSQL & "WHERE IDSocio=" & IDSocio
    sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    
    Set rsFatt = Cn.OpenResultset(sSQL)
    If rsFatt.EOF Then
            
        'Inserimento conferimento senza collegamenti
        sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheEla ("
        sSQL = sSQL & "IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo, TipoRiga, IDAnagrafica, "
        sSQL = sSQL & "IDRV_POCaricoMerceRighe, IDArticolo_Conf, CodiceArticolo_Conf, Articolo_Conf, QuantitaConferita, "
        sSQL = sSQL & "Colli_Conf, PesoLordo_Conf, PesoNetto_Conf, Tara_Conf, Pezzi_Conf, "
        sSQL = sSQL & "DataConferimento, NumeroDocumento, "
        sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Link_TestataLiquidazione & ", "
        sSQL = sSQL & LINK_PERIODO & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & fnNotNullN(rs!IDAnagrafica) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_POCaricoMerceRighe) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDArticolo) & ", "
        sSQL = sSQL & fnNormString(GET_ARTICOLO(fnNotNullN(rs!IDArticolo), True)) & ", "
        sSQL = sSQL & fnNormString(GET_ARTICOLO(fnNotNullN(rs!IDArticolo), False)) & ", "
        sSQL = sSQL & fnNormNumber(rs!Quantita) & ", "
        sSQL = sSQL & fnNormNumber(rs!Colli) & ", "
        sSQL = sSQL & fnNormNumber(rs!PesoLordo) & ", "
        sSQL = sSQL & fnNormNumber(rs!PesoNetto) & ", "
        sSQL = sSQL & fnNormNumber(rs!Tara) & ", "
        sSQL = sSQL & fnNormNumber(rs!Pezzi) & ", "
        sSQL = sSQL & fnNormDate(rs!DataDocumento) & ", "
        sSQL = sSQL & fnNormNumber(rs!NumeroDocumento) & ", "
        sSQL = sSQL & fnNormNumber(rs!IDCategoriaMerceologica) & ", "
        sSQL = sSQL & fnNormNumber(0) & ", "
        sSQL = sSQL & fnNormNumber(rs!Trattenuta) & ", "
        sSQL = sSQL & fnNormNumber(rs!Trattenuta) & ")"
        
        CnDMT.Execute sSQL
        
    Else
        'Inserimento vendite con collegamento alle righe di conferimento
        While Not rsFatt.EOF
            sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheEla ("
            sSQL = sSQL & "IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo, TipoRiga, IDAnagrafica, "
            sSQL = sSQL & "IDArticolo_Conf, CodiceArticolo_Conf, Articolo_Conf, QuantitaConferita, "
            sSQL = sSQL & "Colli_Conf, PesoLordo_Conf, PesoNetto_Conf, Tara_Conf, Pezzi_Conf, "
            sSQL = sSQL & "IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
            sSQL = sSQL & "QuantitaLavorata, QuantitaQuadrata, QuantitaTotaleLavorata, QuantitaQuadrataDiVendita, "
            sSQL = sSQL & "Colli, PesoNetto, PesoLordo, Tara, Pezzi, IDRV_POCaricoMerceRighe, "
            sSQL = sSQL & "DataConferimento, NumeroDocumento, IDOggetto, IDTipoOggetto, IDValoriOggettoDettaglioArticolo, "
            sSQL = sSQL & "QuantitaVenduta, DataDocumentoVendita, Oggetto, "
            sSQL = sSQL & "IDIva_per_imp_vend, AliquotaIva_per_imp_Vend, CodiceIva_per_imp_vend, Iva_per_imp_vend, ImportoUnitarioVend, "
            sSQL = sSQL & "IDIva_per_imp_medio, AliquotaIva_per_imp_medio, CodiceIva_per_imp_medio, Iva_per_imp_medio, ImportoUnitarioMedio, "
            sSQL = sSQL & "ImponibileVenduto, ImponibileMedio, ImpostaImponibileVenduto, ImpostaImponibileMedio, ImportoLordoVenduto, ImportoLordoMedio, "
            sSQL = sSQL & "ImportoUnitarioDaReg, ImponibileDaReg, ImpostaDaReg, ImportoLordoDaReg, "
            sSQL = sSQL & "Colli_Vend, PesoLordo_vend, PesoNetto_Vend, Tara_vend, Pezzi_vend, "
            sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_Totali) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & Link_TestataLiquidazione & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & rsFatt!TipoRiga & ", "
            sSQL = sSQL & fnNotNullN(rs!IDAnagrafica) & ", "
            sSQL = sSQL & fnNotNullN(rs!IDArticolo) & ", "
            sSQL = sSQL & fnNormString(rs!CodiceArticolo) & ", "
            sSQL = sSQL & fnNormString(rs!Articolo) & ", "
            sSQL = sSQL & fnNormNumber(rs!Quantita) & ", "
            sSQL = sSQL & fnNormNumber(rs!Colli) & ", "
            sSQL = sSQL & fnNormNumber(rs!PesoLordo) & ", "
            sSQL = sSQL & fnNormNumber(rs!PesoNetto) & ", "
            sSQL = sSQL & fnNormNumber(rs!Tara) & ", "
            sSQL = sSQL & fnNormNumber(rs!Pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Articolo) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoLavorazione) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDCalibro) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoCategoria) & ", "
            If rsFatt!TipoRiga = 3 Then
                sSQL = sSQL & fnNormNumber(0) & ", "
                sSQL = sSQL & "0" & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
            ElseIf rsFatt!TipoRiga = 2 Then
                sSQL = sSQL & "0" & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Quantita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Quantita) & ", "
            Else
                sSQL = sSQL & fnNormNumber(rsFatt!Quantita) & ", "
                sSQL = sSQL & "0" & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Quantita) & ", "
            End If
            sSQL = sSQL & "0" & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!PesoLordo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!PesoNetto) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rs!IDRV_POCaricoMerceRighe) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!DataConferimento) & ", "
            sSQL = sSQL & fnNormNumber(GET_NUMERO_CONFERIMENTO(rs!IDRV_POCaricoMerceRighe)) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDValoreOggettoDettaglio) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Quantita) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!DataDocumentoVendita) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Oggetto) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIva_Vend) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIva_vend) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIva_Vend) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Iva_Vend) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoUnitario) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIva_medio) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIva_medio) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIva_medio) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Iva_Medio) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoMedioPeriodo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoNettoTotale) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoTotaleSuPeriodo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImpostaTotaleIva) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImpostaImportoMedio) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoTotaleLordoIva) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!ImportoTotaleMedioLordoIva) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!PrezzoUnitarioDaReg) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!TotaleNettoRigaDaReg) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!TotaleImpostaDaReg) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!TotaleLordoRigaDaReg) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!PesoLordo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!PesoNetto) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Pezzi) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDCategoriaMerceologica) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!TrattenutePerLavorazione) & ", "
            
            If EsistenzaArticoloConferitoElaborato(rs!IDRV_POCaricoMerceRighe) = False Then
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!TrattenuteGenerali) + fnNotNullN(rs!Trattenuta)) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!TrattenuteTotali) + fnNotNullN(rs!Trattenuta)) & ", "
            Else
                sSQL = sSQL & fnNormNumber(rsFatt!TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!TrattenuteTotali) & ", "
            End If
            
            sSQL = sSQL & fnNormNumber(rsFatt!Quantita_Per_Totali) & ")"
            
            Cn.Execute sSQL
                                
        rsFatt.MoveNext
        Wend

    rsFatt.CloseResultset
    Set rsFatt = Nothing
        
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub GET_RIGA_VENDITA_SENZA_RIFERIMENTI(IDSocio As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTMPLiquidazioneVendita "
sSQL = sSQL & "WHERE IDSocio=" & IDSocio
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=0"

Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF
    sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheEla ("
    sSQL = sSQL & "IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo, TipoRiga, IDAnagrafica, "
    sSQL = sSQL & "IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
    sSQL = sSQL & "QuantitaLavorata, QuantitaQuadrata, QuantitaTotaleLavorata, QuantitaQuadrataDiVendita, "
    sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
    sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, NumeroDocumento, IDOggetto, IDTipoOggetto, IDValoriOggettoDettaglioArticolo, "
    sSQL = sSQL & "QuantitaVenduta, DataDocumentoVendita, Oggetto, "
    sSQL = sSQL & "IDIva_per_imp_vend, AliquotaIva_per_imp_Vend, CodiceIva_per_imp_vend, Iva_per_imp_vend, ImportoUnitarioVend, "
    sSQL = sSQL & "IDIva_per_imp_medio, AliquotaIva_per_imp_medio, CodiceIva_per_imp_medio, Iva_per_imp_medio, ImportoUnitarioMedio, "
    sSQL = sSQL & "ImponibileVenduto, ImponibileMedio, ImpostaImponibileVenduto, ImpostaImponibileMedio, ImportoLordoVenduto, ImportoLordoMedio, "
    sSQL = sSQL & "ImportoUnitarioDaReg, ImponibileDaReg, ImpostaDaReg, ImportoLordoDaReg, "
    sSQL = sSQL & "Colli_Vend, PesoLordo_vend, PesoNetto_Vend, Tara_vend, Pezzi_vend, "
    sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_TestataLiquidazione & ", "
    sSQL = sSQL & LINK_PERIODO & ", "
    sSQL = sSQL & 3 & ", "
    sSQL = sSQL & fnNotNullN(rs!IDSocio) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDArticolo) & ", "
    sSQL = sSQL & fnNormString(GET_ARTICOLO(fnNotNullN(rs!IDArticolo), True)) & ", "
    sSQL = sSQL & fnNormString(GET_ARTICOLO(fnNotNullN(rs!IDArticolo), False)) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDTipoLavorazione) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDCalibro) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDTipoCategoria) & ", "
    sSQL = sSQL & fnNormNumber(rs!Quantita) & ", "
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & fnNormNumber(rs!Quantita) & ", "
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & fnNormNumber(rs!Colli) & ", "
    sSQL = sSQL & fnNormNumber(rs!PesoLordo) & ", "
    sSQL = sSQL & fnNormNumber(rs!PesoNetto) & ", "
    sSQL = sSQL & fnNormNumber(rs!Tara) & ", "
    sSQL = sSQL & fnNormNumber(rs!Pezzi) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDRV_POCaricoMerceRighe) & ", "
    sSQL = sSQL & fnNormDate(rs!DataConferimento) & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & fnNotNullN(rs!IDOggetto) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDTipoOggetto) & ", "
    sSQL = sSQL & fnNotNullN(rs!IDValoreOggettoDettaglio) & ", "
    sSQL = sSQL & fnNormNumber(rs!Quantita) & ", "
    sSQL = sSQL & fnNormDate(rs!DataDocumentoVendita) & ", "
    sSQL = sSQL & fnNormString(rs!Oggetto) & ", "
    sSQL = sSQL & fnNormNumber(rs!IDIva_Vend) & ", "
    sSQL = sSQL & fnNormNumber(rs!AliquotaIva_vend) & ", "
    sSQL = sSQL & fnNormString(Trim(rs!CodiceIva_Vend)) & ", "
    sSQL = sSQL & fnNormString(rs!Iva_Vend) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoUnitario) & ", "
    sSQL = sSQL & fnNormNumber(rs!IDIva_medio) & ", "
    sSQL = sSQL & fnNormNumber(rs!AliquotaIva_medio) & ", "
    sSQL = sSQL & fnNormString(Trim(rs!CodiceIva_medio)) & ", "
    sSQL = sSQL & fnNormString(rs!Iva_Medio) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoMedioPeriodo) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoNettoTotale) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoTotaleSuPeriodo) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImpostaTotaleIva) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImpostaTotaleMedioIva) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoTotaleLordoIva) & ", "
    sSQL = sSQL & fnNormNumber(rs!ImportoTotaleMedioLordoIva) & ", "
    sSQL = sSQL & fnNormNumber(rs!PrezzoUnitarioDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rs!TotaleNettoRigaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rs!TotaleImpostaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rs!TotaleLordoRigaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rs!Colli) & ", "
    sSQL = sSQL & fnNormNumber(rs!PesoLordo) & ", "
    sSQL = sSQL & fnNormNumber(rs!PesoNetto) & ", "
    sSQL = sSQL & fnNormNumber(rs!Tara) & ", "
    sSQL = sSQL & fnNormNumber(rs!Pezzi) & ", "
    sSQL = sSQL & fnNormNumber(rs!IDCategoriaMerceologica) & ", "
    sSQL = sSQL & fnNormNumber(rs!TrattenutePerLavorazione) & ", "
    sSQL = sSQL & fnNormNumber(rs!TrattenuteGenerali) & ", "
    sSQL = sSQL & fnNormNumber(rs!TrattenuteTotali) & ")"
    
    CnDMT.Execute sSQL

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function ParametroTipoCaloPeso() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCaloPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & VarIDFiliale & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroTipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
Else
    ParametroTipoCaloPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function ParametroTipoAumentoPeso() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoAumentoPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & VarIDFiliale & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroTipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
Else
    ParametroTipoAumentoPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function ParametroTipoScarto() As Long


Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoScarto FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & VarIDFiliale & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroTipoScarto = fnNotNullN(rs!IDTipoScarto)
Else
    ParametroTipoScarto = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function EsistenzaArticoloConferitoElaborato(IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POTMPLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDRV_POLiquidazionePeriodo=" & LINK_PERIODO

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    EsistenzaArticoloConferitoElaborato = False
Else
    EsistenzaArticoloConferitoElaborato = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TotaleDocumento(IDLotto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(TotaleNettoRigaDaReg) as TotaleNettoRiga "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneVendita "
sSQL = sSQL & "WHERE IDLotto=" & IDLotto
sSQL = sSQL & " AND RV_POLiquidazionePeriodo=" & LINK_PERIODO

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TotaleDocumento = 0
Else
    GET_TotaleDocumento = fnNotNullN(rs!TotaleNettoRiga)
End If

rs.CloseResultset
Set rs = Nothing

End Function


Private Sub InserimentoElaborazione()
Dim sSQL As String

sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheEla ("
sSQL = sSQL & "IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo, IDRV_POLavorazione, IDLottoArticolo, CodiceLottoArticolo, LottoArticolo, IDArticolo, "
sSQL = sSQL & "CodiceArticolo, Articolo, IDImballo, CodiceImballo, Imballo, TaraUnitariaImballo, IDTipoImballo, PrezzoImballoListino_Lav, "
sSQL = sSQL & "IDTipoLavorazione, IDTipoCategoria, Calibro, "
sSQL = sSQL & "QuantitaLavorata, QuantitaQuadrata, QuantitaTotaleLavorata, QuantitaQuadrataDiVendita, Colli, PesoNetto, PesoLordo, Tara, Pezzi, "
sSQL = sSQL & "IDRV_POCaricoMerceRighe, IDLottoArticolo_Conf, CodiceLottoArticolo_Conf, LottoArticolo_Conf, IDArticolo_Conf, CodiceArticolo_Conf, "
sSQL = sSQL & "Articolo_Conf, IDImballo_Conf, CodiceImballo_Conf, Imballo_Conf, PrezzoImballoListino_Conf, TaraUnitariaImballo_Conf, IDTipoImballo_Conf, "
sSQL = sSQL & "QuantitaConferita, Colli_Conf, PesoNetto_Conf, PesoLordo_Conf, Tara_Conf, Pezzi_Conf, "
sSQL = sSQL & "IDAnagrafica, Anagrafica, Nome, Indirizzo, Comune, Provincia, Cap, DataConferimento, NumeroDocumento, "
sSQL = sSQL & "IDOggetto, IDTipoOggetto, IDValoriOggettoDettaglioArticolo, Oggetto, DataDocumentoVendita, QuantitaVenduta, IDIva_per_imp_Vend, "
sSQL = sSQL & "AliquotaIva_per_Imp_vend, Iva_per_Imp_Vend, CodiceIva_per_Imp_Vend, ImportoUnitarioVend, "
sSQL = sSQL & "IDIva_per_imp_Medio, AliquotaIva_per_Imp_Medio, Iva_per_Imp_Medio, CodiceIva_per_Imp_Medio, ImportoUnitarioMedio, "
sSQL = sSQL & "ImponibileVenduto, ImponibileMedio, ImpostaImponibileVenduto, ImpostaImponibileMedio, ImportoLordoVenduto, ImportoLordoMedio, "
sSQL = sSQL & "ImportoUnitarioDaReg, ImponibileDaReg, ImpostaDaReg, ImportoLordoDaReg, "
sSQL = sSQL & "Colli_Vend, PesoLordo_Vend, PesoNetto_Vend, Tara_Vend, Pezzi_Vend, "
sSQL = sSQL & "IDValoriOggettoDettaglioImballo, IDImballo_Vend, CodiceImballo_Vend, Imballo_Vend, IDTipoImballo_Vend, PrezzoImballo_Vend, "
sSQL = sSQL & "PrezzoImballoListino_Vend, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
sSQL = sSQL & "VALUES ("
sSQL = sSQL & Link_TestataLiquidazione & ", "
sSQL = sSQL & LINK_PERIODO & ", "
sSQL = sSQL & rsLav!IDRV_POLavorazione & ", "
sSQL = sSQL & rsLav!IDLotto & ", "
sSQL = sSQL & fnNormString(rsLav!CodiceLottoVendita) & ", "
sSQL = sSQL & fnNormString(rsLav!DescrizioneLottoVendita) & ", "
sSQL = sSQL & fnNotNullN(rsLav!IDArticolo) & ", "
sSQL = sSQL & fnNormString(rsLav!CodiceArticolo) & ", "
sSQL = sSQL & fnNormString(rsLav!Articolo) & ", "
sSQL = sSQL & fnNotNullN(rsLav!IDImballo) & ", "
sSQL = sSQL & fnNormString(rsLav!CodiceImballoVendita) & ", "
sSQL = sSQL & fnNormString(rsLav!ImballoVendita) & ", "
sSQL = sSQL & fnNormNumber(rsLav!TaraUnitaria) & ", "
sSQL = sSQL & GET_TIPOIMBALLO(fnNotNullN(rsLav!IDImballo)) & ", "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & fnNotNullN(rsLav!IDTipoLavorazione) & ", "
sSQL = sSQL & fnNotNullN(rsLav!IDRV_POTipoCategoria) & ", "
sSQL = sSQL & fnNormString(rsLav!Calibro) & ", "
If (rsLav!IDTipoProdotto = Link_Tipo_CaloPeso) Or (rsLav!IDTipoProdotto = Link_Tipo_Scarto) Or (rsLav!IDTipoProdotto = Link_Tipo_AumentoPeso) Then
    sSQL = sSQL & fnNormNumber(0) & ", "
Else
    sSQL = sSQL & fnNormNumber(rsLav!Quantita) & ", "
End If
If (rsLav!IDTipoProdotto = Link_Tipo_CaloPeso) Or (rsLav!IDTipoProdotto = Link_Tipo_Scarto) Or (rsLav!IDTipoProdotto = Link_Tipo_AumentoPeso) Then
    sSQL = sSQL & fnNormNumber(rsLav!Quantita) & ", "
Else
    sSQL = sSQL & fnNormNumber(Quantita_Quadratura_Lavorazione) & ", "
End If
sSQL = sSQL & fnNormNumber(QuantitaTotaleRiga) & ", "
sSQL = sSQL & fnNormNumber(GET_QUADRATURA_DI_VENDITA(fnNotNullN(rsLav!IDRV_POLavorazione))) & ", "
sSQL = sSQL & fnNormNumber(rsLav!Colli) & ", "
sSQL = sSQL & fnNormNumber(rsLav!PesoNetto) & ", "
sSQL = sSQL & fnNormNumber(rsLav!PesoLordo) & ", "
sSQL = sSQL & fnNormNumber(rsLav!Tara) & ", "
sSQL = sSQL & fnNormNumber(rsLav!Pezzi) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!IDRV_POCaricoMerceRighe) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!IDCodiceLotto) & ", "
sSQL = sSQL & fnNormString(rsRighe!CodiceLotto) & ", "
sSQL = sSQL & fnNormString(rsRighe!DescrizioneLotto) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!IDArticolo) & ", "
sSQL = sSQL & fnNormString(rsRighe!CodiceArticolo) & ", "
sSQL = sSQL & fnNormString(rsRighe!Articolo) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!IDImballo) & ", "
sSQL = sSQL & fnNormString(rsRighe!CodiceImballo) & ", "
sSQL = sSQL & fnNormString(rsRighe!DescrizioneImballo) & ", "
sSQL = sSQL & fnNormNumber(0) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!TaraUnitaria) & ", "
sSQL = sSQL & GET_TIPOIMBALLO(rsRighe!IDImballo) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!Qta_UM) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!Colli) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!PesoNetto) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!PesoLordo) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!Tara) & ", "
sSQL = sSQL & fnNormNumber(rsRighe!Pezzi) & ", "
sSQL = sSQL & fnNotNullN(rsTesta!IDAnagrafica) & ", "
sSQL = sSQL & fnNormString(rsTesta!Anagrafica) & ", "
sSQL = sSQL & fnNormString(rsTesta!Nome) & ", "
sSQL = sSQL & fnNormString(rsTesta!Indirizzo) & ", "
sSQL = sSQL & fnNormString(rsTesta!Comune) & ", "
sSQL = sSQL & fnNormString(rsTesta!Provincia) & ", "
sSQL = sSQL & fnNormString(rsTesta!Cap) & ", "
sSQL = sSQL & fnNormDate(rsTesta!DataDocumento) & ", "
sSQL = sSQL & fnNormNumber(rsTesta!NumeroDocumento) & ", "
If rsVend.EOF = False Then
    sSQL = sSQL & fnNormNumber(rsVend!IDOggetto) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!IDTipoOggetto) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!IDValoreOggettoDettaglio) & ", "
    Select Case fnNotNullN(rsVend!IDTipoOggetto)
        Case 2
            sSQL = sSQL & fnNormString("D.d.t. n° " & GET_NUMERODOCUMENTO_VENDITA(fnNotNullN(rsVend!IDOggetto), fnNotNullN(rsVend!IDTipoOggetto)) & " del " & fnNotNull(rsVend!DataDocumentoVendita)) & ", "
        Case 114
            sSQL = sSQL & fnNormString("F.A. n° " & GET_NUMERODOCUMENTO_VENDITA(fnNotNullN(rsVend!IDOggetto), fnNotNullN(rsVend!IDTipoOggetto)) & " del " & fnNotNull(rsVend!DataDocumentoVendita)) & ", "
        Case 8
            sSQL = sSQL & fnNormString("S.N.F. n° " & GET_NUMERODOCUMENTO_VENDITA(fnNotNullN(rsVend!IDOggetto), fnNotNullN(rsVend!IDTipoOggetto)) & " del " & fnNotNull(rsVend!DataDocumentoVendita)) & ", "
        Case Else
            sSQL = sSQL & fnNormString("") & ", "
    End Select
    sSQL = sSQL & fnNormDate(rsVend!DataDocumentoVendita) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!Quantita) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!IDIva_Vend) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!AliquotaIva_vend) & ", "
    sSQL = sSQL & fnNormString(rsVend!Iva_Vend) & ", "
    sSQL = sSQL & fnNormString(rsVend!CodiceIva_Vend) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoUnitario) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!IDIva_medio) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!AliquotaIva_medio) & ", "
    sSQL = sSQL & fnNormString(rsVend!Iva_Medio) & ", "
    sSQL = sSQL & fnNormString(rsVend!CodiceIva_medio) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoMedioPeriodo) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoNettoTotale) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoTotaleSuPeriodo) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImpostaTotaleIva) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImpostaTotaleMedioIva) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoTotaleLordoIva) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!ImportoTotaleMedioLordoIva) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!PrezzoUnitarioDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!TotaleNettoRigaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!TotaleImpostaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(rsVend!TotaleLordoRigaDaReg) & ", "
    sSQL = sSQL & fnNormNumber(GET_VALORI_UNITA_COOP(fnNotNullN(rsVend!IDTipoOggetto), fnNotNullN(rsVend("IDValoreOggettoDettaglio")), 1)) & ", "
    sSQL = sSQL & fnNormNumber(GET_VALORI_UNITA_COOP(fnNotNullN(rsVend!IDTipoOggetto), fnNotNullN(rsVend("IDValoreOggettoDettaglio")), 2)) & ", "
    sSQL = sSQL & fnNormNumber(GET_VALORI_UNITA_COOP(fnNotNullN(rsVend!IDTipoOggetto), fnNotNullN(rsVend("IDValoreOggettoDettaglio")), 3)) & ", "
    sSQL = sSQL & fnNormNumber(GET_VALORI_UNITA_COOP(fnNotNullN(rsVend!IDTipoOggetto), fnNotNullN(rsVend("IDValoreOggettoDettaglio")), 4)) & ", "
    sSQL = sSQL & fnNormNumber(GET_VALORI_UNITA_COOP(fnNotNullN(rsVend!IDTipoOggetto), fnNotNullN(rsVend("IDValoreOggettoDettaglio")), 5)) & ", "
    sSQL = sSQL & "0,0,0,0,0,0,0,"

Else
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & fnNormDate("") & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & fnNormNumber(0) & ", "
    sSQL = sSQL & "0,0,0,0,0,0,0,"
End If
sSQL = sSQL & fnNormNumber(TrattenutaPerLavorazione) & ","
sSQL = sSQL & fnNormNumber(TrattenuteGenerale) & ","
sSQL = sSQL & fnNormNumber(TotaleTrattenutaRiga) & ")"

CnDMT.Execute sSQL
End Sub


Private Function GET_TIPOIMBALLO(IDImballo) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDTipoImballo FROM Articolo WHERE IDArticolo=" & IDImballo
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPOIMBALLO = 0
Else
    GET_TIPOIMBALLO = fnNotNullN(rs!RV_POIDTipoImballo)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PREZZO_LISTINO_IMBALLO(IDImballo) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDTipoImballo FROM Articolo WHERE IDArticolo=" & IDImballo
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_LISTINO_IMBALLO = 0
Else
    GET_PREZZO_LISTINO_IMBALLO = fnNotNullN(rs!RV_POIDTipoImballo)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_VALORI_UNITA_COOP(IDTipoOggetto As Long, IDValoreOggettoDettaglio As Long, IDCampo As Integer) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Select Case IDTipoOggetto
    Case 2
        sSQL = "SELECT Art_quantita_pezzi, Art_peso, Art_numero_colli, Art_Tara FROM ValoriOggettoDettaglio0004 "
        sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoreOggettoDettaglio
    Case 114
        sSQL = "SELECT Art_quantita_pezzi, Art_peso, Art_numero_colli, Art_Tara FROM ValoriOggettoDettaglio0001 "
        sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoreOggettoDettaglio
    Case 8
        sSQL = "SELECT Art_quantita_pezzi, Art_peso, Art_numero_colli, Art_Tara FROM ValoriOggettoDettaglio0034 "
        sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoreOggettoDettaglio
            
End Select

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_VALORI_UNITA_COOP = 0
Else
    Select Case IDCampo
        Case 1
            GET_VALORI_UNITA_COOP = fnNotNullN(rs!Art_numero_colli)
        Case 2
            GET_VALORI_UNITA_COOP = fnNotNullN(rs!Art_peso)
        Case 3
            GET_VALORI_UNITA_COOP = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
        Case 4
            GET_VALORI_UNITA_COOP = fnNotNullN(rs!Art_tara)
        Case 5
            GET_VALORI_UNITA_COOP = fnNotNullN(rs!Art_quantita_pezzi)
    End Select
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_VALORI_IMBALLO_VENDITA(IDTipoOggetto As Long, IDValoreOggettoDettaglio As Long, IDCampo As Integer) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Select Case IDTipoOggetto
    Case 2
        sSQL = "SELECT IDValoriOggettoDettaglio, Art_Codice, Art_Descrizione, Link_Art_articolo,Art_prezzo_unitario_neutro FROM ValoriOggettoDettaglio0004 "
        sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoreOggettoDettaglio
    Case 114
        sSQL = "SELECT IDValoriOggettoDettaglio, Art_Codice, Art_Descrizione, Link_Art_articolo,Art_prezzo_unitario_neutro FROM ValoriOggettoDettaglio0001 "
        sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoreOggettoDettaglio
            
End Select

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_VALORI_IMBALLO_VENDITA = 0
Else
    Select Case IDCampo
        Case 1
            GET_VALORI_IMBALLO_VENDITA = fnNotNullN(rs!IDValoriOggettoDettaglio)
        Case 2
            GET_VALORI_IMBALLO_VENDITA = fnNotNull(rs!Art_codice)
        Case 3
            GET_VALORI_IMBALLO_VENDITA = fnNotNull(rs!Art_descrizione)
        Case 4
            GET_VALORI_IMBALLO_VENDITA = fnNotNull(rs!Link_Art_Articolo)
        Case 5
            GET_VALORI_IMBALLO_VENDITA = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    End Select
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function Get_Anagrafica(IDSocio As Long) As String
Dim sSQL As String
Dim rsAna As ADODB.Recordset

sSQL = "SELECT Anagrafica, Nome FROM Anagrafica WHERE IDAnagrafica=" & IDSocio

Set rsAna = New ADODB.Recordset
rsAna.Open sSQL, CnDMT.InternalConnection

 If rsAna.EOF = False Then
    Get_Anagrafica = fnNotNull(rsAna!Anagrafica) & " " & fnNotNull(rsAna!Nome)
Else
    Get_Anagrafica = "IDENTIFICATIVO NON TROVATO"
End If

rsAna.Close
Set rsAna = Nothing
End Function
Private Function GET_NUMERODOCUMENTO_VENDITA(IDOggetto As Long, IDTipoOggetto) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Select Case IDTipoOggetto

    Case 2
        sSQL = "SELECT Doc_numero FROM ValoriOggettoPerTipo0002 "
        sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
        sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    Case 114
        sSQL = "SELECT Doc_numero FROM ValoriOggettoPerTipo0072 "
        sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
        sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    Case 8
        sSQL = "SELECT Doc_numero FROM ValoriOggettoPerTipo0008 "
        sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
        sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    
    Case Else
        GET_NUMERODOCUMENTO_VENDITA = 0
        Exit Function
End Select

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERODOCUMENTO_VENDITA = 0
Else
    GET_NUMERODOCUMENTO_VENDITA = fnNotNull(rs!Doc_Numero)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_TRATTENUTA_PER_LAVORAZIONE(IDTipoTrattenuta As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoTrattenuta "
sSQL = sSQL & "WHERE IDRV_POTipoTrattenuta=" & IDTipoTrattenuta

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TRATTENUTA_PER_LAVORAZIONE = 0
Else
    If fnNotNullN(rs!Tipo4) = 1 Then
        GET_TRATTENUTA_PER_LAVORAZIONE = 1
    Else
        GET_TRATTENUTA_PER_LAVORAZIONE = 0
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_TOTALI_PER_SOCIO(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(TrattenuteGenerali) AS TrattenuteGenerali, SUM(TrattenutePerLavorazione) AS TrattenutePerLavorazione, SUM(TrattenuteTotali) "
sSQL = sSQL & "AS TrattenuteTotali, SUM(ImponibileDaReg) AS ImponibileDaReg, SUM(ImpostaDaReg) AS ImpostaDaReg, SUM(ImportoLordoDaReg) "
sSQL = sSQL & "AS ImportoLordoDaReg, IDAnagrafica "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneRigheEla "
sSQL = sSQL & "GROUP BY IDAnagrafica "
sSQL = sSQL & "HAVING (IDAnagrafica =" & IDAnagrafica & ")"

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    TotaleIva = 0
    TotaleDocumento = 0
    TotaleTrattenutaPerLavorazione = 0
    TotaleTrattenutaGenerale = 0
    TotaleTrattenutaPerSocio = 0
    TotaleImportoLordo = 0
Else
    TotaleIva = fnNotNullN(rs!ImpostaDaReg)
    TotaleDocumento = fnNotNullN(rs!ImponibileDaReg)
    TotaleTrattenutaPerLavorazione = fnNotNullN(rs!TrattenutePerLavorazione)
    TotaleTrattenutaGenerale = fnNotNullN(rs!TrattenuteGenerali)
    TotaleTrattenutaPerSocio = fnNotNullN(rs!TrattenuteTotali)
    TotaleImportoLordo = fnNotNullN(rs!ImportoLordoDaReg)

End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function GET_QUADRATURA_DI_VENDITA(IDLavorazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_QUADRATURA_DI_VENDITA = 0

sSQL = "SELECT * FROM RV_POQuadraturaVendita WHERE IDRV_POLavorazione=" & IDLavorazione

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_QUADRATURA_DI_VENDITA = GET_QUADRATURA_DI_VENDITA + fnNotNullN(rs!QtaCausale)
Wend

rs.CloseResultset
Set rs = Nothing

End Function

Private Function ControllaEsistenzaSocio(IDSocio As Long, IDPeriodo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM RV_POTMPLiquidazione "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDSocio & " AND "
sSQL = sSQL & "IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    ControllaEsistenzaSocio = False
Else
    ControllaEsistenzaSocio = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ARTICOLO(IDArticolo As Long, Codice As Boolean) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT CodiceArticolo, Articolo FROM Articolo "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    sSQL = sSQL & " AND IDAzienda=" & VarIDAzienda
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_ARTICOLO = ""
    Else
        If Codice = True Then
            GET_ARTICOLO = fnNotNull(rs!CodiceArticolo)
        Else
            GET_ARTICOLO = fnNotNull(rs!Articolo)
        End If
    End If
    
rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_NUMERO_CONFERIMENTO(IDRigaConferimento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NumeroDocumento "
sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_CONFERIMENTO = 0
Else
    GET_NUMERO_CONFERIMENTO = fnNotNullN(rs!NumeroDocumento)
End If

rs.CloseResultset
Set rs = Nothing


End Function
