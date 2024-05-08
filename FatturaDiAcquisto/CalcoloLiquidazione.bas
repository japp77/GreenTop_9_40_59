Attribute VB_Name = "CalcoloLiquidazione"
Dim rsTesta As ADODB.Recordset
Dim rsRighe As ADODB.Recordset
Dim rsLav As ADODB.Recordset
Private ArrayQuad(3, 3) As String
Private ArrayTratt(9, 4) As String
Private Link_AddebitoImballo As Integer
Private Link_ListinoImballo As Long
Private Link_Iva As Long
Private Codice_Iva As String
Private Aliquota_Iva As Double
Private Link_Iva_Medio As Long
Private Codice_Iva_Medio As String
Private Aliquota_Iva_Medio As Double
Private Iva_Medio As String
'Variabile per le trattenute
Private TrattenutePerLavorazione As Double
Private TrattenuteGenerali As Double
Private TrattenuteTotali As Double

Private TotaleRighe As Long
Public Sub EsecuzioneElaborazione(IDConferimentoRighe As Long, IDLiquidazione As Long, IDPeriodoLiquidazione As Long)
Dim Data_Minima_Conf As String
Dim Data_Massima_Conf As String

Cn.Execute "DELETE FROM RV_POTMPLiquidazioneRigheConf"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneLavorazione"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneVendita"

If TIPO_LIQUIDAZIONE = 1 Then
    fnCalcolaPrezzoUnitario IDConferimentoRighe, TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
    ElaborazioneConferimento IDConferimentoRighe, DATA_INIZIO, DATA_FINE
    GET_VENDITA_DDT IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_FA IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_SNF IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
End If

If TIPO_LIQUIDAZIONE = 2 Then
    fnCalcolaPrezzoUnitario IDConferimentoRighe, TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
    GET_VENDITA_DDT IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_FA IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_SNF IDConferimentoRighe, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    'Data_Minima_Conf = GET_DATA_MINIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    'Data_Massima_Conf = GET_DATA_MASSIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    ElaborazioneConferimento IDConferimentoRighe, DATA_INIZIO, DATA_FINE
    
End If

ElborazioneTMPLiquidazione IDLiquidazione, IDPeriodoLiquidazione
End Sub
Private Sub ElaborazioneConferimento(IDConferimentoRighe As Long, DataInizio As String, DataFine As String)
Dim sSQL As String
Dim rsSocio As DmtOleDbLib.adoResultset
Dim rsRighe As ADODB.Recordset
Dim TrattenutaConferimento As Double
Dim LINK_SOCIO_LOCAL As Long
Dim Link_CategoriaMerceologica As Long
Dim Unita_progresso As Double







sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe= " & IDConferimentoRighe


sSQL = sSQL & " ORDER BY RV_POCaricoMerceTesta.IDAnagrafica "

Set rsRighe = New ADODB.Recordset
rsRighe.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsRighe.EOF = False Then
    
    While Not rsRighe.EOF
        
        
        Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsRighe!IDArticolo))
        
        GET_SCARTI rsRighe!IDAnagrafica, rsRighe!IDRV_POCaricoMerceRighe
        
        
        TrattenutePerLavorazione = 0
        TrattenuteGenerali = 0
        TrattenuteTotali = 0

        If ((rsRighe!DataDocumento >= DATA_INIZIO) And (rsRighe!DataDocumento <= DATA_FINE)) Then
            If TIPO_QUANTITA = 3 Then
                TrattenuteArticolo fnNotNullN(rsRighe!IDAnagrafica), rsRighe!IDArticolo, Link_CategoriaMerceologica, 0, True, fnNotNullN(rsRighe!Qta_UM), 0
            End If
        End If
        
            sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheConf ("
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, CodiceArticolo, Articolo, IDRV_POCaricoMerceTesta, IDRV_POPeriodoLiquidazione, IDArticolo,  "
            sSQL = sSQL & "IDImballo, Quantita, Colli, PesoLordo, Tara, PesoNetto, Pezzi, "
            sSQL = sSQL & "Trattenuta, IDCategoriaMerceologica, IDAnagrafica, Anagrafica, Nome, "
            sSQL = sSQL & "NumeroDocumento, DataDocumento) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & rsRighe!IDRV_POCaricoMerceRighe & ", "
            sSQL = sSQL & fnNormString(rsRighe!CodiceArticolo) & ", "
            sSQL = sSQL & fnNormString(rsRighe!Articolo) & ", "
            sSQL = sSQL & rsRighe!IDRV_POcaricoMercetesta & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & fnNotNullN(rsRighe!IDArticolo) & ", "
            sSQL = sSQL & fnNotNullN(rsRighe!IDImballo) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!Qta_UM) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!Colli) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!PesoLordo) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!Tara) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!PesoNetto) & ", "
            sSQL = sSQL & fnNormNumber(rsRighe!Pezzi) & ", "
            If TIPO_QUANTITA = 3 Then
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
            Else
                sSQL = sSQL & fnNormNumber(0) & ", "
            End If
            sSQL = sSQL & Link_CategoriaMerceologica & ", "
            sSQL = sSQL & fnNotNullN(rsRighe!IDAnagrafica) & ", "
            sSQL = sSQL & fnNormString(rsRighe!Anagrafica) & ", "
            sSQL = sSQL & fnNormString(rsRighe!Nome) & ", "
            sSQL = sSQL & fnNotNullN(rsRighe!NumeroDocumento) & ", "
            sSQL = sSQL & fnNormDate(rsRighe!DataDocumento) & ")"
            
            Cn.Execute sSQL
            
    rsRighe.MoveNext
        
    Wend
    
End If
rsRighe.Close
Set rsRighe = Nothing



End Sub
Private Sub GET_SCARTI(IDSocio As Long, IDConferimentoRiga As Long)
Dim sSQL As String
Dim rsScarti As DmtOleDbLib.adoResultset
Dim Link_CategoriaMerceologica As Long

sSQL = "SELECT RV_POLavorazione.IDRV_POLavorazione, RV_POLavorazione.IDRV_POCaricoMerceRighe, RV_POLavorazione.IDTipoLavorazione, "
sSQL = sSQL & "RV_POLavorazione.IDArticolo, RV_POLavorazione.CodiceArticolo, RV_POLavorazione.Articolo, RV_POLavorazione.Colli,"
sSQL = sSQL & "RV_POLavorazione.PesoLordo, RV_POLavorazione.PesoNetto, RV_POLavorazione.Tara, RV_POLavorazione.Pezzi, RV_POLavorazione.Qta_UM,"
sSQL = sSQL & "RV_POLavorazione.IDImballoVendita, RV_POLavorazione.CodiceImballoVendita, RV_POLavorazione.ImballoVendita, "
sSQL = sSQL & "RV_POLavorazione.IDRV_POCalibro, RV_POLavorazione.IDRV_POTipoCategoria, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceTesta.DataDocumento , RV_POCaricoMerceTesta.IDAnagrafica "
sSQL = sSQL & "FROM RV_POLavorazione INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POLavorazione.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE RV_POLavorazione.IDRV_POCaricoMerceRighe = " & IDConferimentoRiga

Set rsScarti = Cn.OpenResultset(sSQL)
While Not rsScarti.EOF
    If ARTICOLI_DI_QUAD = False Then
        TrattenutePerLavorazione = 0
        TrattenuteGenerali = 0
        TrattenuteTotali = 0

        TrattenuteArticolo IDSocio, rsScarti!IDArticolo, 0, rsScarti!IDTipoLavorazione, True, fnNotNullN(rsScarti!Qta_UM), 0
        
        Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(rsScarti!IDArticolo)
        
        sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
        sSQL = sSQL & "TipoRiga, IDRV_POCaricoMerceRighe, IDArticolo, CodiceArticolo, Articolo, Quantita, "
        sSQL = sSQL & "DataConferimento, IDSocio, IDValoreOggettoDettaglio, IDRV_POPeriodoLiquidazione, "
        sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, "
        sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
        sSQL = sSQL & "IDImballo, CodiceImballo, Imballo) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNormNumber(2) & ", "
        sSQL = sSQL & IDConferimentoRiga & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDArticolo) & ", "
        sSQL = sSQL & fnNormString(rsScarti!CodiceArticolo) & ", "
        sSQL = sSQL & fnNormString(rsScarti!Articolo) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!Qta_UM) & ", "
        sSQL = sSQL & fnNormDate(rsScarti!DataDocumento) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!IDAnagrafica) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!IDRV_POLavorazione) & ", "
        sSQL = sSQL & LINK_PERIODO & ", "
        sSQL = sSQL & Link_CategoriaMerceologica & ", "
        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!Colli) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!PesoLordo) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!PesoNetto) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!Tara) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!Pezzi) & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDTipoLavorazione) & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDRV_POCalibro) & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDRV_POTipoCategoria) & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDImballoVendita) & ", "
        sSQL = sSQL & fnNormString(rsScarti!CodiceImballoVendita) & ", "
        sSQL = sSQL & fnNormString(rsScarti!ImballoVendita) & ")"
        
        
        Cn.Execute sSQL
    End If

rsScarti.MoveNext
Wend


rsScarti.CloseResultset
Set rsScarti = Nothing

End Sub

Private Sub GET_VENDITA_DDT(IDConferimentoRighe As Long, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim PrezzoMedio As Double
Dim Link_CategoriaMerceologica As Long
Dim PrezzoDaReg As Double
Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_progresso As Double




'DOCUMENTO DI TRASPORTO
sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Iva.Iva AS IvaVendita, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0004.Link_Art_IVA "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " WHERE ValoriOggettoDettaglio0004.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
Else
    sSQL = sSQL & " WHERE ValoriOggettoPerTipo0002.Doc_Data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0002.Doc_Data<=" & fnNormDate(DATA_FINE)
End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe=" & IDConferimentoRighe
sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_Data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento))
            End If
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            TrattenutePerLavorazione = 0
            TrattenuteGenerali = 0
            TrattenuteTotali = 0
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali) "
                sSQL = sSQL & "VALUES ("
                If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                    sSQL = sSQL & 1 & ", "
                Else
                    sSQL = sSQL & 4 & ", "
                End If
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_descrizione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_numero_colli) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("D.D.T. n° " & rsFatt!Doc_Numero & " del " & rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_Numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Link_Art_IVA) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva))) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ")"
                
                Cn.Execute sSQL
                
                InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio)
                InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio)
            End If
        
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing




End Sub

Private Sub GET_VENDITA_FA(IDConferimentoRighe As Long, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double
Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_progresso As Double


'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Iva.Iva AS IvaVendita, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0001.Link_Art_IVA "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " WHERE ValoriOggettoDettaglio0001.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
Else
    sSQL = sSQL & " WHERE ValoriOggettoPerTipo0072.Doc_Data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.Doc_Data<=" & fnNormDate(DATA_FINE)

End If

sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe=" & IDConferimentoRighe
sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_Data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento))
            End If
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            TrattenutePerLavorazione = 0
            TrattenuteGenerali = 0
            TrattenuteTotali = 0
            
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali) "
                sSQL = sSQL & "VALUES ("
                If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                    sSQL = sSQL & 1 & ", "
                Else
                    sSQL = sSQL & 4 & ", "
                End If
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_descrizione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_numero_colli) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("F.A. n° " & rsFatt!Doc_Numero & " del " & rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_Numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Link_Art_IVA) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva))) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ")"
                
                Cn.Execute sSQL
                
                InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio)
                InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio)
            End If
        
        
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing

End Sub
Private Sub GET_VENDITA_SNF(IDConferimentoRighe As Long, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double
Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_progresso As Double


'SCONTRINO NON FISCALE
sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, Iva.Iva AS IvaVendita, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0034.Link_Art_IVA "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " WHERE ValoriOggettoDettaglio0034.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
Else
    sSQL = sSQL & " WHERE ValoriOggettoPerTipo0008.Doc_Data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0008.Doc_Data<=" & fnNormDate(DATA_FINE)

End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe=" & IDConferimentoRighe
sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_Data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento))
            End If
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            TrattenutePerLavorazione = 0
            TrattenuteGenerali = 0
            TrattenuteTotali = 0
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali) "
                sSQL = sSQL & "VALUES ("
                If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                    sSQL = sSQL & 1 & ", "
                Else
                    sSQL = sSQL & 4 & ", "
                End If
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_descrizione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_numero_colli) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("S.N.F. n° " & rsFatt!Doc_Numero & " del " & rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_Numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Link_Art_IVA) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva))) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ")"
                
                CnDMT.Execute sSQL
            End If
        
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing




End Sub


Private Sub TrattenuteArticolo(IDSocio As Long, IDArticolo As Long, IDCategoriaMerceologica As Long, IDTipoLavorazione As Long, DalConferimento As Boolean, Quantita As Double, PrezzoArticolo As Double)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim C As Integer
Dim R As Integer
Dim rsTratt As ADODB.Recordset
Dim sSQLTratt As String

sSQL = "SELECT RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta, RV_POCalcoloLiqTesta.IDFiliale, RV_POCalcoloLiqTesta.IDSocio, "
sSQL = sSQL & "RV_POCalcoloLiqTesta.AddebitoImballo, RV_POCalcoloLiqTesta.IDListinoImballo, RV_POTipoTrattenuta.IDRV_POTipoTrattenuta,"
sSQL = sSQL & "RV_POTipoTrattenuta.TipoTrattenuta, RV_POTipoTrattenuta.Tipo1, RV_POTipoTrattenuta.Tipo2, RV_POTipoTrattenuta.Tipo3,"
sSQL = sSQL & "RV_POTipoTrattenuta.Tipo4 "
sSQL = sSQL & "FROM RV_POCalcoloLiqTesta INNER JOIN "
sSQL = sSQL & "RV_POCalcoloLiqRighe ON RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta = RV_POCalcoloLiqRighe.IDRV_POCalcoloLiqTesta INNER JOIN "
sSQL = sSQL & "RV_POTipoTrattenuta ON RV_POCalcoloLiqRighe.IDRV_POTipoTrattenuta = RV_POTipoTrattenuta.IDRV_POTipoTrattenuta "
sSQL = sSQL & "WHERE IDSocio=" & IDSocio & " AND IDFiliale=" & TheApp.Branch


Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

If rs.EOF Then
    
    rs.Close
    
    
    sSQL = "SELECT RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta, RV_POCalcoloLiqTesta.IDFiliale, RV_POCalcoloLiqTesta.IDSocio, "
    sSQL = sSQL & "RV_POCalcoloLiqTesta.AddebitoImballo, RV_POCalcoloLiqTesta.IDListinoImballo, RV_POTipoTrattenuta.IDRV_POTipoTrattenuta,"
    sSQL = sSQL & "RV_POTipoTrattenuta.TipoTrattenuta, RV_POTipoTrattenuta.Tipo1, RV_POTipoTrattenuta.Tipo2, RV_POTipoTrattenuta.Tipo3,"
    sSQL = sSQL & "RV_POTipoTrattenuta.Tipo4 "
    sSQL = sSQL & "FROM RV_POCalcoloLiqTesta INNER JOIN "
    sSQL = sSQL & "RV_POCalcoloLiqRighe ON RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta = RV_POCalcoloLiqRighe.IDRV_POCalcoloLiqTesta INNER JOIN "
    sSQL = sSQL & "RV_POTipoTrattenuta ON RV_POCalcoloLiqRighe.IDRV_POTipoTrattenuta = RV_POTipoTrattenuta.IDRV_POTipoTrattenuta "
    sSQL = sSQL & "WHERE IDSocio=0 AND IDFiliale = " & TheApp.Branch
    
    rs.Open sSQL, Cn.InternalConnection
End If

If fnNotNullN(rs!AddebitoImballo) = 1 Then
    Link_AddebitoImballo = 1
Else
    Link_AddebitoImballo = 0
End If

Link_ListinoImballo = fnNotNullN(rs!IDListinoImballo)

TrattenutePerLavorazione = 0
TrattenuteGenerali = 0
TrattenuteTotali = 0

While Not rs.EOF
        sSQLTratt = GetSQL(rs!Tipo1, rs!Tipo2, rs!Tipo3, rs!Tipo4, IDSocio, IDCategoriaMerceologica, IDArticolo, IDTipoLavorazione)
        If sSQLTratt <> "" Then
            Set rsTratt = New ADODB.Recordset
            rsTratt.Open sSQLTratt, Cn.InternalConnection
                    
            If rsTratt.EOF = False Then
                If DalConferimento = True Then
                    If rs!Tipo4 = 1 Then
                        'A valore
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                    Else
                        'A valore
                        TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                    End If
                Else
                    If rs!Tipo4 = 1 Then
                        'A valore
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                        'A percentuale
                        TrattenutePerLavorazione = TrattenutePerLavorazione + ((PrezzoUnitario / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                        TrattenutePerLavorazione = TrattenutePerLavorazione + ((PrezzoUnitario / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        
                    Else
                        If TIPO_QUANTITA = 4 Then
                            'A valore
                            TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                            TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                            'A percentuale
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        Else
                            'A percentuale
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        End If
                    End If
                End If
            End If
            
            rsTratt.Close
            Set rsTratt = Nothing
    End If
    
    TrattenuteTotali = TrattenuteGenerali + TrattenutePerLavorazione
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub
Private Function GetSQL(Tipo1 As Integer, Tipo2 As Integer, Tipo3 As Integer, Tipo4 As Integer, ValueTipo1 As Long, ValueTipo2 As Long, ValueTipo3 As Long, ValueTipo4 As Long) As String
Dim sSQL As String
GetSQL = "SELECT * FROM RV_POTrattenutaPerLiquidazione WHERE "
GetSQL = GetSQL & "IDAzienda=" & TheApp.IDFirm & " AND "
GetSQL = GetSQL & "IDFiliale=" & TheApp.Branch

sSQL = ""
If Tipo1 = 1 Then
    If ValueTipo1 > 0 Then
        sSQL = sSQL & " AND IDSocio=" & ValueTipo1
    Else
        sSQL = sSQL
    End If
End If
If Tipo2 = 1 Then
    If ValueTipo2 > 0 Then
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & ValueTipo2
    Else
        sSQL = sSQL
    End If
End If
If Tipo3 = 1 Then
    If ValueTipo3 > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & ValueTipo3
    Else
        sSQL = sSQL
    End If
End If
If Tipo4 = 1 Then
    If ValueTipo4 > 0 Then
        sSQL = sSQL & " AND IDTipoLavorazione=" & ValueTipo4
    Else
        sSQL = sSQL
    End If
End If

If sSQL <> "" Then
    GetSQL = GetSQL & sSQL
Else
    GetSQL = ""
End If
End Function
Private Function GetPrezzoUnitarioPeriodo(IDArticolo As Long, DataRiferimento As String) As Double
Dim sSQL As String
Dim rsAvr As ADODB.Recordset
Dim TotaleQuantita As Double
Dim TotaleVendita As Double

TotaleQuantita = 0
TotaleVendita = 0

'DA DOCUMENTO DI TRASPORTO
sSQL = "SELECT RV_POImportoLiq, Art_quantita_totale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "
sSQL = sSQL & "WHERE Link_Art_Articolo = " & IDArticolo
If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
    sSQL = sSQL & " AND Doc_Data=" & fnNormDate(DataRiferimento)
Else
    sSQL = sSQL & " AND RV_PODataConferimento=" & fnNormDate(DataRiferimento)
End If


Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, Cn.InternalConnection

While Not rsAvr.EOF
    If fnNotNullN(rsAvr!RV_POImportoLiq) > 0 Then
        TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!Art_quantita_totale)
        TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!Art_quantita_totale) * fnNotNullN(rsAvr!RV_POImportoLiq))
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

'DA FATTURA ACCOMPAGNATORIA
sSQL = "SELECT RV_POImportoLiq, Art_quantita_totale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
sSQL = sSQL & "WHERE Link_Art_Articolo = " & IDArticolo
If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
    sSQL = sSQL & " AND Doc_Data=" & fnNormDate(DataRiferimento)
Else
    sSQL = sSQL & " AND RV_PODataConferimento=" & fnNormDate(DataRiferimento)
End If


Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, Cn.InternalConnection

While Not rsAvr.EOF
    If fnNotNullN(rsAvr!RV_POImportoLiq) > 0 Then
        TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!Art_quantita_totale)
        TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!Art_quantita_totale) * fnNotNullN(rsAvr!RV_POImportoLiq))
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing


'DA SCONTRINO NON FISCALE
sSQL = "SELECT RV_POImportoLiq, Art_quantita_totale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "
sSQL = sSQL & "WHERE Link_Art_Articolo = " & IDArticolo
If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
    sSQL = sSQL & " AND Doc_Data=" & fnNormDate(DataRiferimento)
Else
    sSQL = sSQL & " AND RV_PODataConferimento=" & fnNormDate(DataRiferimento)
End If

Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, Cn.InternalConnection

While Not rsAvr.EOF
    If fnNotNullN(rsAvr!RV_POImportoLiq) > 0 Then
        TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!Art_quantita_totale)
        TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!Art_quantita_totale) * fnNotNullN(rsAvr!RV_POImportoLiq))
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

'NOTA DI CREDITO
sSQL = "SELECT RV_POImportoLiq, Art_quantita_totale, RV_POQuantitaOrigine "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto "
sSQL = sSQL & "WHERE Link_Art_Articolo = " & IDArticolo
If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
    sSQL = sSQL & " AND Doc_Data=" & fnNormDate(DataRiferimento)
Else
    sSQL = sSQL & " AND RV_PODataConferimento=" & fnNormDate(DataRiferimento)
End If


Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, Cn.InternalConnection

While Not rsAvr.EOF
    If fnNotNullN(rsAvr!RV_POImportoLiq) > 0 Then
        If fnNotNullN(rsAvr!Art_quantita_totale) - fnNotNullN(rsAvr!RV_POQuantitaOrigine) = 0 Then
            TotaleQuantita = TotaleQuantita
        Else
            TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!Art_quantita_totale)
        End If
            TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!Art_quantita_totale) * fnNotNullN(rsAvr!RV_POImportoLiq))
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

If TotaleQuantita > 0 Then
    GetPrezzoUnitarioPeriodo = TotaleVendita / TotaleQuantita
End If


End Function
Private Function GetPrezzoImballo(IDImballo As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsList As ADODB.Recordset

sSQL = "SELECT Articolo.IDArticolo, RV_POTipoImballo.Rendere, Articolo.RV_POImballoPerAddebito "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoImballo ON Articolo.RV_POIDTipoImballo = RV_POTipoImballo.IDRV_POTipoImballo "
sSQL = sSQL & "WHERE IDArticolo=" & IDImballo

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GetPrezzoImballo = 0
Else
    If rs!Rendere = True Then
        If rs!RV_POImballoPerAddebito = 1 Then
            sSQL = "SELECT PrezzoNettoIva FROM ListinoPerArticolo "
            sSQL = sSQL & "WHERE IDListino=" & Link_ListinoImballo & " AND "
            sSQL = sSQL & "IDArticolo=" & IDImballo
            
            Set rsList = New ADODB.Recordset
            rsList.Open sSQL, CnDMT.InternalConnection
            
            If rsList.EOF Then
                GetPrezzoImballo = 0
            Else
                GetPrezzoImballo = fnNotNullN(rsList!PrezzoNettoIva)
            End If
            
            rsList.Close
            Set rsList = Nothing
        Else
            GetPrezzoImballo = 0
        End If
    Else
        GetPrezzoImballo = 0
    End If
End If

rs.Close
Set rs = Nothing
End Function
Private Function ParametroTipoQuadratura() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoQuadratura FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & VarIDFiliale & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroTipoQuadratura = fnNotNullN(rs!IDTipoQuadratura)
Else
    ParametroTipoQuadratura = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CATEGORIA_MERCEOLOGICA(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaMerceologica FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CATEGORIA_MERCEOLOGICA = 0
Else
    GET_CATEGORIA_MERCEOLOGICA = fnNotNullN(rs!IDCategoriaMerceologica)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub InserisciNotaCredito(ValoriOggettoDettaglio As Long)
'NOTA DI CREDITO
sSQL = "SELECT ValoriOggettoDettaglio0016.*, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero, Iva.Iva AS IvaVendita, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0016.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0016.Link_Art_IVA "


sSQL = sSQL & "WHERE ValoriOggettoDettaglio0016.RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio

Set rsFatt = Cn.OpenResultset(sSQL)

While Not rsFatt.EOF
    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
        
            PrezzoMedio = 0

            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & 3 & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_descrizione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_numero_colli) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNormNumber(-rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("N.C. n° " & rsFatt!Doc_Numero & " del " & rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_Numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Link_Art_IVA) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva))) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * fnNotNullN(-rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(-rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (-rsFatt!Art_quantita_totale)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (-rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (-rsFatt!Art_quantita_totale)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                If (rsFatt!Art_quantita_totale - rsFatt!RV_POQuantitaOrigine) = 0 Then
                    sSQL = sSQL & fnNormNumber(0) & ")"
                Else
                    sSQL = sSQL & fnNormNumber(-rsFatt!Art_quantita_totale) & ")"
                End If
            

        
        
        Cn.Execute sSQL
    End If
    
        
rsFatt.MoveNext
Wend

rsFatt.CloseResultset
Set rsFatt = Nothing
End Sub
Private Sub InserisciNotaDebito(ValoriOggettoDettaglio As Long)
'NOTA DI DEBITO
sSQL = "SELECT ValoriOggettoDettaglio0007.*, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero, Iva.Iva AS IvaVendita, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0007 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0007.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0007.Link_Art_IVA "


sSQL = sSQL & "WHERE ValoriOggettoDettaglio0007.RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio

Set rsFatt = Cn.OpenResultset(sSQL)

While Not rsFatt.EOF
    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
        
            PrezzoMedio = 0

            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & 3 & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Art_descrizione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_numero_colli) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("N.D. n° " & rsFatt!Doc_Numero & " del " & rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_Data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_Numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Link_Art_IVA) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaVendita) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva))) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (rsFatt!Art_quantita_totale)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!Art_aliquota_Iva)) * (rsFatt!Art_quantita_totale)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(-rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                sSQL = sSQL & fnNormNumber(0) & ", "
                If (rsFatt!Art_quantita_totale - rsFatt!RV_POQuantitaOrigine) = 0 Then
                    sSQL = sSQL & fnNormNumber(0) & ")"
                Else
                    sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ")"
                End If
            

        
        
        Cn.Execute sSQL
    End If
    
        
rsFatt.MoveNext
Wend

rsFatt.CloseResultset
Set rsFatt = Nothing
End Sub


Private Function GET_DATA_MINIMA_CONFERIMENTO(Data_Inizio_Vendita As String, data_fine_vendita As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Data_Minima As String

FrmNuovoPeriodo.List1.AddItem "CALCOLO DATA INIZIO CONFERIMENTO"
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1

sSQL = "SELECT MIN(RV_POTMPLiquidazioneVendita.DataConferimento) AS Data_Minima_Conf "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneVendita "

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Data_Minima = ""
Else
    Data_Minima = fnNotNull(rs!Data_Minima_Conf)
End If

rs.CloseResultset
Set rs = Nothing


If Len(Data_Minima) = 0 Then
    GET_DATA_MINIMA_CONFERIMENTO = Data_Inizio_Vendita
Else
    GET_DATA_MINIMA_CONFERIMENTO = Data_Minima
End If

FrmNuovoPeriodo.List1.AddItem "DATA INIZIO CONFERIMENTO: " & GET_DATA_MINIMA_CONFERIMENTO
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1

End Function

Private Function GET_DATA_MASSIMA_CONFERIMENTO(Data_Inizio_Vendita As String, data_fine_vendita As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Data_Massima As String

FrmNuovoPeriodo.List1.AddItem "CALCOLO DATA FINE CONFERIMENTO"
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1

sSQL = "SELECT MAX(RV_POTMPLiquidazioneVendita.DataConferimento) AS Data_Massima_Conf "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneVendita "

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Data_Massima = ""
Else
    Data_Massima = fnNotNull(rs!Data_Massima_Conf)
End If

rs.CloseResultset
Set rs = Nothing


If Len(Data_Massima) = 0 Then
    GET_DATA_MASSIMA_CONFERIMENTO = data_fine_vendita
Else
    GET_DATA_MASSIMA_CONFERIMENTO = Data_Massima
End If

FrmNuovoPeriodo.List1.AddItem "DATA FINE CONFERIMENTO: " & GET_DATA_MASSIMA_CONFERIMENTO
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1


End Function

