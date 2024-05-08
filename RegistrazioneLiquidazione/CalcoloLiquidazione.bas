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

Private TrattValGen1 As Double
Private TrattValGen2 As Double
Private TrattPercGen1 As Double
Private TrattPercGen2 As Double

Private TrattValLav1 As Double
Private TrattValLav2 As Double
Private TrattPercLav1 As Double
Private TrattPercLav2 As Double

Private TrattValPreLiq1 As Double
Private TrattValPreLiq2 As Double

Private TotaleRighe As Long

Private rsConfQtaAbb As ADODB.Recordset
Private rsConfQtaAbbSomma As ADODB.Recordset

Public Sub EsecuzioneElaborazione(prg As ProgressBar)
Dim Data_Minima_Conf As String
Dim Data_Massima_Conf As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneRigheConf"
CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneLavorazione"
CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneVendita"
CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneTrattConf"
CnDMT.Execute "DELETE FROM RV_POTMPLiquidazionePrezzoMedio"
CnDMT.Execute "DELETE FROM RV_POTMPLiquidazionePrezzoMedioRighe"


If TIPO_LIQUIDAZIONE = 1 Then
    If RICALCOLA_VALORI_LIQ = 1 Then
        If LINK_LISTINO = 0 Then
            fnCalcolaPrezzoUnitario prg, TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
        End If
    End If
    
    If ATTIVA_CALCOLO_QTA_DA_ABB = 1 Then
        AVVIA_CALCOLO_QTA_DA_ABBATTERE TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
    End If
    
    If (LIQUIDA_FORNITORE = 0) Then
        ElaborazioneConferimento prg, DATA_INIZIO, DATA_FINE
        
        GET_VENDITA_DDT prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
        GET_VENDITA_FA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
        GET_VENDITA_SNF prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
        GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "DOCUMENTI DI TRASPORTO", "RV_POIELiquidazioneArticoliMixDDT"
        GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "FATTURE ACCOMPAGNATORIE", "RV_POIELiquidazioneArticoliMixFA"
        GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "CORRISPETTIVI", "RV_POIELiquidazioneArticoliMixSNF"

    Else
        If LINK_SOCIO > 0 Then
            ElaborazioneConferimento prg, DATA_INIZIO, DATA_FINE
            
            GET_VENDITA_DDT prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
            GET_VENDITA_FA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
            GET_VENDITA_SNF prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
            GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "DOCUMENTI DI TRASPORTO", "RV_POIELiquidazioneArticoliMixDDT"
            GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "FATTURE ACCOMPAGNATORIE", "RV_POIELiquidazioneArticoliMixFA"
            GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "CORRISPETTIVI", "RV_POIELiquidazioneArticoliMixSNF"
        Else
            sSQL = "SELECT IDAnagrafica, IDTipoDocumentoCoop "
            sSQL = sSQL & "FROM RV_POCaricoMerceTesta "
            sSQL = sSQL & " WHERE DataDocumento>=" & fnNormDate(DATA_INIZIO)
            sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
            sSQL = sSQL & " GROUP BY IDAnagrafica, IDTipoDocumentoCoop "
            sSQL = sSQL & " HAVING (IDTipoDocumentoCoop = 2)"
            
            Set rs = CnDMT.OpenResultset(sSQL)
            
            While Not rs.EOF
                LINK_SOCIO = fnNotNullN(rs!IDAnagrafica)
                
                ElaborazioneConferimento prg, DATA_INIZIO, DATA_FINE
                
                GET_VENDITA_DDT prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
                GET_VENDITA_FA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
                GET_VENDITA_SNF prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
                GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "DOCUMENTI DI TRASPORTO", "RV_POIELiquidazioneArticoliMixDDT"
                GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "FATTURE ACCOMPAGNATORIE", "RV_POIELiquidazioneArticoliMixFA"
                GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "CORRISPETTIVI", "RV_POIELiquidazioneArticoliMixSNF"
                
            rs.MoveNext
            Wend
        End If
    End If
End If

If TIPO_LIQUIDAZIONE = 2 Then
    If RICALCOLA_VALORI_LIQ = 1 Then
        If LINK_LISTINO = 0 Then
            fnCalcolaPrezzoUnitario prg, TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
        End If
    End If
    
    If ATTIVA_CALCOLO_QTA_DA_ABB = 1 Then
        AVVIA_CALCOLO_QTA_DA_ABBATTERE TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
    End If
    
    FrmNuovoPeriodo.List1.Clear
    
    GET_VENDITA_DDT prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_FA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_SNF prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "DOCUMENTI DI TRASPORTO", "RV_POIELiquidazioneArticoliMixDDT"
    GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "FATTURE ACCOMPAGNATORIE", "RV_POIELiquidazioneArticoliMixFA"
    GET_VENDITA_MIX prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE, "CORRISPETTIVI", "RV_POIELiquidazioneArticoliMixSNF"
    GET_SCARTI_VENDITA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    
    Data_Minima_Conf = GET_DATA_MINIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    Data_Massima_Conf = GET_DATA_MASSIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    ElaborazioneConferimento prg, Data_Minima_Conf, Data_Massima_Conf

End If
If TIPO_LIQUIDAZIONE = 3 Then
    frmVisDoc.Show vbModal
    
    If CONFERMA_SEL_DOCUMENTI = 0 Then Exit Sub
    
    If RICALCOLA_VALORI_LIQ = 1 Then
        fnCalcolaPrezzoUnitario prg, TIPO_LIQUIDAZIONE, DATA_INIZIO, DATA_FINE
    End If

    GET_VENDITA_DDT_SU_RICHIESTA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_FA_SU_RICHIESTA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE
    GET_VENDITA_SNF_SU_RICHIESTA prg, DATA_INIZIO, DATA_FINE, TIPO_LIQUIDAZIONE

    Data_Minima_Conf = GET_DATA_MINIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    Data_Massima_Conf = GET_DATA_MASSIMA_CONFERIMENTO(DATA_INIZIO, DATA_FINE)
    ElaborazioneConferimento prg, Data_Minima_Conf, Data_Massima_Conf
    
End If

ElborazioneTMPLiquidazione prg

End Sub
Private Function GET_FLAG_NON_LIQUIDARE(IDRigaConferimentoMerce As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo.RV_PONonLiquidare, RV_POCaricoMerceRighe.IDArticolo "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "Articolo ON RV_POCaricoMerceRighe.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe=" & IDRigaConferimentoMerce

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_FLAG_NON_LIQUIDARE = False
Else
    GET_FLAG_NON_LIQUIDARE = fnNotNullN(rs!RV_PONonLiquidare)
    If LINK_ARTICOLO > 0 Then
        If LINK_ARTICOLO = fnNotNullN(rs!IDArticolo) Then
            GET_FLAG_NON_LIQUIDARE = False
        Else
            GET_FLAG_NON_LIQUIDARE = True
        End If
    End If
    
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ElaborazioneConferimento(prg As ProgressBar, DataInizio As String, DataFine As String)
Dim sSQL As String
Dim rsSocio As DmtOleDbLib.adoResultset
Dim rsRighe As ADODB.Recordset
Dim TrattenutaConferimento As Double
Dim LINK_SOCIO_LOCAL As Long
Dim Link_CategoriaMerceologica As Long
Dim Unita_Progresso As Double
Dim QuantitaConferita As Double
Dim ColliAbbattuti As Double
Dim PesoLordoAbbattuto As Double
Dim TaraAbbattuto As Double
Dim PesoNettoAbbattuto As Double
Dim PezziAbbattuto As Double

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE CONFERIMENTO"

'sSQL = "SELECT * "
'sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
'sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
'sSQL = sSQL & " WHERE RV_POCaricoMerceTesta.IDAzienda=" & TheApp.IDFirm
'sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DataInizio)
'sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DataFine)

sSQL = "SELECT RV_POCaricoMerceRighe.*, RV_POCaricoMerceTesta.*, Articolo.RV_POIDCategoriaLiquidazione, Articolo.IDCategoriaMerceologica, Articolo.RV_POPercentualeAbbattimentoLiquidazione "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON RV_POCaricoMerceRighe.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & " WHERE RV_POCaricoMerceTesta.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DataInizio)
sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DataFine)
sSQL = sSQL & " AND RV_POCaricoMerceTesta.PreConferimento=0"

If LIQ_CONGUAGLIO = 0 Then 'SE NON SONO CONGUAGLI ALLORA SI GESTISCE LO STATO DELLA RIGA DI CONFERIMENTO
    If LINK_TIPO_LIQ_CONF = 2 Then
        sSQL = sSQL & " AND IDRV_POTipoConfLiquidazione=1"
    End If
End If

If LIQUIDA_FORNITORE = 0 Then
    If TIPO_LIQUIDAZIONE < 3 Then
        sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDTipoDocumentoCoop=1"
    End If
End If

If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDAnagrafica=" & LINK_SOCIO
End If
If LINK_ARTICOLO > 0 Then
    sSQL = sSQL & " AND RV_POCaricoMerceRighe.IDArticolo=" & LINK_ARTICOLO
End If
'If TIPO_LIQUIDAZIONE = 1 Then
If LINK_CAT_MERCE > 0 Then
    sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
End If
'End If

sSQL = sSQL & " ORDER BY RV_POCaricoMerceTesta.IDAnagrafica "

Set rsRighe = New ADODB.Recordset
rsRighe.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsRighe.EOF = False Then
    prg.Value = 0
    prg.Max = 1000
    
    Unita_Progresso = prg.Max / rsRighe.RecordCount
    While Not rsRighe.EOF
        
        FrmNuovoPeriodo.lblInfoStatus.Caption = "Conferimento del socio " & rsRighe!Anagrafica & " n° " & fnNotNullN(rsRighe!NumeroDocumento) & " del " & fnNotNull(rsRighe!DataDocumento) & " - Cod. Art.: " & fnNotNull(rsRighe!CodiceArticolo) & " (" & fnNotNull(rsRighe!Articolo) & ")"
        FrmNuovoPeriodo.List1.AddItem "- Conferimento del socio " & rsRighe!Anagrafica & " n° " & fnNotNullN(rsRighe!NumeroDocumento) & " del " & fnNotNull(rsRighe!DataDocumento) & " - Cod. Art.: " & fnNotNull(rsRighe!CodiceArticolo) & " (" & fnNotNullN(rsRighe!Qta_UM) & ")"
        FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
        
        If GET_CONTROLLO_CONF_VEND(fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe)) = True Then
        
            If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe)) = False Then
                
                'If LINK_CAT_MERCE = 0 Then
                '    Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsRighe!IDArticolo))
                'Else
                '    Link_CategoriaMerceologica = LINK_CAT_MERCE
                'End If
                QuantitaConferita = fnNotNullN(rsRighe!Qta_UM) - ((fnNotNullN(rsRighe!Qta_UM) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                ColliAbbattuti = fnNotNullN(rsRighe!Colli) - ((fnNotNullN(rsRighe!Colli) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                PesoLordoAbbattuto = fnNotNullN(rsRighe!PesoLordo) - ((fnNotNullN(rsRighe!PesoLordo) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                TaraAbbattuto = fnNotNullN(rsRighe!Tara) - ((fnNotNullN(rsRighe!Tara) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                PesoNettoAbbattuto = fnNotNullN(rsRighe!PesoNetto) - ((fnNotNullN(rsRighe!PesoNetto) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                PezziAbbattuto = fnNotNullN(rsRighe!Pezzi) - ((fnNotNullN(rsRighe!Pezzi) / 100) * fnNotNullN(rsRighe!RV_POPercentualeAbbattimentoLiquidazione))
                
                Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsRighe!IDCategoriaMerceologica))
                
                If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe)) = False Then
                    If TIPO_LIQUIDAZIONE <> 2 Then
                        GET_SCARTI rsRighe!IDAnagrafica, rsRighe!IDRV_POCaricoMerceRighe
                    End If
                Else
                    'If TIPO_LIQUIDAZIONE = 1 Then
                        GET_CAMPIONATURA fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe)
                    'End If
                End If
                
                TrattenutePerLavorazione = 0
                TrattenuteGenerali = 0
                TrattenuteTotali = 0
                TrattenutaConferimento = 0
                
                If ((rsRighe!DataDocumento >= DATA_INIZIO) And (rsRighe!DataDocumento <= DATA_FINE)) Then
                    If TIPO_QUANTITA = 3 Then
                        TrattenuteArticolo fnNotNullN(rsRighe!IDAnagrafica), rsRighe!IDArticolo, Link_CategoriaMerceologica, 0, True, QuantitaConferita, 0, fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe), 0, 0, fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe), 0
                    Else
                        TrattenutaConferimento = GET_TRATTENUTA_ARTICOLO_CONFERITO(fnNotNullN(rsRighe!IDArticolo), ColliAbbattuti, PesoLordoAbbattuto, TaraAbbattuto, PesoNettoAbbattuto, PezziAbbattuto, fnNotNullN(rsRighe!IDUnitaDiMisura), fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe), fnNotNullN(rsRighe!IDAnagrafica))
                    End If
                End If
                
                sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheConf ("
                sSQL = sSQL & "IDRV_POCaricoMerceRighe, CodiceArticolo, Articolo, IDRV_POCaricoMerceTesta, IDRV_POPeriodoLiquidazione, IDArticolo,  "
                sSQL = sSQL & "IDImballo, Quantita, Colli, PesoLordo, Tara, PesoNetto, Pezzi, "
                sSQL = sSQL & "Trattenuta, IDCategoriaMerceologica, IDAnagrafica, Anagrafica, Nome, "
                sSQL = sSQL & "TotaleTrattenutaCOnferimento, NumeroDocumento, DataDocumento) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & rsRighe!IDRV_POCaricoMerceRighe & ", "
                sSQL = sSQL & fnNormString(rsRighe!CodiceArticolo) & ", "
                sSQL = sSQL & fnNormString(rsRighe!Articolo) & ", "
                sSQL = sSQL & rsRighe!IDRV_POcaricoMercetesta & ", "
                sSQL = sSQL & LINK_PERIODO & ", "
                sSQL = sSQL & fnNotNullN(rsRighe!IDArticolo) & ", "
                sSQL = sSQL & fnNotNullN(rsRighe!IDImballo) & ", "
                sSQL = sSQL & fnNormNumber(QuantitaConferita) & ", "
                sSQL = sSQL & fnNormNumber(ColliAbbattuti) & ", "
                sSQL = sSQL & fnNormNumber(PesoLordoAbbattuto) & ", "
                sSQL = sSQL & fnNormNumber(TaraAbbattuto) & ", "
                sSQL = sSQL & fnNormNumber(PesoNettoAbbattuto) & ", "
                sSQL = sSQL & fnNormNumber(PezziAbbattuto) & ", "
                If TIPO_QUANTITA = 3 Then
                    sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                Else
                    sSQL = sSQL & fnNormNumber(0) & ", "
                End If
                sSQL = sSQL & Link_CategoriaMerceologica & ", "
                sSQL = sSQL & fnNotNullN(rsRighe!IDAnagrafica) & ", "
                sSQL = sSQL & fnNormString(rsRighe!Anagrafica) & ", "
                sSQL = sSQL & fnNormString(rsRighe!Nome) & ","
                sSQL = sSQL & fnNormNumber(TrattenutaConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsRighe!NumeroDocumento) & ", "
                sSQL = sSQL & fnNormDate(rsRighe!DataDocumento) & ")"
                
                CnDMT.Execute sSQL
            End If
        End If

        
        If (Unita_Progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
            
        DoEvents
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
Dim Link_UM_Coop As Long
Dim QuantitaLiq As Double
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiqAbbattuta As Double

sSQL = "SELECT RV_POLavorazione.IDRV_POLavorazione, RV_POLavorazione.IDRV_POCaricoMerceRighe, RV_POLavorazione.IDTipoLavorazione, "
sSQL = sSQL & "RV_POLavorazione.IDArticolo, RV_POLavorazione.CodiceArticolo, RV_POLavorazione.Articolo, RV_POLavorazione.Colli, "
sSQL = sSQL & "RV_POLavorazione.PesoLordo, RV_POLavorazione.PesoNetto, RV_POLavorazione.Tara, RV_POLavorazione.Pezzi, RV_POLavorazione.Qta_UM, "
sSQL = sSQL & "RV_POLavorazione.IDImballoVendita, RV_POLavorazione.CodiceImballoVendita, RV_POLavorazione.ImballoVendita, "
sSQL = sSQL & "RV_POLavorazione.IDRV_POCalibro, RV_POLavorazione.IDRV_POTipoCategoria, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceTesta.DataDocumento , RV_POCaricoMerceTesta.IDAnagrafica, RV_POLavorazione.DataDocumento AS DataLavorazioneScarto "
sSQL = sSQL & "FROM RV_POLavorazione INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POLavorazione.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE RV_POLavorazione.IDRV_POCaricoMerceRighe = " & IDConferimentoRiga
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND RV_POLavorazione.DataDocumento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND RV_POLavorazione.DataDocumento<=" & fnNormDate(DATA_FINE)
End If


Set rsScarti = CnDMT.OpenResultset(sSQL)
While Not rsScarti.EOF
    If ARTICOLI_DI_QUAD = False Then
        TrattenutePerLavorazione = 0
        TrattenuteGenerali = 0
        TrattenuteTotali = 0
        TrattValGen1 = 0
        TrattValGen2 = 0
        TrattPercGen1 = 0
        TrattPercGen2 = 0
        
        TrattValLav1 = 0
        TrattValLav2 = 0
        TrattPercLav1 = 0
        TrattPercLav2 = 0
            
        If fnNotNullN(rsScarti!IDArticolo) > 0 Then
            AvviaLiquidazioneRiga = 1
            If NO_LIQ_VEND_UFF = 1 Then
                If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsScarti!IDRV_POLavorazione), 0, 0, 2) = True) Then
                AvviaLiquidazioneRiga = 0
                End If
            End If
'            If NO_RIP_SCARTI_IN_LIQ = 1 Then
'                AvviaLiquidazioneRiga = 0
'            End If
            If AvviaLiquidazioneRiga = 1 Then
                Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsScarti!IDArticolo))
                QuantitaLiq = fnNotNullN(rsScarti!Qta_UM)
                Link_UM_Coop = GET_LINK_UM_COOP(fnNotNullN(rsScarti!IDArticolo))
                Select Case Link_UM_Coop
                    Case 1
                        QuantitaLiq = fnNotNullN(rsScarti!Colli)
                    Case 2
                        QuantitaLiq = fnNotNullN(rsScarti!PesoLordo)
                    Case 3
                        QuantitaLiq = fnNotNullN(rsScarti!PesoNetto)
                    Case 4
                        QuantitaLiq = fnNotNullN(rsScarti!Tara)
                    Case 5
                        QuantitaLiq = fnNotNullN(rsScarti!Pezzi)
                End Select
                If LINK_TIPO_AUMENTO_PESO = GET_TIPO_SCARTO_ARTICOLO(fnNotNullN(rsScarti!IDArticolo)) Then
                    QuantitaLiq = QuantitaLiq * -1
                End If
                If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                    QuantitaLiqAbbattuta = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), fnNotNullN(rsScarti!IDRV_POLavorazione), 0, 0, 2)
                    If QuantitaLiqAbbattuta = 0 Then
                        QuantitaLiq = QuantitaLiq
                    Else
                        QuantitaLiq = QuantitaLiqAbbattuta
                    End If
                End If
                
                'TrattenuteArticolo IDSocio, fnNotNullN(rsScarti!IDArticolo), Link_CategoriaMerceologica, fnNotNullN(rsScarti!IDTipoLavorazione), False, fnNotNullN(rsScarti!Qta_UM), 0, fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), 0, 0, fnNotNullN(rsScarti!IDRV_POLavorazione), 0
                TrattenuteArticolo IDSocio, fnNotNullN(rsScarti!IDArticolo), Link_CategoriaMerceologica, fnNotNullN(rsScarti!IDTipoLavorazione), False, QuantitaLiq, 0, fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), 0, 0, fnNotNullN(rsScarti!IDRV_POLavorazione), 0
                            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDRV_POCaricoMerceRighe, IDArticolo, CodiceArticolo, Articolo, Quantita, "
                sSQL = sSQL & "DataConferimento, IDSocio, IDValoreOggettoDettaglio, IDRV_POPeriodoLiquidazione, "
                sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
                sSQL = sSQL & "IDImballo, CodiceImballo, Imballo, "
                sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2) "
    
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnNormNumber(2) & ", "
                sSQL = sSQL & IDConferimentoRiga & ", "
                sSQL = sSQL & fnNotNullN(rsScarti!IDArticolo) & ", "
                sSQL = sSQL & fnNormString(rsScarti!CodiceArticolo) & ", "
                sSQL = sSQL & fnNormString(rsScarti!Articolo) & ", "
                'sSQL = sSQL & fnNormNumber(rsScarti!Qta_UM) & ", "
                sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
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
                sSQL = sSQL & fnNormString(rsScarti!ImballoVendita) & ", "
                sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercLav2) & ")"
                
                CnDMT.Execute sSQL
            End If
        End If
    End If

rsScarti.MoveNext
Wend


rsScarti.CloseResultset
Set rsScarti = Nothing

End Sub
Private Sub GET_CAMPIONATURA(IDConferimentoRiga As Long)
Dim sSQL As String
Dim rsCamp As DmtOleDbLib.adoResultset
Dim Link_CategoriaMerceologica As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long

sSQL = "SELECT * FROM RV_POIECampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rsCamp = CnDMT.OpenResultset(sSQL)
While Not rsCamp.EOF
    
    
    If fnNotNullN(rsCamp!IDArticolo) > 0 Then
        
        If fnNotNullN(rsCamp!IDArticoloQuadratura) = 0 Then
            
'            If (IDConferimentoRiga = 47932) Then
'                MsgBox "STOP"
'            End If
            
            ImportoScontiPM = 0
            ImportoVarImballoPM = 0
            ImportoCommissioniPM = 0
            ImportoNettoIvaPM = 0
            Link_TMP_Prezzo_Medio = 0
            PrezzoDaReg = fnNotNullN(rsCamp!ImportoUnitario)
            Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsCamp!IDArticolo))
            
'            If (fnNotNullN(rsCamp!IDArticolo) = 928) Then
'                MsgBox "STOP"
'            End If
            
            If CALCOLA_PM_CAMP = 1 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsCamp!IDAnagraficaConferimento), fnNotNullN(rsCamp!IDArticolo), fnNotNull(rsCamp!DataDocumento), fnNotNull(rsCamp!DataCampionatura), fnNotNull(rsCamp!DataCampionatura), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
    
                If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                    If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                        PrezzoDaReg = fnNotNullN(rsCamp!ImportoUnitario)
                        ImportoScontiDaReg = 0
                        ImportoVarImballoDaReg = 0
                        ImportoCommissioniDaReg = 0
                        ImportoNettoIvaDaReg = fnNotNullN(rsCamp!ImportoUnitario)
                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Link_TMP_Prezzo_Medio = 0
                    Else
                        PrezzoDaReg = PrezzoMedio
                        ImportoScontiDaReg = ImportoScontiPM
                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                        ImportoCommissioniDaReg = ImportoCommissioniPM
                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                        
                        SALVA_PREZZO_MEDIO_IN_CAMPIONATURA fnNotNullN(rsCamp!IDRV_POCampionaturaRighe), PrezzoDaReg, fnNotNullN(rsCamp!AliquotaIva)
                        
                    End If
                Else
                    PrezzoDaReg = PrezzoMedio
                    ImportoScontiDaReg = ImportoScontiPM
                    'ImportoVarImballoDaReg = ImportoVarImballoPM
                    ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                    ImportoCommissioniDaReg = ImportoCommissioniPM
                    ImportoNettoIvaDaReg = ImportoNettoIvaPM
                    Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                    TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                    
                    SALVA_PREZZO_MEDIO_IN_CAMPIONATURA fnNotNullN(rsCamp!IDRV_POCampionaturaRighe), PrezzoDaReg, fnNotNullN(rsCamp!AliquotaIva)
    
                End If
    
    
            End If
            
            TrattenutePerLavorazione = 0
            TrattenuteGenerali = 0
            TrattenuteTotali = 0
            
            TrattValGen1 = 0
            TrattValGen2 = 0
            TrattPercGen1 = 0
            TrattPercGen2 = 0
            
            TrattValLav1 = 0
            TrattValLav2 = 0
            TrattPercLav1 = 0
            TrattPercLav2 = 0
            
            If CALCOLA_TRATT_CAMP = 1 Then
                'If fnNotNullN(rsCamp!Invenduto) = 0 Then
                    TrattenuteArticolo fnNotNullN(rsCamp!IDAnagraficaConferimento), fnNotNullN(rsCamp!IDArticolo), Link_CategoriaMerceologica, fnNotNullN(rsCamp!IDRV_POTipoLavorazioneConf), False, fnNotNullN(rsCamp!QuantitaDefinitiva), PrezzoDaReg, fnNotNullN(rsCamp!IDRV_POCaricoMerceRighe), 0, 0, 0, fnNotNullN(rsCamp!Invenduto)
                'End If
            End If
            
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
            sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, "
            sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
            sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2) "


            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & 1 & ", " 'TipoRiga
            sSQL = sSQL & fnNotNullN(rsCamp!IDArticolo) & ", " 'IDArticolo di vendita
            sSQL = sSQL & fnNormString(rsCamp!CodiceArticolo) & ", " 'Codice articolo di vendita
            sSQL = sSQL & fnNormString(rsCamp!Articolo) & ", " 'Descrizione articolo di vendita
            sSQL = sSQL & fnNotNullN(rsCamp!IDRV_POTipoLavorazioneConf) & ", " 'Tipo Lavorazione nella vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Tipo categoria nella vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Calibro nella vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Numero colli vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Peso lordo della vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'peso netto della vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Tara della vendita
            sSQL = sSQL & fnNormNumber(0) & ", " 'Pezzi della vendita
            sSQL = sSQL & fnNotNullN(rsCamp!IDRV_POCaricoMerceRighe) & ", " 'Link della riga conferita
            sSQL = sSQL & fnNormDate(rsCamp!DataConferimento) & ", " 'data conferimento
            sSQL = sSQL & fnNotNullN(rsCamp!IDAnagraficaConferimento) & ", " 'Link dell'anagrafica del socio
            sSQL = sSQL & LINK_PERIODO & ", " 'link  del periodo di liquidazione
            sSQL = sSQL & fnNormNumber(rsCamp!QuantitaDefinitiva) & ", " 'Quantità di liquidazione
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDOggetto di vendita
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDDella riga di campionatura
            sSQL = sSQL & fnNormString("Campionatura n° " & fnNotNullN(rsCamp!AnnoCampionatura) & "-" & fnNotNullN(rsCamp!NumeroCampionatura) & " del " & fnNotNull(rsCamp!DataCampionatura)) & ", " 'Descrizione del tipo oggetto
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDTipoOggetto
            sSQL = sSQL & fnNotNullN(0) & ", " 'Tipo oggetto variante
            sSQL = sSQL & fnNormDate(rsCamp!DataCampionatura) & ", " 'Data di vendita
            sSQL = sSQL & fnNormString(fnNotNullN(rsCamp!AnnoCampionatura) & "-" & fnNotNullN(rsCamp!NumeroCampionatura)) & ", " 'Numero di vendita
            
            sSQL = sSQL & fnNormNumber(rsCamp!ImportoUnitario) & ", " 'Importo unitario di liquidazione
            sSQL = sSQL & fnNormNumber(rsCamp!ImportoNettoRiga) & ", " 'Imponibile riga di liquidazione
            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", " 'Prezzo medio unitario di liquidazione
            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", " 'Imponibile medio della riga
            sSQL = sSQL & fnNormNumber(rsCamp!IDIva) & ", " 'IDIva dell'articolo venduto
            sSQL = sSQL & fnNormString(rsCamp!CodiceIva) & ", " 'Codice iva dell'articolo venduto
            sSQL = sSQL & fnNormNumber(rsCamp!AliquotaIva) & ", " 'Aliquota Iva dell'articolo venduto
            sSQL = sSQL & fnNormString(rsCamp!Iva) & ", " 'Descrizione IVA dell'articolo venduto
            sSQL = sSQL & fnNormNumber(rsCamp!IDIva) & ", " 'IDIva dell'articolo venduto a prezzo medio
            sSQL = sSQL & fnNormString(rsCamp!CodiceIva) & ", " 'Codice IVA dell'articolo venduto a prezzo medio
            sSQL = sSQL & fnNormNumber(rsCamp!AliquotaIva) & ", " 'Aliquota IVA dell'articolo venduto a prezzo medio
            sSQL = sSQL & fnNormString(rsCamp!Iva) & ", " 'Descrizione IVA dell'articolo venduto a prezzo medio
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsCamp!ImportoUnitario) / 100) * fnNotNullN(rsCamp!AliquotaIva))) & ", " 'Importo unitario dell'imposta sul prezzo di vendita di liquidazione
            sSQL = sSQL & fnNormNumber(((PrezzoMedio / 100) * fnNotNullN(rsCamp!AliquotaIva))) & "," ''Importo unitario dell'imposta sul prezzo medio di liquidazione
            sSQL = sSQL & fnNormNumber(rsCamp!ImportoImpostaRiga) & ", " 'Imposta totale sul venduto
            sSQL = sSQL & fnNormNumber(((PrezzoMedio / 100) * fnNotNullN(rsCamp!AliquotaIva)) * fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", " 'Imposta totale sul prezzo medio
            sSQL = sSQL & fnNormNumber(rsCamp!ImportoLordoRiga) & ", " 'Totale lordo sul venduto
            sSQL = sSQL & fnNormNumber((PrezzoMedio + ((PrezzoMedio / 100) * fnNotNullN(rsCamp!AliquotaIva))) * fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", " 'Imposta totale sul prezzo medio
            
            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", " 'Importo unitario di liquidazione da prendere in considerazione
            sSQL = sSQL & fnNormNumber(PrezzoDaReg * fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", " 'Imponibile riga di liquidazione
            sSQL = sSQL & fnNormNumber(((PrezzoDaReg / 100) * fnNotNullN(rsCamp!AliquotaIva)) * fnNotNullN(rsCamp!QuantitaDefinitiva)) & "," 'Imposta totale sul venduto
            sSQL = sSQL & fnNormNumber((PrezzoDaReg + ((PrezzoDaReg / 100) * fnNotNullN(rsCamp!AliquotaIva))) * fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", "  'Totale lordo sul venduto
                    
            sSQL = sSQL & Link_CategoriaMerceologica & ", " 'Link della categoria merceologica
            sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", " 'Trattenute per lavorazione
            sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", " 'Trattenute generali
            sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", " 'Trattenute totali
            sSQL = sSQL & fnNormNumber(rsCamp!QuantitaDefinitiva) & ", " 'Quantita per totali
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            'IMPORTO SCONTI
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
            'IMPORTO VARIAZIONE IMBALLO
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
            'IMPORTO COMMISSIONI
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
            'IMPORTO UNITARIO NETTO IVA
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsCamp!ImportoUnitario)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
            sSQL = sSQL & fnNotNullN(rsCamp!Invenduto) & ", "
            sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
            sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
            sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
            sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercLav2) & ")"
            
            CnDMT.Execute sSQL
        Else
            '''''SCARTI IN CAMPIONATURA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsCamp!IDArticoloQuadratura))
            
            TrattenutePerLavorazione = 0
            TrattenuteGenerali = 0
            TrattenuteTotali = 0
            
            TrattValGen1 = 0
            TrattValGen2 = 0
            TrattPercGen1 = 0
            TrattPercGen2 = 0
            
            TrattValLav1 = 0
            TrattValLav2 = 0
            TrattPercLav1 = 0
            TrattPercLav2 = 0

            If CALCOLA_TRATT_CAMP = 1 Then
                TrattenuteArticolo fnNotNullN(rsCamp!IDAnagraficaConferimento), fnNotNullN(rsCamp!IDArticoloQuadratura), Link_CategoriaMerceologica, fnNotNullN(rsCamp!IDRV_POTipoLavorazioneConf), False, fnNotNullN(rsCamp!QuantitaDefinitiva), 0, fnNotNullN(rsCamp!IDRV_POCaricoMerceRighe), 0, 0, 0, 0
            End If
                                    
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDRV_POCaricoMerceRighe, IDArticolo, CodiceArticolo, Articolo, Quantita, "
            sSQL = sSQL & "DataConferimento, IDSocio, IDValoreOggettoDettaglio, IDRV_POPeriodoLiquidazione, "
            sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
            sSQL = sSQL & "IDImballo, CodiceImballo, Imballo, "
            sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
            sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2) "


            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & fnNormNumber(2) & ", " 'TipoRiga
            sSQL = sSQL & IDConferimentoRiga & ", " 'IDRV_POCaricoMerceRighe
            sSQL = sSQL & fnNotNullN(rsCamp!IDArticoloQuadratura) & ", " 'IDArticolo
            sSQL = sSQL & fnNormString(rsCamp!CodiceArticoloQuadratura) & ", " 'CodiceArticolo
            sSQL = sSQL & fnNormString(rsCamp!ArticoloQuadratura) & ", " 'Articolo
            If fnNotNullN(rsCamp!QuantitaDefinitiva) <= 0 Then
                sSQL = sSQL & fnNormNumber(Abs(fnNotNullN(rsCamp!QuantitaDefinitiva))) & ", " 'Quantita di scarto o calo peso
            Else
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsCamp!QuantitaDefinitiva)) & ", " 'Quantita di aumeto peso
            End If
            sSQL = sSQL & fnNormDate(rsCamp!DataDocumento) & ", " 'DataConferimento
            sSQL = sSQL & fnNormNumber(rsCamp!IDAnagraficaConferimento) & ", " 'IDSocio
            sSQL = sSQL & fnNormNumber(rsCamp!IDRV_POCampionaturaRighe) & ", " 'IDValoreOggettoDettaglio
            sSQL = sSQL & LINK_PERIODO & ", " 'IDRV_POPeriodoLiquidazione
            sSQL = sSQL & Link_CategoriaMerceologica & ", " 'IDCategoriaMerceologica
            sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", " 'TrattenutePerLavorazione
            sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", " 'TrattenuteGenerali
            sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", " 'TrattenuteTotali
            sSQL = sSQL & fnNormNumber(0) & ", " 'Colli
            sSQL = sSQL & fnNormNumber(0) & ", " 'PesoLordo
            sSQL = sSQL & fnNormNumber(0) & ", " 'PesoNetto
            sSQL = sSQL & fnNormNumber(0) & ", " 'Tara
            sSQL = sSQL & fnNormNumber(0) & ", " 'Pezzi
            sSQL = sSQL & fnNotNullN(rsCamp!IDRV_POTipoLavorazioneConf) & ", " 'IDTipoLavorazione
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDCalibro
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDTipoCategoria
            sSQL = sSQL & fnNotNullN(0) & ", " 'IDImballo
            sSQL = sSQL & fnNormString("") & ", " 'CodiceImballo
            sSQL = sSQL & fnNormString("") & ", " 'Imballo
            sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
            sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
            sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
            sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
            sSQL = sSQL & fnNormNumber(TrattPercLav2) & ")"
            
            CnDMT.Execute sSQL
        End If 'CHIUSURA DELL'ARTICOLO DI QUADRATURA
    End If


rsCamp.MoveNext
Wend


rsCamp.CloseResultset
Set rsCamp = Nothing

End Sub

Private Sub GET_VENDITA_DDT(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Link_CategoriaMerceologica As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiq As Double

Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE DOCUMENTI DI TRASPORTO"
DoEvents

'DOCUMENTO DI TRASPORTO
sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Oggetto.Oggetto, "
sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0002.Doc_Prefisso, Articolo.RV_POIDCategoriaLiquidazione, ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0004.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0004.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0004.Link_Art_articolo "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If LINK_CAT_MERCE > 0 Then
    sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
End If

Select Case TipoLiquidazione
    Case 1
        sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
        sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
        If LINK_SOCIO > 0 Then
            sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POIDSocio=" & LINK_SOCIO
        End If
    Case 2
        sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO)
        sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE)
        If LINK_SOCIO > 0 Then
            sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POIDSocio=" & LINK_SOCIO
        End If
    Case 3
End Select

sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1 "
sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    prg.Value = 0
    prg.Max = 1000
    Unita_Progresso = prg.Max / rsFatt.RecordCount
    
    While Not rsFatt.EOF
        FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione " & fnNotNull(rsFatt!Oggetto) & " n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
        FrmNuovoPeriodo.List1.AddItem "- " & FrmNuovoPeriodo.lblInfoStatus.Caption
        FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
        DoEvents
        
        AvviaLiquidazioneRiga = 1
        If NO_LIQ_VEND_UFF = 1 Then
            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                AvviaLiquidazioneRiga = 0
            End If
        End If
        If fnNotNullN(rsFatt!RV_PORigaRiscontroPeso) = 1 Then
            Select Case fnNotNullN(rsFatt!RV_POIDTipoDocumentoCoop)
                Case 1
                    If RISCONTRO_PESO_SOCIO_VAL = 1 Then
                        AvviaLiquidazioneRiga = 0
                    End If
                Case 2
                    If RISCONTRO_PESO_FORN_VAL = 1 Then
                        AvviaLiquidazioneRiga = 0
                    End If
             End Select
        End If
        
        If AvviaLiquidazioneRiga = 1 Then
            If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                        ImportoScontiPM = 0
                        ImportoVarImballoPM = 0
                        ImportoCommissioniPM = 0
                        ImportoNettoIvaPM = 0
                        Link_TMP_Prezzo_Medio = 0
                        
                        If LINK_LISTINO = 0 Then
                            PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!RV_PODataCompetenzaLiq), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                        End If
                                            
                        If LINK_LISTINO > 0 Then
                            PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                            ImportoScontiDaReg = 0
                            PrezzoMedio = 0
                            ImportoVarImballoDaReg = 0
                            ImportoCommissioniDaReg = 0
                            ImportoNettoIvaDaReg = PrezzoDaReg
                            Link_TMP_Prezzo_Medio = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Else
                            If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                    Case 1
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        Link_TMP_Prezzo_Medio = 0
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                    Case 2
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    Case Else
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                End Select
                            Else
                                If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                    If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        Link_TMP_Prezzo_Medio = 0
                                    Else
'                                        'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
'                                        Else
'                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                            Link_TMP_Prezzo_Medio = 0
'                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        End If
                                    End If
                                Else
                                    'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
'                                    Else
'                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                        Link_TMP_Prezzo_Medio = 0
'                                    End If
                                End If
                            End If
                        End If
                        
                        TrattenutePerLavorazione = 0
                        TrattenuteGenerali = 0
                        TrattenuteTotali = 0
    
                        TrattValGen1 = 0
                        TrattValGen2 = 0
                        TrattPercGen1 = 0
                        TrattPercGen2 = 0
                        
                        TrattValLav1 = 0
                        TrattValLav2 = 0
                        TrattPercLav1 = 0
                        TrattPercLav2 = 0
                        
                        TrattValPreLiq1 = 0
                        TrattValPreLiq2 = 0
                        
                        QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                        
                        If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                            QuantitaLiq = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1)
                            If QuantitaLiq = 0 Then
                                QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            End If
                        End If
                        
                        If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
                        If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)
                        
                        
                       'TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, QuantitaLiq, (fnNotNullN(PrezzoDaReg) + fnNotNullN(rsFatt!RV_POImportoDaLiq))
                        'If fnNotNullN(rsFatt!RV_POInvenduto) = 0 Then
                            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, QuantitaLiq, fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                        'End If
                        
                        sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                        sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                        sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                        sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                        sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                        sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                        sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                        sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                        sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                        sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                        sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                        sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                        sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                        sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                        sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                        sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                        sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, "
                        sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                        sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2, "
                        sSQL = sSQL & "TrattenutaValorePreLiq1, TrattenutaValorePreLiq2 ) "
                        
                        sSQL = sSQL & "VALUES ("
                        If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                            sSQL = sSQL & 1 & ", "
                        Else
                            sSQL = sSQL & 4 & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                        sSQL = sSQL & LINK_PERIODO & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                        sSQL = sSQL & fnNormString("D.D.T. n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(0) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                        
                        TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 5)
                        TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 5)
                        TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
                        
                        If LINK_LISTINO > 0 Then
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                        Else
'                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
'
'                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
'                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
'                            End Select
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq), 2)) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                             
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                        If LINK_LISTINO = 0 Then
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq2) & ")"
                        
                        
                        CnDMT.Execute sSQL
    
                        InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        
                    End If
                End If
            End If
        End If


        If (Unita_Progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
            
        DoEvents
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing

End Sub

Private Sub GET_VENDITA_FA(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Link_CategoriaMerceologica As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiq As Double

Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE DOCUMENTI DI FATTURA ACCOMPAGNATORIA"

'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq, "
sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0072.Doc_Prefisso, Articolo.RV_POIDCategoriaLiquidazione  "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0001.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0001.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0001.Link_Art_articolo "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If LINK_CAT_MERCE > 0 Then
    sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
End If


If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
Else
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE)
End If

sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1 "

If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POIDSocio=" & LINK_SOCIO
End If

sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    prg.Value = 0
    prg.Max = 1000
    Unita_Progresso = prg.Max / rsFatt.RecordCount
    
    While Not rsFatt.EOF
        AvviaLiquidazioneRiga = 1
        If NO_LIQ_VEND_UFF = 1 Then
            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                AvviaLiquidazioneRiga = 0
            End If
        End If
        If AvviaLiquidazioneRiga = 1 Then
            If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                    
                        ImportoScontiPM = 0
                        ImportoVarImballoPM = 0
                        ImportoCommissioniPM = 0
                        ImportoNettoIvaPM = 0
                        Link_TMP_Prezzo_Medio = 0
                        
                        If LINK_LISTINO = 0 Then
                            PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!RV_PODataCompetenzaLiq), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                        End If
                        
                        If LINK_LISTINO > 0 Then
                            PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                            ImportoScontiDaReg = 0
                            PrezzoMedio = 0
                            ImportoVarImballoDaReg = 0
                            ImportoCommissioniDaReg = 0
                            ImportoNettoIvaDaReg = PrezzoDaReg
                            Link_TMP_Prezzo_Medio = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Else
                            If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                    Case 1
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        Link_TMP_Prezzo_Medio = 0
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        
                                    Case 2
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    Case Else
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                End Select
                            Else
                                If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                    If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        Link_TMP_Prezzo_Medio = 0
                                    Else
                                        'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
'                                        Else
'                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                            Link_TMP_Prezzo_Medio = 0
'                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        End If
                                    End If
                                Else
                                    'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
'                                    Else
'                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                        Link_TMP_Prezzo_Medio = 0
'                                    End If
                                End If
                            End If
                        End If
                    
                    TrattenutePerLavorazione = 0
                    TrattenuteGenerali = 0
                    TrattenuteTotali = 0
                    
                    TrattValGen1 = 0
                    TrattValGen2 = 0
                    TrattPercGen1 = 0
                    TrattPercGen2 = 0
                    
                    TrattValLav1 = 0
                    TrattValLav2 = 0
                    TrattPercLav1 = 0
                    TrattPercLav2 = 0
                    
                    
                    TrattValPreLiq1 = 0
                    TrattValPreLiq2 = 0
                    
                    QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                    
                    If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                        QuantitaLiq = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1)
                        If QuantitaLiq = 0 Then
                            QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                        End If
                    End If

                    If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
                    If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)


                    'If fnNotNullN(rsFatt!RV_POInvenduto) = 0 Then
                        TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, QuantitaLiq, fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                    'End If
                    
                        sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                        sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                        sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                        sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                        sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                        sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                        sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                        sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                        sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                        sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                        sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                        sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                        sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                        sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                        sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                        sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                        sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, "
                        sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                        sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2, "
                        sSQL = sSQL & "TrattenutaValorePreLiq1, TrattenutaValorePreLiq2 ) "
                        
                        sSQL = sSQL & "VALUES ("
                        If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                            sSQL = sSQL & 1 & ", "
                        Else
                            sSQL = sSQL & 4 & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                        sSQL = sSQL & LINK_PERIODO & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                        sSQL = sSQL & fnNormString("F.A. n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(0) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                        
                        TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 5)
                        TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 5)
                        TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
                        
                        If LINK_LISTINO > 0 Then
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                        Else
'                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
'
'                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
'                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
'                            End Select
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq), 2)) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                             
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                        If LINK_LISTINO = 0 Then
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq2) & ")"
                        
                        
                        CnDMT.Execute sSQL
    
                        InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                                        
                    
                    End If
                End If
            End If
        End If
        FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione F.A. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
        FrmNuovoPeriodo.List1.AddItem "- Elaborazione F.A. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
        FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
        
        If (Unita_Progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
            
        DoEvents
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing

End Sub
Private Sub GET_VENDITA_SNF(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Link_CategoriaMerceologica As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long

Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiq As Double
Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double


FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE CORRISPETTIVI"
'SCONTRINO NON FISCALE
sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq, "
sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0008.Doc_Prefisso, Articolo.RV_POIDCategoriaLiquidazione "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0034.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0034.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0034.Link_Art_articolo "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If LINK_CAT_MERCE > 0 Then
    sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
End If


If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_PODataConferimento<=" & fnNormDate(DATA_FINE)
Else
    sSQL = sSQL & " AND ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE)
End If

sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1 "

If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POIDSocio=" & LINK_SOCIO
End If
sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    prg.Value = 0
    prg.Max = 1000
    Unita_Progresso = prg.Max / rsFatt.RecordCount
    
    While Not rsFatt.EOF
        AvviaLiquidazioneRiga = 1
        If NO_LIQ_VEND_UFF = 1 Then
            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                AvviaLiquidazioneRiga = 0
            End If
        End If
        If AvviaLiquidazioneRiga = 1 Then
            If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                                 
                        ImportoScontiPM = 0
                        ImportoVarImballoPM = 0
                        ImportoCommissioniPM = 0
                        ImportoNettoIvaPM = 0
                        Link_TMP_Prezzo_Medio = 0
                        
                        
                        If LINK_LISTINO = 0 Then
                            PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!RV_PODataCompetenzaLiq), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                        End If
                        
                        If LINK_LISTINO > 0 Then
                            PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                            PrezzoMedio = 0
                            ImportoScontiDaReg = 0
                            ImportoVarImballoDaReg = 0
                            ImportoCommissioniDaReg = 0
                            ImportoNettoIvaDaReg = PrezzoDaReg
                            Link_TMP_Prezzo_Medio = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Else
                     
                            If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                    Case 1
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        Link_TMP_Prezzo_Medio = 0
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        
                                    Case 2
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    Case Else
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                End Select
                            Else
                                If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                    If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        Link_TMP_Prezzo_Medio = 0
                                    Else
                                        'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
'                                        Else
'                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                            Link_TMP_Prezzo_Medio = 0
'                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        End If
                                    End If
                                Else
                                    'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
'                                    Else
'                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                        Link_TMP_Prezzo_Medio = 0
'                                    End If
                                End If
                            End If
                        End If
                        TrattenutePerLavorazione = 0
                        TrattenuteGenerali = 0
                        TrattenuteTotali = 0
                        
                        TrattValGen1 = 0
                        TrattValGen2 = 0
                        TrattPercGen1 = 0
                        TrattPercGen2 = 0
                        
                        TrattValLav1 = 0
                        TrattValLav2 = 0
                        TrattPercLav1 = 0
                        TrattPercLav2 = 0
                        
                        TrattValPreLiq1 = 0
                        TrattValPreLiq2 = 0
                        
                        QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                        
                        If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                            QuantitaLiq = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1)
                            If QuantitaLiq = 0 Then
                                QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            End If
                        End If
                        
                        If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
                        If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

                        
                       'TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, QuantitaLiq, (fnNotNullN(PrezzoDaReg) + fnNotNullN(rsFatt!RV_POImportoDaLiq))
                        'If fnNotNullN(rsFatt!RV_POInvenduto) = 0 Then
                            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, QuantitaLiq, fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                        'End If
                        
                        sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                        sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                        sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                        sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                        sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                        sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                        sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                        sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                        sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                        sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                        sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                        sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                        sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                        sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                        sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                        sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                        sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, "
                        sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                        sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2, "
                        sSQL = sSQL & "TrattenutaValorePreLiq1, TrattenutaValorePreLiq2 ) "
                        
                        sSQL = sSQL & "VALUES ("
                        If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                            sSQL = sSQL & 1 & ", "
                        Else
                            sSQL = sSQL & 4 & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                        sSQL = sSQL & LINK_PERIODO & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                        sSQL = sSQL & fnNormString("S.N.F. n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(0) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                        
                        TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 5)
                        TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 5)
                        TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
                        
                        If LINK_LISTINO > 0 Then
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                        Else
'                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
'
'                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
'                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
'                            End Select
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq), 2)) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                             
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                        If LINK_LISTINO = 0 Then
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq2) & ")"
                        
                        CnDMT.Execute sSQL
    
                        InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        
                    End If
                End If
            End If
        End If
        FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione D.d.t. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
        FrmNuovoPeriodo.List1.AddItem "- Elaborazione D.d.t. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
        FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
        
        
        If (Unita_Progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
            
        DoEvents
                        
                        
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing




End Sub
Private Sub TrattenuteArticolo(IDSocio As Long, IDArticolo As Long, IDCategoriaMerceologica As Long, IDTipoLavorazione As Long, DalConferimento As Boolean, Quantita As Double, PrezzoArticolo As Double, IDRigaConferimento As Long, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long, Invenduto As Long)
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

rs.Open sSQL, CnDMT.InternalConnection

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
    
    rs.Open sSQL, CnDMT.InternalConnection

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
            rsTratt.Open sSQLTratt, CnDMT.InternalConnection
                                
            If rsTratt.EOF = False Then
                
                TrattValPreLiq1 = fnNotNullN(rsTratt!ValoreTrattenutaPreLiq1)
                TrattValPreLiq2 = fnNotNullN(rsTratt!ValoreTrattenutaPreLiq2)
            
                
                If DalConferimento = True Then
                    If rs!Tipo4 = 1 Then
                        'A valore
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattValLav1 = TrattValLav1 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                        TrattValLav2 = TrattValLav2 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
    
                    Else
                        'A valore
                        TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattValGen1 = TrattValGen1 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                        TrattValGen2 = TrattValGen2 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                    End If
                Else
                    If rs!Tipo4 = 1 Then
                        'A valore
                        If Invenduto = 0 Then
                            TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                            TrattValLav1 = TrattValLav1 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                            TrattenutePerLavorazione = TrattenutePerLavorazione + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                            TrattValLav2 = TrattValLav2 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                        End If
                        'A percentuale
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                        TrattPercLav1 = TrattPercLav1 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                        TrattenutePerLavorazione = TrattenutePerLavorazione + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        TrattPercLav2 = TrattPercLav2 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                    Else
                        If TIPO_QUANTITA = 4 Then
                            'A valore
                            If Invenduto = 0 Then
                                TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                                TrattValGen1 = TrattValGen1 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta1))
                                TrattenuteGenerali = TrattenuteGenerali + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                                TrattValGen2 = TrattValGen2 + (Quantita * fnNotNullN(rsTratt!ValoreTrattenuta2))
                            End If
                            'A percentuale
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattPercGen1 = TrattPercGen1 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                            TrattPercGen2 = TrattPercGen2 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        Else
                            'A percentuale
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattPercGen1 = TrattPercGen1 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                            TrattenuteGenerali = TrattenuteGenerali + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                            TrattPercGen2 = TrattPercGen2 + (((PrezzoArticolo * Quantita) / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                        End If
                    End If
                End If
            End If
            
            If DalConferimento = False Then
                SCRIVI_TRATTENUTA_RIGA_CONFERIMENTO IDArticolo, LINK_PERIODO, IDRigaConferimento, rsTratt, IDTipoOggetto, IDOggetto, IDValoriOggettoDettaglio, fnNotNullN(rs!IDRV_POTipoTrattenuta), False, IDSocio
            Else
                SCRIVI_TRATTENUTA_RIGA_CONFERIMENTO 0, LINK_PERIODO, IDRigaConferimento, rsTratt, IDTipoOggetto, IDOggetto, IDValoriOggettoDettaglio, fnNotNullN(rs!IDRV_POTipoTrattenuta), True, IDSocio
            End If
            
            rsTratt.Close
            Set rsTratt = Nothing
    End If
'
'    TrattenuteGenerali = FormatNumber(TrattenuteGenerali, 5)
'    TrattenutePerLavorazione = FormatNumber(TrattenutePerLavorazione, 5)
    TrattenuteGenerali = TrattenuteGenerali
    TrattenutePerLavorazione = TrattenutePerLavorazione
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
        
        GetSQL = ""
        Exit Function
    End If
Else
    sSQL = sSQL & " AND IDSocio=" & 0
End If
If Tipo2 = 1 Then
    If ValueTipo2 > 0 Then
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & ValueTipo2
    Else
        
        GetSQL = ""
        Exit Function
    End If
Else
    sSQL = sSQL & " AND IDCategoriaMerceologica=" & 0
End If
If Tipo3 = 1 Then
    If ValueTipo3 > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & ValueTipo3
    Else
        GetSQL = ""
        Exit Function
    End If
Else
    sSQL = sSQL & " AND IDArticolo=" & 0
End If
If Tipo4 = 1 Then
    If ValueTipo4 > 0 Then
        sSQL = sSQL & " AND IDTipoLavorazione=" & ValueTipo4
    Else
        GetSQL = ""
        Exit Function
    End If
Else
    sSQL = sSQL & " AND IDTipoLavorazione=" & 0
End If




If sSQL <> "" Then
    GetSQL = GetSQL & sSQL
Else
    GetSQL = ""
End If
End Function
Private Function GetPrezzoUnitarioPeriodo(IDSocio As Long, IDArticolo As Long, DataConferimento As String, DataVendita As String, DataLavorazione As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, ImportoNettoVenditaPM As Double, IDTipoPrezzoMedioTMP As Long) As Double
Dim sSQL As String
Dim rsAvr As ADODB.Recordset
Dim IDTipoPrezzoMedio As Long

Dim TotaleQuantita As Double
Dim TotaleVendita As Double

Dim TotaleVenditaSconti As Double
Dim TotaleVenditaComm As Double
Dim TotaleVenditaVarImb As Double
Dim TotaleVenditaNettoIva As Double
Dim Moltiplicatore As Double


Dim DATA_INIZIO_PERIODO As String
Dim DATA_FINE_PERIODO As String
Dim Base As Integer
Dim Settimana As Long
Dim rsArt As DmtOleDbLib.adoResultset
Dim CalcoloPerCategoria As Boolean
Dim IDCategoriaMerceologica_local As Long
Dim IDSocio_local As Long
Dim LINK_TIPO_CALCOLO_PREZZO_MEDIO_NC As Long
Dim LINK_TIPO_CALCOLO_PREZZO_MEDIO_ND As Long
Dim PrezzoMedio_Local As String

Dim rsOggetto As ADODB.Recordset
Dim Link_Riga_Oggetto As Long
Dim QantitaPerPrezzoMedio As Double

Dim AvviaPrezzoMedio As Boolean

'CREAZIONE RECORDSET'''''''''''''''''''''''''''''''''''''''''''''''''
Set rsOggetto = New ADODB.Recordset
rsOggetto.CursorLocation = adUseClient

rsOggetto.Fields.Append "IDRiga", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "IDTipoOggetto", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "IDValoriOggettoDettaglio", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "IDCategoriaMerceologica", adInteger, , adFldIsNullable
rsOggetto.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "QuantitaDoc", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "ImportoNettoVendita", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "ImportoSconti", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "ImportoVarImballo", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "ImportoCommissioni", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "ImportoLiquidazione", adDouble, , adFldIsNullable
rsOggetto.Fields.Append "DataVendita", adDBTimeStamp, , adFldIsNullable
rsOggetto.Fields.Append "DataConferimento", adDBTimeStamp, , adFldIsNullable
rsOggetto.Fields.Append "DataLavorazione", adDBTimeStamp, , adFldIsNullable

rsOggetto.Open , , adOpenKeyset, adLockBatchOptimistic
Link_Riga_Oggetto = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

PrezzoMedio_Local = ""
IDSocio_local = 0

IDCategoriaMerceologica_local = GET_CATEGORIA_MERCEOLOGICA(IDArticolo)


'If ((IDCategoriaMerceologica_local = 300) And (DataConferimento = "16/03/2019")) Then
'    MsgBox "STOP"
'End If

sSQL = "SELECT RV_POIDTipoPrezzoMedio "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rsArt = CnDMT.OpenResultset(sSQL)

LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0

If rsArt.EOF Then
    LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0
Else
    LINK_TIPO_PREZZO_MEDIO_ARTICOLO = fnNotNullN(rsArt!RV_POIDTipoPrezzoMedio)
End If

rsArt.CloseResultset
Set rsArt = Nothing

If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then
    LINK_TIPO_PREZZO_MEDIO_ARTICOLO = TIPO_CALCOLO_PREZZO_MEDIO
End If

Select Case LINK_TIPO_PREZZO_MEDIO_ARTICOLO

    Case 1 'Giorno di conferimento
        DATA_INIZIO_PERIODO = DataConferimento
        DATA_FINE_PERIODO = DataConferimento
        Base = 1
        CalcoloPerCategoria = False
        Link_TMP_Prezzo_Medio = 1
    Case 2 'Giorno di vendita
        DATA_INIZIO_PERIODO = DataVendita
        DATA_FINE_PERIODO = DataVendita
        CalcoloPerCategoria = False
        Base = 2
        Link_TMP_Prezzo_Medio = 2
    Case 3 'Settimana di conferimento
        Settimana = DatePart("w", DataConferimento, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataConferimento)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 1
        CalcoloPerCategoria = False
        Link_TMP_Prezzo_Medio = 3
    Case 4 'Settimana di vendita
        Settimana = DatePart("w", DataVendita, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataVendita)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 2
        CalcoloPerCategoria = False
        Link_TMP_Prezzo_Medio = 4
    Case 5 'Mese di conferimento
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataConferimento)) = 1), "0" & Month(DataConferimento), Month(DataConferimento)) & "/" & Year(DataConferimento)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataConferimento)
        Base = 1
        CalcoloPerCategoria = False
        Link_TMP_Prezzo_Medio = 5
    Case 6 'Mese di vendita
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataVendita)) = 1), "0" & Month(DataVendita), Month(DataVendita)) & "/" & Year(DataVendita)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataVendita)
        Base = 2
        CalcoloPerCategoria = False
        Link_TMP_Prezzo_Medio = 6
    Case 7 'Periodo di liquidazione per data di conferimento
        Base = 1
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = False
    Case 8 'Periodo di liquidazione per data di vendita
        Base = 2
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = False
    
    Case 9 'Giorno di conferimento per categoria merceologica
        DATA_INIZIO_PERIODO = DataConferimento
        DATA_FINE_PERIODO = DataConferimento
        Base = 1
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 10 'Giorno di vendita per categoria merceologica
        DATA_INIZIO_PERIODO = DataVendita
        DATA_FINE_PERIODO = DataVendita
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
        Base = 2
        
    Case 11 'Settimana di conferimento per categoria merceologica
        Settimana = DatePart("w", DataConferimento, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataConferimento)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 1
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 12 'Settimana di vendita per categoria merceologica
        Settimana = DatePart("w", DataVendita, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataVendita)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 2
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 13 'Mese di conferimento per categoria merceologica
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataConferimento)) = 1), "0" & Month(DataConferimento), Month(DataConferimento)) & "/" & Year(DataConferimento)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataConferimento)
        Base = 1
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 14 'Mese di vendita per categoria merceologica
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataVendita)) = 1), "0" & Month(DataVendita), Month(DataVendita)) & "/" & Year(DataVendita)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataVendita)
        Base = 2
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 15 'Periodo di liquidazione per data di conferimento per categoria merceologica
        Base = 1
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 16 'Periodo di liquidazione per data di vendita per categoria merceologica
        Base = 2
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    
    Case 17 'Per socio per articolo giorno di conferimento
        DATA_INIZIO_PERIODO = DataConferimento
        DATA_FINE_PERIODO = DataConferimento
        Base = 1
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 18 'Per socio per articolo giorno di vendita
        DATA_INIZIO_PERIODO = DataVendita
        DATA_FINE_PERIODO = DataVendita
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
        Base = 2
    Case 19 'Per socio per articolo settimana di conferimento
        Settimana = DatePart("w", DataConferimento, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataConferimento)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 1
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 20 'per socio per articolo settimana di vendita
        Settimana = DatePart("w", DataVendita, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataVendita)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 2
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 21 'Per socio per articolo mese di conferimento
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataConferimento)) = 1), "0" & Month(DataConferimento), Month(DataConferimento)) & "/" & Year(DataConferimento)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataConferimento)
        Base = 1
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 22 'Per socio per articolo mese di vendita
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataVendita)) = 1), "0" & Month(DataVendita), Month(DataVendita)) & "/" & Year(DataVendita)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataVendita)
        Base = 2
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 23 'Per socio per articolo periodo di liquidazione per data di conferimento
        Base = 1
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    Case 24 'Per socio per articolo periodo di liquidazione per data di vendita
        Base = 2
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    
    Case 25 'Per socio per categoria merceologica giorno di conferimento
        DATA_INIZIO_PERIODO = DataConferimento
        DATA_FINE_PERIODO = DataConferimento
        Base = 1
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 26 'Per socio per categoria merceologica giorno di vendita
        DATA_INIZIO_PERIODO = DataVendita
        DATA_FINE_PERIODO = DataVendita
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
        Base = 2
    Case 27 'Per socio per categoria merceologica settimana di conferimento
        Settimana = DatePart("w", DataConferimento, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataConferimento)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 1
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 28 'per socio per categoria merceologica settimana di vendita
        Settimana = DatePart("w", DataVendita, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataVendita)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 2
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 29 'Per socio per categoria merceologica mese di conferimento
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataConferimento)) = 1), "0" & Month(DataConferimento), Month(DataConferimento)) & "/" & Year(DataConferimento)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataConferimento)
        Base = 1
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 30 'Per socio per categoria merceologica mese di vendita
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataVendita)) = 1), "0" & Month(DataVendita), Month(DataVendita)) & "/" & Year(DataVendita)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataVendita)
        Base = 2
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 31 'Per socio per categoria merceologica periodo di liquidazione per data di conferimento
        Base = 1
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 32 'Per socio per categoria merceologica periodo di liquidazione per data di vendita
        Base = 2
        DATA_INIZIO_PERIODO = DATA_INIZIO
        DATA_FINE_PERIODO = DATA_FINE
        CalcoloPerCategoria = True
        IDSocio_local = IDSocio
    Case 33 'Per giorno di lavorazione
        DATA_INIZIO_PERIODO = DataLavorazione
        DATA_FINE_PERIODO = DataLavorazione
        CalcoloPerCategoria = False
        Base = 3
    Case 34 'Per settimana del giorno di lavorazione
        Settimana = DatePart("w", DataLavorazione, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataLavorazione)
        DATA_FINE_PERIODO = DateAdd("d", 6, DataLavorazione)
        Base = 3
        CalcoloPerCategoria = False
    Case 35 'Mese del giorno di lavorazione
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataLavorazione)) = 1), "0" & Month(DataLavorazione), Month(DataLavorazione)) & "/" & Year(DataLavorazione)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataLavorazione)
        Base = 3
        CalcoloPerCategoria = False
    Case 36 'Per categoria merceologica articolo per giorno di lavorazione
        DATA_INIZIO_PERIODO = DataLavorazione
        DATA_FINE_PERIODO = DataLavorazione
        Base = 3
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    Case 37 'Per categoria merceologica articolo settimana del giorno di lavorazione
        If Len(DataLavorazione) > 0 Then
            Settimana = DatePart("w", DataLavorazione, vbMonday) - 1
            DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataLavorazione)
            DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
            Base = 3
            If IDCategoriaMerceologica_local > 0 Then
                CalcoloPerCategoria = True
            Else
                CalcoloPerCategoria = False
            End If
        Else
            Settimana = DatePart("w", DataVendita, vbMonday) - 1
            DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataVendita)
            DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
            Base = 3
            If IDCategoriaMerceologica_local > 0 Then
                CalcoloPerCategoria = True
            Else
                CalcoloPerCategoria = False
            End If
        End If
    Case 38 'Per categoria merceologica articolo per mese del giorno di lavorazione
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataLavorazione)) = 1), "0" & Month(DataLavorazione), Month(DataLavorazione)) & "/" & Year(DataLavorazione)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataLavorazione)
        Base = 3
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
    
    Case 39 'Per socio per articolo per giorno di lavorazione
        DATA_INIZIO_PERIODO = DataLavorazione
        DATA_FINE_PERIODO = DataLavorazione
        Base = 3
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    
    Case 40 'Per socio per articolo per settimana del giorno di lavorazione
        Settimana = DatePart("w", DataLavorazione, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataLavorazione)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 3
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    
    Case 41 'Per socio per articolo del mese di lavorazione
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataLavorazione)) = 1), "0" & Month(DataLavorazione), Month(DataLavorazione)) & "/" & Year(DataLavorazione)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataLavorazione)
        Base = 3
        CalcoloPerCategoria = False
        IDSocio_local = IDSocio
    
    Case 42 'Per socio per categoria merceologica per  giorno di lavorazione
        DATA_INIZIO_PERIODO = DataLavorazione
        DATA_FINE_PERIODO = DataLavorazione
        Base = 3
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
        IDSocio_local = IDSocio
    Case 43 'Per socio per categoria merceologica per settimana del giorno di lavorazione
        Settimana = DatePart("w", DataLavorazione, vbMonday) - 1
        DATA_INIZIO_PERIODO = DateAdd("d", (-Settimana), DataLavorazione)
        DATA_FINE_PERIODO = DateAdd("d", 6, DATA_INIZIO_PERIODO)
        Base = 3
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
        IDSocio_local = IDSocio
    
    Case 44 'Per socio per categoria merceologica per mese di lavorazione
        DATA_INIZIO_PERIODO = "01/" & IIf((Len(Month(DataLavorazione)) = 1), "0" & Month(DataLavorazione), Month(DataLavorazione)) & "/" & Year(DataLavorazione)
        DATA_FINE_PERIODO = GET_DATA_FINE_MESE(DataLavorazione)
        Base = 3
        If IDCategoriaMerceologica_local > 0 Then
            CalcoloPerCategoria = True
        Else
            CalcoloPerCategoria = False
        End If
        IDSocio_local = IDSocio
        
    Case Else
        DATA_INIZIO_PERIODO = DataConferimento
        DATA_FINE_PERIODO = DataConferimento
        CalcoloPerCategoria = False
        Base = 1
End Select

TotaleQuantita = 0
TotaleVendita = 0
TotaleVenditaSconti = 0
TotaleVenditaVarImb = 0
TotaleVenditaComm = 0
TotaleVenditaNettoIva = 0

If LIQUIDA_FORNITORE = 1 Then
    IDSocio_local = IDSocio
End If

If CalcoloPerCategoria = False Then
    PrezzoMedio_Local = GET_CALCOLO_PREZZO_MEDIO_ARTICOLO(CalcoloPerCategoria, IDArticolo, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, IDSocio_local, DATA_INIZIO_PERIODO, DATA_FINE_PERIODO, Base, ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoVenditaPM, IDTipoPrezzoMedioTMP)
Else
    PrezzoMedio_Local = GET_CALCOLO_PREZZO_MEDIO_ARTICOLO(CalcoloPerCategoria, IDCategoriaMerceologica_local, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, IDSocio_local, DATA_INIZIO_PERIODO, DATA_FINE_PERIODO, Base, ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoVenditaPM, IDTipoPrezzoMedioTMP)
End If

If Len(PrezzoMedio_Local) > 0 Then
    GetPrezzoUnitarioPeriodo = PrezzoMedio_Local
    Exit Function
End If

If CalcoloPerCategoria = False Then
    sSQL = "SELECT ValoriOggettoDettaglio0004.RV_POImportoLiq, ValoriOggettoDettaglio0004.RV_POQuantitaLiq, ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0004.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0004.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0004.IDTipoOggetto, ValoriOggettoDettaglio0004.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.Link_art_articolo, ValoriOggettoDettaglio0004.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0004.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PODataLavorazione, ValoriOggettoPerTipo0002.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POInvenduto, ValoriOggettoDettaglio0004.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0004.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0004.RV_POIDTipoUtilizzoLinea, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PORigaRiscontroPeso, ValoriOggettoDettaglio0004.RV_POIDTipoDocumentoCoop, ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0004.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0004.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Art_Articolo=" & IDArticolo
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select
    
    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
    
Else
    sSQL = "SELECT ValoriOggettoDettaglio0004.RV_POImportoLiq, ValoriOggettoDettaglio0004.RV_POQuantitaLiq, ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0004.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0004.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0004.IDTipoOggetto, ValoriOggettoDettaglio0004.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.Link_art_articolo, ValoriOggettoDettaglio0004.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0004.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PODataLavorazione, ValoriOggettoPerTipo0002.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POInvenduto, ValoriOggettoDettaglio0004.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0004.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0004.RV_POIDTipoUtilizzoLinea, "
    sSQL = sSQL & "ValoriOggettoDettaglio0004.RV_PORigaRiscontroPeso, ValoriOggettoDettaglio0004.RV_POIDTipoDocumentoCoop, ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0004.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0004.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE IDCategoriaMerceologica = " & IDCategoriaMerceologica_local
    sSQL = sSQL & " AND Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"

    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
End If

Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, CnDMT.InternalConnection
X = 0
While Not rsAvr.EOF
    'If GET_CONFERIMENTO_MERCE(fnNotNullN(rsAvr!RV_POIDConferimentoRighe)) = False Then
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(rsAvr!Link_Art_Articolo)
    
    Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
        Case 0
            If fnNotNullN(rsAvr!RV_POIDTipoUtilizzoLinea) = 0 Then
                If FrmNuovoPeriodo.chkPrzMedioIVGamma.Value = vbChecked Then
                    AvviaPrezzoMedio = True
                Else
                    AvviaPrezzoMedio = False
                End If
            Else
                AvviaPrezzoMedio = True
            End If
        Case 1
            AvviaPrezzoMedio = True
        Case 2
            If LIQUIDA_FORNITORE = 1 Then
                AvviaPrezzoMedio = True
            Else
                AvviaPrezzoMedio = False
            End If

        Case Else
            AvviaPrezzoMedio = True
    End Select
    
    If AvviaPrezzoMedio = False Then
        TotaleQuantita = TotaleQuantita
        TotaleVendita = TotaleVendita
        
        TotaleVenditaSconti = TotaleVenditaSconti
        TotaleVenditaVarImb = TotaleVenditaVarImb
        TotaleVenditaComm = TotaleVenditaComm
        TotaleVenditaNettoIva = TotaleVenditaNettoIva
        
        X = X + fnNotNullN(rsAvr!RV_POQuantitaLiq)
    Else
        If fnNotNullN(rsAvr!RV_POImportoLiq) <> 0 Then
            If fnNotNullN(rsAvr!RV_PORigaRiscontroPeso) = 1 Then
                Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
                    Case 1
                        If RISCONTRO_PESO_SOCIO_VAL = 1 Then
                            TotaleQuantita = TotaleQuantita
                            TotaleVenditaSconti = TotaleVenditaSconti
                            TotaleVenditaVarImb = TotaleVenditaVarImb
                            TotaleVenditaComm = TotaleVenditaComm
                            TotaleVenditaNettoIva = TotaleVenditaNettoIva
                        Else
                            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
                            TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                            TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                            TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                            TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                        End If
                    Case 2
                        If RISCONTRO_PESO_FORN_VAL = 1 Then
                            TotaleQuantita = TotaleQuantita
                            TotaleVenditaSconti = TotaleVenditaSconti
                            TotaleVenditaVarImb = TotaleVenditaVarImb
                            TotaleVenditaComm = TotaleVenditaComm
                            TotaleVenditaNettoIva = TotaleVenditaNettoIva
                        Else
                            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
                            TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                            TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                            TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                            TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                        End If
                    Case Else
                        TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
                        TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                        TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                        TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                End Select
            Else
                TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
            End If
            
            TotaleVendita = TotaleVendita + (FormatNumber(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * FormatNumber(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
            
        Else
            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
            'TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
            TotaleVendita = TotaleVendita
            
            TotaleVenditaSconti = TotaleVenditaSconti
            TotaleVenditaVarImb = TotaleVenditaVarImb
            TotaleVenditaComm = TotaleVenditaComm
            TotaleVenditaNettoIva = TotaleVenditaNettoIva
            
        End If
        
        GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
        fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
        fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)

    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing
'Close #1
'DA FATTURA ACCOMPAGNATORIA
If CalcoloPerCategoria = False Then
    sSQL = "SELECT ValoriOggettoDettaglio0001.RV_POImportoLiq, ValoriOggettoDettaglio0001.RV_POQuantitaLiq, ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0001.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0001.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0001.IDTipoOggetto, ValoriOggettoDettaglio0001.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.Link_art_articolo, ValoriOggettoDettaglio0001.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0001.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_PODataLavorazione, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POInvenduto, ValoriOggettoDettaglio0001.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0001.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0001.RV_POIDTipoUtilizzoLinea "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0001.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0001.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Art_Articolo=" & IDArticolo
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"

    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If

Else
    sSQL = "SELECT ValoriOggettoDettaglio0001.RV_POImportoLiq, ValoriOggettoDettaglio0001.RV_POQuantitaLiq, ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0001.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0001.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0001.IDTipoOggetto, ValoriOggettoDettaglio0001.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.Link_art_articolo, ValoriOggettoDettaglio0001.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0001.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_PODataLavorazione, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POInvenduto, ValoriOggettoDettaglio0001.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0001.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0001.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0001.RV_POIDTipoUtilizzoLinea "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0001.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0001.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE IDCategoriaMerceologica = " & IDCategoriaMerceologica_local
    sSQL = sSQL & " AND Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select
    
    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
End If


Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, CnDMT.InternalConnection

While Not rsAvr.EOF
    'If GET_CONFERIMENTO_MERCE(fnNotNullN(rsAvr!RV_POIDConferimentoRighe)) = False Then
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(rsAvr!Link_Art_Articolo)
    
    'If Moltiplicatore > 1 Then
    '    MsgBox "STOP"
    'End If
    
    Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
        Case 0
            If fnNotNullN(rsAvr!RV_POIDTipoUtilizzoLinea) = 0 Then
                If FrmNuovoPeriodo.chkPrzMedioIVGamma.Value = vbChecked Then
                    AvviaPrezzoMedio = True
                Else
                    AvviaPrezzoMedio = False
                End If
            Else
                AvviaPrezzoMedio = True
            End If
        Case 1
            AvviaPrezzoMedio = True
        Case 2
            If LIQUIDA_FORNITORE = 1 Then
                AvviaPrezzoMedio = True
            Else
                AvviaPrezzoMedio = False
            End If
        Case Else
            AvviaPrezzoMedio = True
    End Select
    
    If AvviaPrezzoMedio = False Then
        
        TotaleQuantita = TotaleQuantita
        TotaleVendita = TotaleVendita
        TotaleVenditaSconti = TotaleVenditaSconti
        TotaleVenditaVarImb = TotaleVenditaVarImb
        TotaleVenditaComm = TotaleVenditaComm
        TotaleVenditaNettoIva = TotaleVenditaNettoIva


    Else
        If fnNotNullN(rsAvr!RV_POImportoLiq) <> 0 Then
            'TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
            
            'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
            TotaleVendita = TotaleVendita + (FormatNumber(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * FormatNumber(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
            TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
            TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
            TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
            TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
        
        Else
            'TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
            TotaleVendita = TotaleVendita
        
            TotaleVenditaSconti = TotaleVenditaSconti
            TotaleVenditaVarImb = TotaleVenditaVarImb
            TotaleVenditaComm = TotaleVenditaComm
            TotaleVenditaNettoIva = TotaleVenditaNettoIva
            
        End If
        
        GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
        fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
        fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)

    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing


'DA SCONTRINO NON FISCALE
If CalcoloPerCategoria = False Then
    sSQL = "SELECT ValoriOggettoDettaglio0034.RV_POImportoLiq, ValoriOggettoDettaglio0034.RV_POQuantitaLiq, ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0034.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0034.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0034.IDTipoOggetto, ValoriOggettoDettaglio0034.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.Link_art_articolo, ValoriOggettoDettaglio0034.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0034.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_PODataLavorazione, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POInvenduto, ValoriOggettoDettaglio0034.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0034.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0034.RV_POIDTipoUtilizzoLinea  "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0034.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0034.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Art_Articolo = " & IDArticolo
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If

Else
    sSQL = "SELECT ValoriOggettoDettaglio0034.RV_POImportoLiq, ValoriOggettoDettaglio0034.RV_POQuantitaLiq, ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0034.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0034.RV_POImportoDaLiq, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0034.IDTipoOggetto, ValoriOggettoDettaglio0034.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.Link_art_articolo, ValoriOggettoDettaglio0034.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0034.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_PODataLavorazione, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq , "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POInvenduto, ValoriOggettoDettaglio0034.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0034.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0034.RV_POIDTipoDocumentoCoop, ValoriOggettoDettaglio0034.RV_POIDTipoUtilizzoLinea  "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0034.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0034.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDCategoriaMerceologica = " & IDCategoriaMerceologica_local
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
End If

Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, CnDMT.InternalConnection

While Not rsAvr.EOF
    'If GET_CONFERIMENTO_MERCE(fnNotNullN(rsAvr!RV_POIDConferimentoRighe)) = False Then
    
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(rsAvr!Link_Art_Articolo)
    
    'If Moltiplicatore > 1 Then
    '    MsgBox "STOP"
    'End If
    
    Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
        Case 0
            If fnNotNullN(rsAvr!RV_POIDTipoUtilizzoLinea) = 0 Then
                If FrmNuovoPeriodo.chkPrzMedioIVGamma.Value = vbChecked Then
                    AvviaPrezzoMedio = True
                Else
                    AvviaPrezzoMedio = False
                End If
            Else
                AvviaPrezzoMedio = True
            End If
        Case 1
            AvviaPrezzoMedio = True
        Case 2
            If LIQUIDA_FORNITORE = 1 Then
                AvviaPrezzoMedio = True
            Else
                AvviaPrezzoMedio = False
            End If

        Case Else
            AvviaPrezzoMedio = True
    End Select
    
    If AvviaPrezzoMedio = False Then
        
        TotaleQuantita = TotaleQuantita
        TotaleVendita = TotaleVendita
        TotaleVenditaSconti = TotaleVenditaSconti
        TotaleVenditaVarImb = TotaleVenditaVarImb
        TotaleVenditaComm = TotaleVenditaComm
        TotaleVenditaNettoIva = TotaleVenditaNettoIva
    Else
        If fnNotNullN(rsAvr!RV_POImportoLiq) <> 0 Then
            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
            'TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
            'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
            TotaleVendita = TotaleVendita + (FormatNumber(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * FormatNumber(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
            TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
            TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
            TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
            TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
        
        Else
            'TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
            TotaleQuantita = Round(TotaleQuantita, 5) + Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5)
            TotaleVendita = TotaleVendita
        
            TotaleVenditaSconti = TotaleVenditaSconti
            TotaleVenditaVarImb = TotaleVenditaVarImb
            TotaleVenditaComm = TotaleVenditaComm
            TotaleVenditaNettoIva = TotaleVenditaNettoIva
        
        End If
        
        GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
        fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
        fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)

    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

'NOTA DI CREDITO
If CalcoloPerCategoria = False Then
    sSQL = "SELECT ValoriOggettoDettaglio0016.RV_POImportoLiq, ValoriOggettoDettaglio0016.Art_quantita_totale, ValoriOggettoDettaglio0016.RV_POQuantitaOrigine, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POQuantitaLiq, ValoriOggettoDettaglio0016.RV_POIDConferimentoRighe, ValoriOggettoDettaglio0016.RV_POIDTipoVariazione, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDTipoOggetto, ValoriOggettoDettaglio0016.RV_POIDValoriOggettoDettaglio, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POCodiceLotto, ValoriOggettoDettaglio0016.RV_POIDOggetto, ValoriOggettoDettaglio0016.Art_prezzo_unitario_netto_IVA, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0016.RV_POImportoDaLiq, ValoriOggettoDettaglio0016.RV_POImportoRigaCommissioni, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0016.IDTipoOggetto, ValoriOggettoDettaglio0016.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Link_art_articolo, ValoriOggettoDettaglio0016.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0016.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_PODataLavorazione, ValoriOggettoPerTipo000B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0016.RV_POIDTipoDocumentoCoop, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDTipoUtilizzoLinea, ValoriOggettoPerTipo000B.RV_PODataCompetenzaLiq "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0016.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0016.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Art_Articolo = " & IDArticolo
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If

    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If

Else
    
    sSQL = "SELECT ValoriOggettoDettaglio0016.RV_POImportoLiq, ValoriOggettoDettaglio0016.Art_quantita_totale, ValoriOggettoDettaglio0016.RV_POQuantitaOrigine, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POQuantitaLiq, ValoriOggettoDettaglio0016.RV_POIDConferimentoRighe, ValoriOggettoDettaglio0016.Art_prezzo_unitario_netto_IVA, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDTipoOggetto,ValoriOggettoDettaglio0016.RV_POIDOggetto, ValoriOggettoDettaglio0016.RV_POIDValoriOggettoDettaglio, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POCodiceLotto, ValoriOggettoDettaglio0016.RV_POIDTipoVariazione,  "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0016.RV_POImportoDaLiq, ValoriOggettoDettaglio0016.RV_POImportoRigaCommissioni, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0016.IDTipoOggetto, ValoriOggettoDettaglio0016.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Link_art_articolo, ValoriOggettoDettaglio0016.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0016.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_PODataLavorazione, ValoriOggettoPerTipo000B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0016.RV_POIDTipoDocumentoCoop, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.RV_POIDTipoUtilizzoLinea, ValoriOggettoPerTipo000B.RV_PODataCompetenzaLiq "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0016.Link_Art_articolo = Articolo.IDArticolo INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0016.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0016.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDCategoriaMerceologica = " & IDCategoriaMerceologica_local
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
End If

Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, CnDMT.InternalConnection

While Not rsAvr.EOF
    LINK_TIPO_CALCOLO_PREZZO_MEDIO_NC = GET_TIPO_CALCOLO_PREZZO_MEDIO_NC
    
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(rsAvr!Link_Art_Articolo)
    
    Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
        Case 0
            If fnNotNullN(rsAvr!RV_POIDTipoUtilizzoLinea) = 0 Then
                If FrmNuovoPeriodo.chkPrzMedioIVGamma.Value = vbChecked Then
                    AvviaPrezzoMedio = True
                Else
                    AvviaPrezzoMedio = False
                End If
            Else
                AvviaPrezzoMedio = True
            End If
        Case 1
            AvviaPrezzoMedio = True
        Case 2
            If LIQUIDA_FORNITORE = 1 Then
                AvviaPrezzoMedio = True
            Else
                AvviaPrezzoMedio = False
            End If

        Case Else
            AvviaPrezzoMedio = True
    End Select
    'If GET_CONFERIMENTO_MERCE(fnNotNullN(rsAvr!RV_POIDConferimentoRighe)) = True Then
    If AvviaPrezzoMedio = False Then
        TotaleQuantita = TotaleQuantita
        TotaleVendita = TotaleVendita
        TotaleVenditaSconti = TotaleVenditaSconti
        TotaleVenditaVarImb = TotaleVenditaVarImb
        TotaleVenditaComm = TotaleVenditaComm
        TotaleVenditaNettoIva = TotaleVenditaNettoIva
    Else
        If GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO(fnNotNullN(rsAvr!RV_POIDTipoOggetto), fnNotNullN(rsAvr!RV_POIDValoriOggettoDettaglio), fnNotNull(rsAvr!RV_POCodiceLotto), fnNotNullN(rsAvr!RV_POIDOggetto)) = 0 Then
            TotaleQuantita = TotaleQuantita
            TotaleVendita = TotaleVendita
        Else
            If fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) <> 0 Then
                Select Case LINK_TIPO_CALCOLO_PREZZO_MEDIO_NC
                    Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1 'VARIAZIONE PARZIALE DI PREZZO
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                            Case 2 'VARIAZIONE PARZIALE DI PESO
                                TotaleQuantita = TotaleQuantita + (-rsAvr!RV_POQuantitaLiq) ' fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                QantitaPerPrezzoMedio = fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            Case 3 'VARIAZIONE TOTALE DI PESO
                                TotaleQuantita = TotaleQuantita + (-rsAvr!RV_POQuantitaLiq) 'fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                QantitaPerPrezzoMedio = fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            Case 4 'VARIAZIONE PARZIALE DI PREZZO
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                        End Select
                        
                        TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                    
                        TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                        TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                        TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                                            
                                            
                        GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                        fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                        fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                    
                    
                    Case 2 'INCLUSO VARIAZIONE PESO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1
                                TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(-rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            
                            
                            Case 2
                                TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(-rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            
                            
                            Case 3
                                TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(-rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            
                            Case 4
                                TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                        
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(-rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                        
                        End Select
                        
                    Case 3 'INCLUSO VARIAZIONE PREZZO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1
                                TotaleQuantita = TotaleQuantita
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))

                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, 0, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)

                            Case 2
                                TotaleQuantita = TotaleQuantita
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, 0, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                            
                            Case 3
                                TotaleQuantita = TotaleQuantita
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                                
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, 0, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                                
                            Case 4
                                TotaleQuantita = TotaleQuantita
                                'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(-rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))

                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, 0, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)

                        End Select
                    
                    Case 4 'ESCLUDI VARIAZIONE PESO E PREZZO
                        TotaleQuantita = TotaleQuantita
                        TotaleVendita = TotaleVendita
                        
                        TotaleVenditaSconti = TotaleVenditaSconti
                        TotaleVenditaVarImb = TotaleVenditaVarImb
                        TotaleVenditaComm = TotaleVenditaComm
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva

                    Case Else 'INCLUSO VARIAZIONE PREZZO E PESO
                    
                        TotaleQuantita = TotaleQuantita
                        TotaleVendita = TotaleVendita
                        
                        TotaleVenditaSconti = TotaleVenditaSconti
                        TotaleVenditaVarImb = TotaleVenditaVarImb
                        TotaleVenditaComm = TotaleVenditaComm
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva
                        
                        'Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                        '    Case 1
                        '        TotaleQuantita = TotaleQuantita
                        '        QantitaPerPrezzoMedio = 0
                        '    Case 2
                        '        TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                        '        QantitaPerPrezzoMedio = fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                        '    Case 3
                        '        TotaleQuantita = TotaleQuantita + fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                        '        QantitaPerPrezzoMedio = fnNotNullN(-rsAvr!RV_POQuantitaLiq)
                        '    Case 4
                        '        TotaleQuantita = TotaleQuantita
                        '        QantitaPerPrezzoMedio = 0
                        '
                        'End Select
                        
                        'TotaleVendita = TotaleVendita + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                        'TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                        'TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                        'TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(-rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                        'TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(-rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                
                        'GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                        'fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                        'fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(-rsAvr!RV_POQuantitaLiq)

                
                End Select
                
            
            End If
        End If
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

'NOTA DI DEBITO
If CalcoloPerCategoria = False Then
    sSQL = "SELECT ValoriOggettoDettaglio0007.RV_POImportoLiq, ValoriOggettoDettaglio0007.Art_quantita_totale, ValoriOggettoDettaglio0007.RV_POQuantitaOrigine, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POQuantitaLiq, ValoriOggettoDettaglio0007.RV_POIDConferimentoRighe, ValoriOggettoDettaglio0007.RV_POIDOggetto, ValoriOggettoDettaglio0007.RV_POIDTipoVariazione, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDTipoOggetto, ValoriOggettoDettaglio0007.RV_POIDValoriOggettoDettaglio, ValoriOggettoDettaglio0007.RV_POCodiceLotto, ValoriOggettoDettaglio0007.Art_prezzo_unitario_netto_IVA, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0007.RV_POImportoDaLiq, ValoriOggettoDettaglio0007.RV_POImportoRigaCommissioni, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0007.IDTipoOggetto, ValoriOggettoDettaglio0007.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Link_art_articolo, ValoriOggettoDettaglio0007.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0007.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_PODataLavorazione, ValoriOggettoPerTipo006B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0007.RV_POIDTipoDocumentoCoop, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDTipoUtilizzoLinea, ValoriOggettoPerTipo006B.RV_PODataCompetenzaLiq "

    sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0007.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0007.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Art_Articolo = " & IDArticolo
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select

    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If

Else
    sSQL = "SELECT ValoriOggettoDettaglio0007.RV_POImportoLiq, ValoriOggettoDettaglio0007.Art_quantita_totale, ValoriOggettoDettaglio0007.RV_POQuantitaOrigine, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POQuantitaLiq, ValoriOggettoDettaglio0007.RV_POIDConferimentoRighe, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDTipoOggetto, ValoriOggettoDettaglio0007.RV_POIDValoriOggettoDettaglio, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POCodiceLotto, ValoriOggettoDettaglio0007.RV_POIDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Art_prezzo_unitario_netto_IVA, ValoriOggettoDettaglio0007.RV_POIDTipoVariazione, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Art_pre_uni_net_sco_net_IVA, ValoriOggettoDettaglio0007.RV_POImportoDaLiq, ValoriOggettoDettaglio0007.RV_POImportoRigaCommissioni, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POImportoRigaCommissioni, ValoriOggettoDettaglio0007.IDTipoOggetto, ValoriOggettoDettaglio0007.IDOggetto, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Link_art_articolo, ValoriOggettoDettaglio0007.IDValoriOggettoDettaglio,ValoriOggettoDettaglio0007.RV_PODataConferimento, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_PODataLavorazione, ValoriOggettoPerTipo006B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDProcessoIVGamma, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDAssegnazioneMerce, ValoriOggettoDettaglio0007.RV_POIDTipoDocumentoCoop, "
    sSQL = sSQL & "ValoriOggettoDettaglio0007.RV_POIDTipoUtilizzoLinea, ValoriOggettoPerTipo006B.RV_PODataCompetenzaLiq "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0007.Link_Art_articolo = Articolo.IDArticolo INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0007.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0007.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDCategoriaMerceologica = " & IDCategoriaMerceologica_local
    sSQL = sSQL & " AND RV_POPrezzoMedioInLiq=1"
    'If Base = 2 Then
    '    sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
    '    sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
    'Else
    '    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
    ' '   sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    'End If
    Select Case Base
        Case 1
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
        Case 2
            If TIPO_LIQUIDAZIONE = 2 Then
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DATA_FINE_PERIODO)
            Else
                sSQL = sSQL & " AND Doc_Data>=" & fnNormDate(DATA_INIZIO_PERIODO)
                sSQL = sSQL & " AND Doc_Data<=" & fnNormDate(DATA_FINE_PERIODO)
            End If
        Case 3
            sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DATA_FINE_PERIODO)
        Case Else
            sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DATA_INIZIO_PERIODO)
            sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DATA_FINE_PERIODO)
    End Select


    If IDSocio_local > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & IDSocio_local
    End If
End If

Set rsAvr = New ADODB.Recordset
rsAvr.Open sSQL, CnDMT.InternalConnection

While Not rsAvr.EOF
    LINK_TIPO_CALCOLO_PREZZO_MEDIO_ND = GET_TIPO_CALCOLO_PREZZO_MEDIO_ND
    
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(rsAvr!Link_Art_Articolo)
    
    Select Case fnNotNullN(rsAvr!RV_POIDTipoDocumentoCoop)
        Case 0
            If fnNotNullN(rsAvr!RV_POIDTipoUtilizzoLinea) = 0 Then
                If FrmNuovoPeriodo.chkPrzMedioIVGamma.Value = vbChecked Then
                    AvviaPrezzoMedio = True
                Else
                    AvviaPrezzoMedio = False
                End If
            Else
                AvviaPrezzoMedio = True
            End If
        Case 1
            AvviaPrezzoMedio = True
        Case 2
            If LIQUIDA_FORNITORE = 1 Then
                AvviaPrezzoMedio = True
            Else
                AvviaPrezzoMedio = False
            End If

        Case Else
            AvviaPrezzoMedio = True
    End Select
    'If GET_CONFERIMENTO_MERCE(fnNotNullN(rsAvr!RV_POIDConferimentoRighe)) = True Then
    If AvviaPrezzoMedio = False Then
        TotaleQuantita = TotaleQuantita
        TotaleVendita = TotaleVendita
        TotaleVenditaSconti = TotaleVenditaSconti
        TotaleVenditaVarImb = TotaleVenditaVarImb
        TotaleVenditaComm = TotaleVenditaComm
        TotaleVenditaNettoIva = TotaleVenditaNettoIva
    Else
    
        If GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO(fnNotNullN(rsAvr!RV_POIDTipoOggetto), fnNotNullN(rsAvr!RV_POIDValoriOggettoDettaglio), fnNotNull(rsAvr!RV_POCodiceLotto), fnNotNullN(rsAvr!RV_POIDOggetto)) = 0 Then
            TotaleQuantita = TotaleQuantita
            TotaleVendita = TotaleVendita
        Else
            If fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) <> 0 Then
'                If fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) < 0 Then
'                    MsgBox "STOP"
'                End If
            
                Select Case LINK_TIPO_CALCOLO_PREZZO_MEDIO_ND
                    Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                            Case 2
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                QantitaPerPrezzoMedio = fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                
                            Case 3
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                QantitaPerPrezzoMedio = fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            Case 4
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                        End Select
                        
                        TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                        
                        TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                        TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                        TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))

                        GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                        fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                        fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                    
                    
                    Case 2 'INCLUSO VARIAZIONE PESO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            
                            
                            Case 2
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            
                            Case 3
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            
                            Case 4
                                TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                        
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, fnNotNullN(rsAvr!RV_POQuantitaLiq), fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                
                        End Select
                    Case 3 'INCLUSO VARIAZIONE PREZZO
                        Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                            Case 1
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))

                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                                
                                
                            Case 2
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            
                            Case 3
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                            
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                            
                            Case 4
                                TotaleQuantita = TotaleQuantita
                                QantitaPerPrezzoMedio = 0
                                'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                                TotaleVendita = TotaleVendita + (Round(fnNotNullN(rsAvr!RV_POQuantitaLiq), 5) * Round(fnNotNullN(rsAvr!RV_POImportoLiq), 5))
                                TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                                TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                                TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                                TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                        
                                GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                                fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                                fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                        
                        End Select
                    
                    Case 4 'ESCLUDI VARIAZIONE PESO E PREZZO
                        TotaleQuantita = TotaleQuantita
                        TotaleVendita = TotaleVendita
                    
                        TotaleVenditaSconti = TotaleVenditaSconti
                        TotaleVenditaVarImb = TotaleVenditaVarImb
                        TotaleVenditaComm = TotaleVenditaComm
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva
                    
                    Case Else 'INCLUSO VARIAZIONE PREZZO E PESO
                        TotaleQuantita = TotaleQuantita
                        TotaleVendita = TotaleVendita
                    
                        TotaleVenditaSconti = TotaleVenditaSconti
                        TotaleVenditaVarImb = TotaleVenditaVarImb
                        TotaleVenditaComm = TotaleVenditaComm
                        TotaleVenditaNettoIva = TotaleVenditaNettoIva
                    
                    
                    
                        'Select Case fnNotNullN(rsAvr!RV_POIDTipoVariazione)
                        '    Case 1
                        '        TotaleQuantita = TotaleQuantita
                        '        QantitaPerPrezzoMedio = 0
                        '    Case 2
                        '        TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                        '        QantitaPerPrezzoMedio = fnNotNullN(rsAvr!RV_POQuantitaLiq)
                        '    Case 3
                        '        TotaleQuantita = TotaleQuantita + fnNotNullN(rsAvr!RV_POQuantitaLiq)
                        '        QantitaPerPrezzoMedio = fnNotNullN(rsAvr!RV_POQuantitaLiq)
                        '    Case 4
                        '        TotaleQuantita = TotaleQuantita
                        '        QantitaPerPrezzoMedio = 0
                        'End Select

                        'TotaleVendita = TotaleVendita + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoLiq))
                
                        'TotaleVenditaSconti = TotaleVenditaSconti + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)))
                        'TotaleVenditaVarImb = TotaleVenditaVarImb + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoDaLiq))
                        'TotaleVenditaComm = TotaleVenditaComm + (fnNotNullN(rsAvr!RV_POQuantitaLiq) * fnNotNullN(rsAvr!RV_POImportoRigaCommissioni))
                        'TotaleVenditaNettoIva = TotaleVenditaNettoIva + ((fnNotNullN(rsAvr!RV_POQuantitaLiq) / Moltiplicatore) * fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA))
                
                        'GET_RIGA_PER_PREZZO_MEDIO rsOggetto, Link_Riga_Oggetto, fnNotNullN(rsAvr!IDTipoOggetto), fnNotNullN(rsAvr!IDOggetto), fnNotNullN(rsAvr!IDValoriOggettoDettaglio), _
                        'fnNotNullN(rsAvr!Link_Art_Articolo), IDCategoriaMerceologica_local, QantitaPerPrezzoMedio, fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) / Moltiplicatore, (fnNotNullN(rsAvr!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsAvr!Art_pre_uni_net_sco_net_IVA)) / Moltiplicatore, _
                        'fnNotNullN(rsAvr!RV_POImportoDaLiq), fnNotNullN(rsAvr!RV_POImportoRigaCommissioni), fnNotNullN(rsAvr!RV_POImportoLiq), fnNotNull(rsAvr!doc_data), fnNotNull(rsAvr!RV_PODataConferimento), fnNotNull(rsAvr!RV_PODataLavorazione), fnNotNullN(rsAvr!RV_POQuantitaLiq)
                
                End Select
                
            
            End If


        End If
    End If
rsAvr.MoveNext
Wend

rsAvr.Close
Set rsAvr = Nothing

If TotaleQuantita > 0 Then
    GetPrezzoUnitarioPeriodo = FormatNumber((TotaleVendita / TotaleQuantita), 18)
    ImportoScontiPM = FormatNumber((TotaleVenditaSconti / TotaleQuantita), 18)
    ImportoVarImballoPM = FormatNumber((TotaleVenditaVarImb / TotaleQuantita), 18)
    ImportoCommissioniPM = FormatNumber((TotaleVenditaComm / TotaleQuantita), 18)
    ImportoNettoVenditaPM = FormatNumber((TotaleVenditaNettoIva / TotaleQuantita), 18)
Else
    GetPrezzoUnitarioPeriodo = 0
    ImportoScontiPM = 0
    ImportoVarImballoPM = 0
    ImportoCommissioniPM = 0
    ImportoNettoVenditaPM = 0
End If

    If CalcoloPerCategoria = False Then
        SALVA_PREZZO_MEDIO_ARTICOLO False, IDArticolo, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, IDSocio_local, GetPrezzoUnitarioPeriodo, Base, DATA_INIZIO_PERIODO, DATA_FINE_PERIODO, ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoVenditaPM, IDTipoPrezzoMedioTMP, rsOggetto
    Else
        SALVA_PREZZO_MEDIO_ARTICOLO True, IDCategoriaMerceologica_local, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, IDSocio_local, GetPrezzoUnitarioPeriodo, Base, DATA_INIZIO_PERIODO, DATA_FINE_PERIODO, ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoVenditaPM, IDTipoPrezzoMedioTMP, rsOggetto
    End If

End Function
Private Function GET_DATA_FINE_MESE(DataRiferimento As String) As String
Dim Mese As Integer
Dim Anno As Integer

Mese = Month(DataRiferimento)
Anno = Year(DataRiferimento)

Select Case Mese
    Case 1
        GET_DATA_FINE_MESE = "31/01/" & Anno
    Case 2
        If Anno Mod 4 = 0 Then
            GET_DATA_FINE_MESE = "29/02/" & Anno
        Else
            GET_DATA_FINE_MESE = "28/02/" & Anno
        End If
    Case 3
        GET_DATA_FINE_MESE = "31/03/" & Anno
    Case 4
        GET_DATA_FINE_MESE = "30/04/" & Anno
    Case 5
        GET_DATA_FINE_MESE = "31/05/" & Anno
    Case 6
        GET_DATA_FINE_MESE = "30/06/" & Anno
    Case 7
        GET_DATA_FINE_MESE = "31/07/" & Anno
    Case 8
        GET_DATA_FINE_MESE = "31/08/" & Anno
    Case 9
        GET_DATA_FINE_MESE = "30/09/" & Anno
    Case 10
        GET_DATA_FINE_MESE = "31/10/" & Anno
    Case 11
        GET_DATA_FINE_MESE = "30/11/" & Anno
    Case 12
        GET_DATA_FINE_MESE = "31/12/" & Anno
End Select



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
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
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

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CATEGORIA_MERCEOLOGICA = 0
Else
    GET_CATEGORIA_MERCEOLOGICA = fnNotNullN(rs!IDCategoriaMerceologica)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub InserisciNotaCredito(ValoriOggettoDettaglio As Long, IDTipoOggetto As Long, PrezzoMedio As Double, PrezzoMedioInFattura As Long, CodiceLottoVendita As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, ImportoNettoIvaPM As Double, Link_TMP_Prezzo_Medio As Long, IDTipoPrezzoMedio As Long)
Dim LINK_TIPO_CALCOLO_NC As Long
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim PrezzoDaReg As Double
Dim ImportoScontiDaReg As Double
Dim ImportoVarImballoDaReg As Double
Dim ImportoCommissioniDaReg As Double
Dim ImportoNettoIvaDaReg As Double
Dim Link_Riga_Prezzo_Medio_TMP As Long
Dim QuantitaLiq As Double

Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double
Dim LiquidaAPrezzoMedio As Boolean


'NOTA DI CREDITO
sSQL = "SELECT ValoriOggettoDettaglio0016.*, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero, "
sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo000B.Doc_Prefisso "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0016.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0016.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0016.Link_Art_articolo "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDTipoOggetto=" & IDTipoOggetto
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio 'Tolto per incongruenza con la riga del documento
End If


Set rsFatt = CnDMT.OpenResultset(sSQL)

While Not rsFatt.EOF
    'If (rsFatt!IDOggetto = 188224) And (rsFatt!IDValoriOggettoDettaglio = 7971) Then
    '    MsgBox "STOP"
    'End If
    
    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
            LiquidaAPrezzoMedio = False
            If LINK_LISTINO > 0 Then
                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                ImportoScontiDaReg = 0
                'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                ImportoVarImballoDaReg = 0
                
                ImportoCommissioniDaReg = 0
                ImportoNettoIvaDaReg = 0
                Link_Riga_Prezzo_Medio_TMP = 0
                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                LiquidaAPrezzoMedio = False
                
            Else
            
                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                        Case 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Case 2
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                        Case Else
                            PrezzoDaReg = PrezzoMedio
                            
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            
                            
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                    End Select
                Else
                    If TIPO_IMPORTO_ARTICOLO = 1 Then
                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                            
                        Else
                            'If PrezzoMedioInFattura = 1 Then
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                PrezzoDaReg = PrezzoMedio
                                ImportoScontiDaReg = ImportoScontiPM
                                'ImportoVarImballoDaReg = ImportoVarImballoPM
                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                
                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                                LiquidaAPrezzoMedio = True
'                            Else
'                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                ImportoScontiDaReg = ImportoScontiPM
'                                'ImportoVarImballoDaReg = ImportoVarImballoPM
'                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                                ImportoCommissioniDaReg = ImportoCommissioniPM
'                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
'                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
'
'                            End If
                        End If
                    Else
                        'If PrezzoMedioInFattura = 1 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            LiquidaAPrezzoMedio = True
'                        Else
'                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                            Link_Riga_Prezzo_Medio_TMP = 0
'                        End If
                    End If
                End If
            End If
            If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
                QuantitaLiq = -fnNotNullN(rsFatt!RV_POQuantitaLiq)
            Else
                If LiquidaAPrezzoMedio = True Then
                    QuantitaLiq = 0
                Else
                    QuantitaLiq = -fnNotNullN(rsFatt!RV_POQuantitaLiq)
                End If
            End If
            If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                QuantitaLiq = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1)
                If QuantitaLiq = 0 Then
                    QuantitaLiq = -fnNotNullN(rsFatt!RV_POQuantitaLiq)
                End If
            End If
            If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
                QuantitaLiq = QuantitaLiq
            Else
                If LiquidaAPrezzoMedio = True Then
                    QuantitaLiq = 0
                Else
                    QuantitaLiq = QuantitaLiq
                End If
            End If
            
            If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
            If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

            
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
            sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, "
            ''''PEZZO NUOVO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiDaReg, ImportoVarImpImballoDaReg, ImportoCommissioniDaReg, ImpUniVendDocNettoIvaVenditaDaReg, "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            sSQL = sSQL & "IDRV_POTipoOggettoVariante, IDCategoriaMerceologica, "
            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali,  "
            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato,  "
            
            '''PEZZO NUOVO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, "
            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, "
            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, "
            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, "
            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto) "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & 3 & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
            sSQL = sSQL & fnNormString(Mid(fnNotNull(rsFatt!Art_Descrizione), 1, 150)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
            sSQL = sSQL & fnNormString("N.C. n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) + (PrezzoMedio * fnNotNullN(QuantitaLiq))) & ", "
            
            TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 5)
            TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 5)
            TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
            
            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
            
                Case 1
                    If LINK_LISTINO = 0 Then
'                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        
                        sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                        
                        sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                    Else
                    
                        If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
'                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                            
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                            
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        End If
                    End If
                Case 3
                    LINK_TIPO_CALCOLO_NC = GET_TIPO_CALCOLO_PREZZO_MEDIO_NC
                    
                    Select Case LINK_TIPO_CALCOLO_NC
                        Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case Else
                                
                            End Select
                        Case 2 'INCLUSO VARIAZIONE PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                    
                                Case 2  'Variazione parziale di peso
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "

                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 3  'Variazione di peso totale
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select
                                
                                                        
                        Case 3 'INCLUSO VARAZIONE PREZZO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case Else
                                
                            End Select

                        
                        Case 4 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select

                        
                        Case Else 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select



                            'Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                            '    Case 1  'Variazione parziale di prezzo
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0 * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                            '        sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (0 * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            '
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '
                            '    Case 2  'Variazione parziale di peso
                            '        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                            '        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(-rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (-rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (-rsFatt!RV_POQuantitaLiq)) + (PrezzoMedio * fnNotNullN(-rsFatt!RV_POQuantitaLiq))) & ", "
                            '
                            '        sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            '
                            '
                            '    Case 3  'Variazione di peso totale
                            '        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                            '        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(-rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (-rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (-rsFatt!RV_POQuantitaLiq)) + (PrezzoMedio * fnNotNullN(-rsFatt!RV_POQuantitaLiq))) & ", "'
'
                            '        sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            '        sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            '    Case 4  'Variazione di prezzo totale
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0 * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            '        sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                            '        sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (0 * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            '
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            ' '       sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '        sSQL = sSQL & fnNormNumber(0) & ", "
                            '    Case Else
                            'End Select
                    End Select
            End Select
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNotNullN(1) & ", "
                Case 2
                    sSQL = sSQL & fnNotNullN(2) & ", "
                Case 3
                    sSQL = sSQL & fnNotNullN(2) & ", "
                Case 4
                    sSQL = sSQL & fnNotNullN(1) & ", "
            End Select
            
            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case 2
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 3
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 4
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case Else
                    
            End Select
            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ","
            'IMPORTO SCONTI
            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
            'IMPORTO VARIAZIONE IMBALLO
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
            'IMPORTO COMMISSIONI
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
            'IMPORTO UNITARIO NETTO IVA
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
            sSQL = sSQL & Link_Riga_Prezzo_Medio_TMP & ", "
            sSQL = sSQL & IDTipoPrezzoMedio & ", "
            sSQL = sSQL & fnNotNullN(0) & ")"

    CnDMT.Execute sSQL
End If
    
        
rsFatt.MoveNext
Wend

rsFatt.CloseResultset
Set rsFatt = Nothing
End Sub
Private Sub InserisciNotaDebito(ValoriOggettoDettaglio As Long, IDTipoOggetto As Long, PrezzoMedio As Double, PrezzoMedioInFattura As Long, CodiceLottoVendita As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, ImportoNettoIvaPM As Double, Link_TMP_Prezzo_Medio As Long, IDTipoPrezzoMedio As Long)
Dim LINK_TIPO_CALCOLO_ND As Long
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim PrezzoDaReg As Double
Dim ImportoScontiDaReg As Double
Dim ImportoVarImballoDaReg As Double
Dim ImportoCommissioniDaReg As Double
Dim ImportoNettoIvaDaReg As Double
Dim Link_Riga_Prezzo_Medio_TMP As Long
Dim QuantitaLiq As Double

Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double
Dim LiquidaAPrezzoMedio As Boolean

'NOTA DI DEBITO
sSQL = "SELECT ValoriOggettoDettaglio0007.*, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero, "
sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo006B.Doc_prefisso "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
sSQL = sSQL & "ValoriOggettoDettaglio0007 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0007.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0007.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0007.Link_Art_articolo "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
'sSQL = sSQL & " AND ValoriOggettoDettaglio0007.RV_POIDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND ValoriOggettoDettaglio0007.RV_POIDTipoOggetto=" & IDTipoOggetto
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0007.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND ValoriOggettoDettaglio0007.RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio
End If

Set rsFatt = CnDMT.OpenResultset(sSQL)

While Not rsFatt.EOF
    If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
            LiquidaAPrezzoMedio = False
            If LINK_LISTINO > 0 Then
                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                ImportoScontiDaReg = 0
                'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                ImportoVarImballoDaReg = 0
                
                ImportoCommissioniDaReg = 0
                ImportoNettoIvaDaReg = 0
                Link_Riga_Prezzo_Medio_TMP = 0
                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
            
            Else
                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                        Case 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Case 2
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                        Case Else
                            PrezzoDaReg = PrezzoMedio
                            
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            
                            
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                    End Select
                Else
                    If TIPO_IMPORTO_ARTICOLO = 1 Then
                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                        Else
                            'If PrezzoMedioInFattura = 1 Then
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                PrezzoDaReg = PrezzoMedio
                                ImportoScontiDaReg = ImportoScontiPM
                                'ImportoVarImballoDaReg = ImportoVarImballoPM
                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                
                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                                LiquidaAPrezzoMedio = True
'                            Else
'                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                ImportoScontiDaReg = ImportoScontiPM
'                                'ImportoVarImballoDaReg = ImportoVarImballoPM
'                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                                ImportoCommissioniDaReg = ImportoCommissioniPM
'                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
'                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
'
'                            End If
                        End If
                    Else
                        'If PrezzoMedioInFattura = 1 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            LiquidaAPrezzoMedio = True
'                        Else
'                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                            Link_Riga_Prezzo_Medio_TMP = 0
                        'End If
                    End If
                End If
            End If
            
            QuantitaLiq = fnNotNullN(rsFatt!RV_POQuantitaLiq)
            
            If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                QuantitaLiq = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1)
                If QuantitaLiq = 0 Then
                    QuantitaLiq = -fnNotNullN(rsFatt!RV_POQuantitaLiq)
                End If
            End If
            
            If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
                QuantitaLiq = QuantitaLiq
            Else
                If LiquidaAPrezzoMedio = True Then
                    QuantitaLiq = 0
                Else
                    QuantitaLiq = QuantitaLiq
                End If
            End If
            
            If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
            If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

            
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
            sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, "
            ''''PEZZO NUOVO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiDaReg, ImportoVarImpImballoDaReg, ImportoCommissioniDaReg, ImpUniVendDocNettoIvaVenditaDaReg, "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            sSQL = sSQL & "IDRV_POTipoOggettoVariante, IDCategoriaMerceologica, "
            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali,  "
            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato,  "
            
            '''PEZZO NUOVO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, "
            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, "
            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, "
            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, "
            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto) "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & 3 & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
            sSQL = sSQL & fnNormString("N.D. n° " & fnNotNull(rsFatt!Doc_Prefisso) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) + (PrezzoMedio * fnNotNullN(QuantitaLiq))) & ", "
            
            
            TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)
            TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 2)
            TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
            
            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
            
                Case 1
                    If LINK_LISTINO = 0 Then
'                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
'                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) & ", "
'                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                        
                        sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                    Else
                        If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
'                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                            
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        End If
                    End If
                Case 3
                    LINK_TIPO_CALCOLO_ND = GET_TIPO_CALCOLO_PREZZO_MEDIO_ND
                    
                    
                    Select Case LINK_TIPO_CALCOLO_ND
                        Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                
                                Case 1
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
        
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                Case 4
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "

                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                            
                            End Select
                        Case 2 'INCLUSO VARIAZIONE PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
        
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "

        
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                    
                                    
                            End Select
                                                        
                        Case 3 'INCLUSO VARAZIONE PREZZO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "

                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "

                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                            End Select
                        
                        Case 4 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            
                            End Select
                        
                        Case Else 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            
                            End Select
                    End Select
            End Select
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNotNullN(3) & ", "
                Case 2
                    sSQL = sSQL & fnNotNullN(4) & ", "
                Case 3
                    sSQL = sSQL & fnNotNullN(4) & ", "
                Case 4
                    sSQL = sSQL & fnNotNullN(3) & ", "
            End Select
            
            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case 2
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 3
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 4
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case Else
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            End Select
            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ","
            'IMPORTO SCONTI
            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
            'IMPORTO VARIAZIONE IMBALLO
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
            'IMPORTO COMMISSIONI
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
            'IMPORTO UNITARIO NETTO IVA
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
            sSQL = sSQL & Link_Riga_Prezzo_Medio_TMP & ", "
            sSQL = sSQL & IDTipoPrezzoMedio & ", "
            sSQL = sSQL & fnNotNullN(0) & ")"
            
        CnDMT.Execute sSQL
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

Private Function GET_CONFERIMENTO_MERCE(IDConferimentoRiga As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If LIQUIDA_FORNITORE = 1 Then
    GET_CONFERIMENTO_MERCE = True
    Exit Function
End If

sSQL = "SELECT IDTipoDocumentoCoop "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONFERIMENTO_MERCE = False
Else
    If fnNotNullN(rs!IDTipoDocumentoCoop) = 1 Then
        GET_CONFERIMENTO_MERCE = True
    Else
        GET_CONFERIMENTO_MERCE = False
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_IMPORTO_DOCUMENTO_DI_RIFERIMENTO(IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, CodiceLottoVendita As String, IDOggetto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


If IDTipoOggetto = 0 Then
    GET_DOCUMENTO_DI_RIFERIMENTO = 0
    Exit Function
End If

Select Case IDTipoOggetto
    Case 114
        sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0001.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0001.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
    Case 2
        sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0004.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0004.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
    
    Case 8
        sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0034.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0034.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)

End Select

'sSQL = sSQL & "  "
'sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_DOCUMENTO_DI_RIFERIMENTO = 0
Else
    GET_IMPORTO_DOCUMENTO_DI_RIFERIMENTO = FormatNumber(fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), 3)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO(IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, CodiceLottoVendita As String, IDOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


If IDTipoOggetto = 0 Then
    GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO = 0
    Exit Function
End If

Select Case IDTipoOggetto
    Case 114
        sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0001.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0001.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
    Case 2
        sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0004.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0004.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
    
    Case 8
        sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0034.Link_Art_IVA "
        sSQL = sSQL & " WHERE ValoriOggettoDettaglio0034.IDOggetto=" & IDOggetto
        sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)

End Select

'sSQL = sSQL & "  "
'sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO = 0
Else
    GET_PREZZO_MEDIO_DOCUMENTO_DI_RIFERIMENTO = fnNotNullN(rs!RV_POPrezzoMedioInLiq)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Function GET_TIPO_CALCOLO_PREZZO_MEDIO_NC() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCalcoloPrezzoMedioNC FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_CALCOLO_PREZZO_MEDIO_NC = 0
Else
    GET_TIPO_CALCOLO_PREZZO_MEDIO_NC = fnNotNullN(rs!IDTipoCalcoloPrezzoMedioNC)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TIPO_CALCOLO_PREZZO_MEDIO_ND() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCalcoloPrezzoMedioND FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_CALCOLO_PREZZO_MEDIO_ND = 0
Else
    GET_TIPO_CALCOLO_PREZZO_MEDIO_ND = fnNotNullN(rs!IDTipoCalcoloPrezzoMedioND)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CALCOLO_PREZZO_MEDIO_ARTICOLO(CalcoloPerCategoria As Boolean, IDArticolo As Long, IDTipoPrezzoMedio As Long, IDSocio As Long, DaData As String, AData As String, Base As Integer, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, PrezzoNettoIvaPM As Double, IDTipoPrezzoMedioTMP As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTMPLiquidazionePrezzoMedio, PrezzoMedio, PrezzoSconti, PrezzoCommissioni, PrezzoVarInclusoImballo, PrezzoMedioNettoIva "
sSQL = sSQL & "FROM RV_POTMPLiquidazionePrezzoMedio "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
If CalcoloPerCategoria = False Then
    sSQL = sSQL & " AND IDArticolo=" & IDArticolo
Else
    sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDArticolo
End If
sSQL = sSQL & " AND IDTipoPrezzoMedio=" & IDTipoPrezzoMedio
sSQL = sSQL & " AND IDSocio=" & IDSocio


Select Case Base
    Case 1
        sSQL = sSQL & " AND DaDataConferimento=" & fnNormDate(DaData)
        sSQL = sSQL & " AND ADataConferimento=" & fnNormDate(AData)
    Case 2
        sSQL = sSQL & " AND DaDataVendita=" & fnNormDate(DaData)
        sSQL = sSQL & " AND ADataVendita=" & fnNormDate(AData)
    Case 3
        sSQL = sSQL & " AND DaDataLavorazione=" & fnNormDate(DaData)
        sSQL = sSQL & " AND ADataLavorazione=" & fnNormDate(AData)
End Select

'If Base = 1 Then
'    sSQL = sSQL & " AND DaDataConferimento=" & fnNormDate(DaData)
'    sSQL = sSQL & " AND ADataConferimento=" & fnNormDate(AData)

'Else
'    sSQL = sSQL & " AND DaDataVendita=" & fnNormDate(DaData)
'    sSQL = sSQL & " AND ADataVendita=" & fnNormDate(AData)
'End If


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CALCOLO_PREZZO_MEDIO_ARTICOLO = ""
    ImportoScontiPM = 0
    ImportoVarImballoPM = 0
    ImportoCommissioniPM = 0
    PrezzoNettoIvaPM = 0
    IDTipoPrezzoMedioTMP = 0
Else
    GET_CALCOLO_PREZZO_MEDIO_ARTICOLO = fnNotNullN(rs!PrezzoMedio)
    ImportoScontiPM = fnNotNullN(rs!PrezzoSconti)
    ImportoVarImballoPM = fnNotNullN(rs!PrezzoVarInclusoImballo)
    ImportoCommissioniPM = fnNotNullN(rs!PrezzoCommissioni)
    PrezzoNettoIvaPM = fnNotNullN(rs!PrezzoMedioNettoIva)
    IDTipoPrezzoMedioTMP = fnNotNullN(rs!IDRV_POTMPLiquidazionePrezzoMedio)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function SALVA_PREZZO_MEDIO_ARTICOLO(PerCategoria As Boolean, IDArticolo As Long, IDTipoPrezzoMedio As Long, IDSocio As Long, PrezzoMedio As Double, Base As Integer, DaData As String, AData As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, PrezzoNettoIvaPM As Double, IDTipoPrezzoMedioTMP As Long, rsOggetto As ADODB.Recordset)
Dim sSQL As String
Dim IDRiga As Long
Dim rsNew As ADODB.Recordset

IDRiga = fnGetNewKey("RV_POTMPLiquidazionePrezzoMedio", "IDRV_POTMPLiquidazionePrezzoMedio")

sSQL = "INSERT INTO RV_POTMPLiquidazionePrezzoMedio ("
sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDUtente, IDArticolo, IDCategoriaMerceologica, "
sSQL = sSQL & "IDTipoPrezzoMedio, IDSocio, PrezzoMedio, "
sSQL = sSQL & "DaDataConferimento, ADataConferimento, DaDataVendita, ADataVendita, DaDataLavorazione, ADataLavorazione, "
sSQL = sSQL & "PrezzoSconti, PrezzoCommissioni, PrezzoVarInclusoImballo,PrezzoMedioNettoIva)"
sSQL = sSQL & "VALUES ("
sSQL = sSQL & IDRiga & ", "
sSQL = sSQL & TheApp.IDUser & ", "
If PerCategoria = False Then
    sSQL = sSQL & IDArticolo & ", "
    sSQL = sSQL & "0" & ", "
Else
    sSQL = sSQL & "0" & ", "
    sSQL = sSQL & IDArticolo & ", "
End If
sSQL = sSQL & IDTipoPrezzoMedio & ", "
sSQL = sSQL & IDSocio & ", "
sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
Select Case Base
    Case 1
        sSQL = sSQL & fnNormDate(DaData) & ", "
        sSQL = sSQL & fnNormDate(AData) & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
    Case 2
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate(DaData) & ", "
        sSQL = sSQL & fnNormDate(AData) & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
    
    Case 3
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate("") & ", "
        sSQL = sSQL & fnNormDate(DaData) & ", "
        sSQL = sSQL & fnNormDate(AData) & ", "
    Case Else
    
End Select

sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
sSQL = sSQL & fnNormNumber(PrezzoNettoIvaPM) & ")"

CnDMT.Execute sSQL

IDTipoPrezzoMedioTMP = IDRiga

If ((rsOggetto.EOF = True) And (rsOggetto.BOF = True)) Then Exit Function

sSQL = "SELECT * FROM RV_POTMPLiquidazionePrezzoMedioRighe "
sSQL = sSQL & "WHERE IDRV_POTMPLiquidazionePrezzoMedio=" & IDRiga

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rsOggetto.MoveFirst

While Not rsOggetto.EOF
    rsNew.AddNew
        rsNew!IDRV_POTMPLiquidazionePrezzoMedioRighe = fnGetNewKey("RV_POTMPLiquidazionePrezzoMedioRighe", "IDRV_POTMPLiquidazionePrezzoMedioRighe")
        rsNew!IDRV_POTMPLiquidazionePrezzoMedio = IDRiga
        rsNew!IDTipoOggetto = rsOggetto!IDTipoOggetto
        rsNew!IDOggetto = rsOggetto!IDOggetto
        rsNew!IDValoriOggettoDettaglio = rsOggetto!IDValoriOggettoDettaglio
        rsNew!IDArticolo = rsOggetto!IDArticolo
        rsNew!IDCategoriaMerceologica = rsOggetto!IDCategoriaMerceologica
        rsNew!Quantita = rsOggetto!Quantita
        rsNew!QuantitaDocumento = rsOggetto!QuantitaDoc
        rsNew!ImportoNettoVendita = rsOggetto!ImportoNettoVendita
        rsNew!ImportoSconti = rsOggetto!ImportoSconti
        rsNew!ImportoVarImballo = rsOggetto!ImportoVarImballo
        rsNew!ImportoCommissioni = rsOggetto!ImportoCommissioni
        rsNew!ImportoLiquidazioni = rsOggetto!ImportoLiquidazione
        rsNew!DataVendita = rsOggetto!DataVendita
        rsNew!DataConferimento = rsOggetto!DataConferimento
        rsNew!DataLavorazione = rsOggetto!DataLavorazione
    rsNew.Update
rsOggetto.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rsOggetto.Close
Set rsOggetto = Nothing

End Function

Private Function GET_TRATTENUTA_ARTICOLO_CONFERITO(IDArticoloConf As Long, ValColli As Double, ValPesoLordo As Double, ValTara As Double, ValPesoNetto As Double, ValPezzi As Double, IDUMConf As Long, IDRigaConferimento As Long, IDAnagraficaSocio As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_UM_Coop As Long
Dim VALORE_CONFERITO As Double
GET_TRATTENUTA_ARTICOLO_CONFERITO = 0

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POTrattenutaPerLiquidazione "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloConf
sSQL = sSQL & " AND IDSocio=0"
sSQL = sSQL & " AND IDCategoriaMerceologica=0"
sSQL = sSQL & " AND IDTipoLavorazione=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection


If Not rs.EOF Then
    
    If fnNotNullN(rs!IDUnitaDiMisuraConf) = 0 Then
        Link_UM_Coop = IDUMConf
    Else
        Link_UM_Coop = fnNotNullN(rs!IDUnitaDiMisuraConf)
    End If
    
    Select Case Link_UM_Coop
        Case 1
            VALORE_CONFERITO = ValColli
        Case 2
            VALORE_CONFERITO = ValPesoLordo
        Case 3
            VALORE_CONFERITO = ValPesoNetto
        Case 4
            VALORE_CONFERITO = ValTara
        Case 5
            VALORE_CONFERITO = ValPezzi
        Case Else
            VALORE_CONFERITO = 0
    End Select
    
    GET_TRATTENUTA_ARTICOLO_CONFERITO = GET_TRATTENUTA_ARTICOLO_CONFERITO + (VALORE_CONFERITO * fnNotNullN(rs!ValoreTrattenuta1Conf))
    GET_TRATTENUTA_ARTICOLO_CONFERITO = GET_TRATTENUTA_ARTICOLO_CONFERITO + (VALORE_CONFERITO * fnNotNullN(rs!ValoreTrattenuta2Conf))
    

    SCRIVI_TRATTENUTA_RIGA_CONFERIMENTO 0, LINK_PERIODO, IDRigaConferimento, rs, 0, 0, IDRigaConferimento, 0, True, 0

End If



rs.Close
Set rs = Nothing


End Function
Private Sub SCRIVI_TRATTENUTA_RIGA_CONFERIMENTO(IDArticoloVenduto As Long, IDPeriodoLiquidazione As Long, IDRigaConferimento As Long, rsTratt As ADODB.Recordset, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long, IDTipoTrattenuta As Long, SoloRigaConferimento As Boolean, IDAnagraficaSocio As Long)
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_POTMPLiquidazioneTrattConf "
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rsNew.AddNew
    rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
    rsNew!IDArticoloVendita = IDArticoloVenduto
    rsNew!IDRV_POPeriodoLiquidazione = IDPeriodoLiquidazione
    If Not rsTratt.EOF Then
        rsNew!IDRV_POTrattenutaPerLiquidazione = fnNotNullN(rsTratt!IDRV_POTrattenutaPerLiquidazione)
        rsNew!ValoreTrattenuta1 = fnNotNullN(rsTratt!ValoreTrattenuta1)
        rsNew!ValoreTrattenuta2 = fnNotNullN(rsTratt!ValoreTrattenuta2)
        rsNew!PercTrattenuta1 = fnNotNullN(rsTratt!PercTrattenuta1)
        rsNew!PercTrattenuta2 = fnNotNullN(rsTratt!PercTrattenuta2)
        rsNew!ValoreTrattenuta1Conf = fnNotNullN(rsTratt!ValoreTrattenuta1Conf)
        rsNew!ValoreTrattenuta2Conf = fnNotNullN(rsTratt!ValoreTrattenuta2Conf)
    Else
        rsNew!IDRV_POTrattenutaPerLiquidazione = 0
        rsNew!ValoreTrattenuta1 = 0
        rsNew!ValoreTrattenuta2 = 0
        rsNew!PercTrattenuta1 = 0
        rsNew!PercTrattenuta2 = 0
        rsNew!ValoreTrattenuta1Conf = 0
        rsNew!ValoreTrattenuta2Conf = 0
        
    End If
    rsNew!IDTipoOggetto = IDTipoOggetto
    rsNew!IDOggetto = IDOggetto
    rsNew!IDValoriOggettoDettaglio = IDValoriOggettoDettaglio
    rsNew!IDRV_POTipoTrattenuta = IDTipoTrattenuta
    rsNew!SoloRigaConferimento = SoloRigaConferimento
    rsNew!IDAnagraficaSocio = IDAnagraficaSocio
rsNew.Update

rsNew.Close
Set rsNew = Nothing

End Sub

Private Sub GET_VENDITA_DDT_SU_RICHIESTA(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long

Dim rsOgg As ADODB.Recordset
Dim NumeroRecord As Long
Dim AvviaLiquidazioneRiga As Long

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE DOCUMENTI DI TRASPORTO SELEZIONATI"

''''''''''''''''''''''CONTEGGIO DOCUMENTI DA ELABORARE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_POTMPLiqFattureSel) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=2 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

If rsOgg.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rsOgg!NumeroRecord)
End If

rsOgg.Close
Set rsOgg = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumeroRecord = 0 Then Exit Sub

Unita_Progresso = FormatNumber((prg.Max / NumeroRecord), 2)

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=2 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

While Not rsOgg.EOF
    'DOCUMENTO DI TRASPORTO
    sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, "
    sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
    sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0002.Doc_prefisso, Articolo.RV_POIDCategoriaLiquidazione "
    sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
    sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
    sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0004.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0004.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
    sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0004.Link_Art_articolo "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1 "
    sSQL = sSQL & " AND Oggetto.IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
    sSQL = sSQL & " AND Oggetto.IDTipoOggetto=" & fnNotNullN(rsOgg!IDTipoOggetto)

    If LINK_CAT_MERCE > 0 Then
        sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
    End If
    sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"
    
    
    Set rsFatt = New ADODB.Recordset
    rsFatt.Open sSQL, CnDMT.InternalConnection
    
    If rsFatt.EOF = False Then
        
        While Not rsFatt.EOF
            AvviaLiquidazioneRiga = 1
            If NO_LIQ_VEND_UFF = 1 Then
                If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                    AvviaLiquidazioneRiga = 0
                End If
            End If
            If AvviaLiquidazioneRiga = 1 Then
                If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                    If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                            
                            ImportoScontiPM = 0
                            ImportoVarImballoPM = 0
                            ImportoCommissioniPM = 0
                            ImportoNettoIvaPM = 0
                            Link_TMP_Prezzo_Medio = 0
                                
                            If LINK_LISTINO > 0 Then
                                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!doc_data), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                            End If
                                                
                            If LINK_LISTINO > 0 Then
                                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(sFatt!Link_Art_Articolo), LINK_LISTINO)
                                ImportoScontiDaReg = 0
                                PrezzoMedio = 0
                                ImportoVarImballoDaReg = 0
                                ImportoCommissioniDaReg = 0
                                ImportoNettoIvaDaReg = PrezzoDaReg
                                Link_TMP_Prezzo_Medio = 0
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            Else
                                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                        Case 1
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            
                                        Case 2
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Case Else
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    End Select
                                Else
                                    If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            Link_TMP_Prezzo_Medio = 0
                                        Else
                                            If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                                PrezzoDaReg = PrezzoMedio
                                                ImportoScontiDaReg = ImportoScontiPM
                                                ImportoVarImballoDaReg = ImportoVarImballoPM
                                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                                Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            Else
                                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                                ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                                ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                                ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                                ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                                Link_TMP_Prezzo_Medio = 0
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            End If
                                        End If
                                    Else
                                        If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Else
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                        End If
                                    End If
                                End If
                            End If
                            
                            TrattenutePerLavorazione = 0
                            TrattenuteGenerali = 0
                            TrattenuteTotali = 0
                            
                           'TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), (fnNotNullN(PrezzoDaReg) + fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                            
                            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                            sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto) "
                            
                            sSQL = sSQL & "VALUES ("
                            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                                sSQL = sSQL & 1 & ", "
                            Else
                                sSQL = sSQL & 4 & ", "
                            End If
                            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                            sSQL = sSQL & LINK_PERIODO & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                            sSQL = sSQL & fnNormString("D.D.T. n° " & fnNotNull(rsFatt!Doc_Prefisso) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(0) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
                            
                                Case 1
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                                Case 3
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            End Select
                            
                            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ")"
                            
                            
                            CnDMT.Execute sSQL
        
                            InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                            InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        End If
                    End If
                End If
            End If
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione D.d.t. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione D.d.t. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
                
            DoEvents
                            
                            
        rsFatt.MoveNext
        Wend
    End If
    rsFatt.Close
    Set rsFatt = Nothing


If (Unita_Progresso + prg.Value) >= prg.Max Then
    prg.Value = prg.Max
Else
    prg.Value = prg.Value + Unita_Progresso
End If
DoEvents

rsOgg.MoveNext
Wend
rsOgg.Close
Set rsOgg = Nothing
End Sub
Private Sub GET_VENDITA_FA_SU_RICHIESTA(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Unita_Progresso As Double
Dim rsOgg As ADODB.Recordset
Dim NumeroRecord As Long


Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double

Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim AvviaLiquidazioneRiga As Long

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE FATTURE ACCOMPAGNATORIE SELEZIONATI"

''''''''''''''''''''''CONTEGGIO DOCUMENTI DA ELABORARE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_POTMPLiqFattureSel) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=114 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

If rsOgg.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rsOgg!NumeroRecord)
End If

rsOgg.Close
Set rsOgg = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumeroRecord = 0 Then Exit Sub

Unita_Progresso = FormatNumber((prg.Max / NumeroRecord), 2)

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=114 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

While Not rsOgg.EOF
    'DOCUMENTO DI TRASPORTO
    sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, "
    sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
    sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0072.Doc_prefisso, Articolo.RV_POIDCategoriaLiquidazione "
    sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
    sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
    sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0001.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0001.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
    sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0001.Link_Art_articolo "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1 "
    sSQL = sSQL & " AND Oggetto.IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
    sSQL = sSQL & " AND Oggetto.IDTipoOggetto=" & fnNotNullN(rsOgg!IDTipoOggetto)

    If LINK_CAT_MERCE > 0 Then
        sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
    End If

    sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"
    
    
    Set rsFatt = New ADODB.Recordset
    rsFatt.Open sSQL, CnDMT.InternalConnection
    
    If rsFatt.EOF = False Then
        
        While Not rsFatt.EOF
            AvviaLiquidazioneRiga = 1
            
            If NO_LIQ_VEND_UFF = 1 Then
                If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                    AvviaLiquidazioneRiga = 0
                End If
            End If
            If AvviaLiquidazioneRiga = 1 Then
                If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                    If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                            
                            ImportoScontiPM = 0
                            ImportoVarImballoPM = 0
                            ImportoCommissioniPM = 0
                            ImportoNettoIvaPM = 0
                            Link_TMP_Prezzo_Medio = 0
                            
                            If LINK_LISTINO = 0 Then
                                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!doc_data), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                            End If
                            
                            If LINK_LISTINO > 0 Then
                                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(sFatt!Link_Art_Articolo), LINK_LISTINO)
                                PrezzoMedio = 0
                                ImportoScontiDaReg = 0
                                ImportoVarImballoDaReg = 0
                                ImportoCommissioniDaReg = 0
                                ImportoNettoIvaDaReg = PrezzoDaReg
                                Link_TMP_Prezzo_Medio = 0
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            Else
                                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                        Case 1
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            
                                        Case 2
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Case Else
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    End Select
                                Else
                                    If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            Link_TMP_Prezzo_Medio = 0
                                        Else
                                            If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                                PrezzoDaReg = PrezzoMedio
                                                ImportoScontiDaReg = ImportoScontiPM
                                                ImportoVarImballoDaReg = ImportoVarImballoPM
                                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                                Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            Else
                                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                                ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                                ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                                ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                                ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                                Link_TMP_Prezzo_Medio = 0
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            End If
                                        End If
                                    Else
                                        If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Else
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                        End If
                                    End If
                                End If
                            End If
                            TrattenutePerLavorazione = 0
                            TrattenuteGenerali = 0
                            TrattenuteTotali = 0
                            
                           'TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), (fnNotNullN(PrezzoDaReg) + fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                            
                            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                            sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto) "
                            
                            sSQL = sSQL & "VALUES ("
                            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                                sSQL = sSQL & 1 & ", "
                            Else
                                sSQL = sSQL & 4 & ", "
                            End If
                            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                            sSQL = sSQL & LINK_PERIODO & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                            sSQL = sSQL & fnNormString("F.A. n° " & fnNotNull(rsFatt!Doc_Prefisso) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(0) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
                            
                                Case 1
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                                Case 3
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            End Select
                            
                            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ")"
                            
                            
                            CnDMT.Execute sSQL
        
                            InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                            InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        End If
                    End If
                End If
            End If
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione F.A. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione F.A. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
                
            DoEvents
                            
                            
        rsFatt.MoveNext
        Wend
    End If
    rsFatt.Close
    Set rsFatt = Nothing


If (Unita_Progresso + prg.Value) >= prg.Max Then
    prg.Value = prg.Max
Else
    prg.Value = prg.Value + Unita_Progresso
End If
DoEvents

rsOgg.MoveNext
Wend
rsOgg.Close
Set rsOgg = Nothing
End Sub

Private Sub GET_VENDITA_SNF_SU_RICHIESTA(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Unita_Progresso As Double
Dim rsOgg As ADODB.Recordset
Dim NumeroRecord As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double

Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim AvviaLiquidazioneRiga As Long

FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE CORRISPETTIVI SELEZIONATI"

''''''''''''''''''''''CONTEGGIO DOCUMENTI DA ELABORARE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_POTMPLiqFattureSel) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=8 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

If rsOgg.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rsOgg!NumeroRecord)
End If

rsOgg.Close
Set rsOgg = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumeroRecord = 0 Then Exit Sub

Unita_Progresso = FormatNumber((prg.Max / NumeroRecord), 2)

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDTipoOggetto=8 "
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"

Set rsOgg = New ADODB.Recordset

rsOgg.Open sSQL, CnDMT.InternalConnection

While Not rsOgg.EOF
    'CORRISPETTIVI
    sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, "
    sSQL = sSQL & "Iva.IDIva AS IDIvaArticolo, Iva.Iva AS IvaArticolo, Iva.AliquotaIva AS AliquotaIvaArticolo, "
    sSQL = sSQL & "Iva.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0008.Doc_Prefisso, Articolo.RV_POIDCategoriaLiquidazione "
    sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
    sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
    sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoDettaglio0034.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoDettaglio0034.IDTipoOggetto = Oggetto.IDTipoOggetto ON "
    sSQL = sSQL & "Articolo.IDArticolo = ValoriOggettoDettaglio0034.Link_Art_articolo "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1 "
    sSQL = sSQL & " AND Oggetto.IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
    sSQL = sSQL & " AND Oggetto.IDTipoOggetto=" & fnNotNullN(rsOgg!IDTipoOggetto)

    If LINK_CAT_MERCE > 0 Then
        sSQL = sSQL & " AND Articolo.RV_POIDCategoriaLiquidazione=" & LINK_CAT_MERCE
    End If
    sSQL = sSQL & " ORDER BY RV_POIDSocio, doc_data"
    
    
    Set rsFatt = New ADODB.Recordset
    rsFatt.Open sSQL, CnDMT.InternalConnection
    
    If rsFatt.EOF = False Then
        
        While Not rsFatt.EOF
            AvviaLiquidazioneRiga = 1
            
            If NO_LIQ_VEND_UFF = 1 Then
                If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto)) = True) Then
                    AvviaLiquidazioneRiga = 0
                End If
            End If
            If AvviaLiquidazioneRiga = 1 Then
                If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
        
                    If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!RV_POIDConferimentoRighe)) = False Then
                        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
                            
                            ImportoScontiPM = 0
                            ImportoVarImballoPM = 0
                            ImportoCommissioniPM = 0
                            ImportoNettoIvaPM = 0
                            Link_TMP_Prezzo_Medio = 0
                            
                            PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento), fnNotNull(rsFatt!doc_data), fnNotNull(rsFatt!RV_PODataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                            If LINK_LISTINO > 0 Then
                                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(sFatt!Link_Art_Articolo), LINK_LISTINO)
                                PrezzoMedio = 0
                                ImportoScontiDaReg = 0
                                ImportoVarImballoDaReg = 0
                                ImportoCommissioniDaReg = 0
                                ImportoNettoIvaDaReg = PrezzoDaReg
                                Link_TMP_Prezzo_Medio = 0
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            Else
                                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                        Case 1
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            
                                        Case 2
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Case Else
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    End Select
                                Else
                                    If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            Link_TMP_Prezzo_Medio = 0
                                        Else
                                            If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                                PrezzoDaReg = PrezzoMedio
                                                ImportoScontiDaReg = ImportoScontiPM
                                                ImportoVarImballoDaReg = ImportoVarImballoPM
                                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                                Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            Else
                                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                                ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                                ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                                ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                                ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                                Link_TMP_Prezzo_Medio = 0
                                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            End If
                                        End If
                                    Else
                                        If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                        Else
                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                            ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                            Link_TMP_Prezzo_Medio = 0
                                        End If
                                    End If
                                End If
                            End If
                            TrattenutePerLavorazione = 0
                            TrattenuteGenerali = 0
                            TrattenuteTotali = 0
                            
                           'TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), (fnNotNullN(PrezzoDaReg) + fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNullN(rsFatt!IDCategoriaMerceologica), fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!RV_POQuantitaLiq), fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!RV_POInvenduto)
                            
                            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                            sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto) "
                            
                            sSQL = sSQL & "VALUES ("
                            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                                sSQL = sSQL & 1 & ", "
                            Else
                                sSQL = sSQL & 4 & ", "
                            End If
                            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                            sSQL = sSQL & LINK_PERIODO & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                            sSQL = sSQL & fnNormString("S.N.F. n° " & fnNotNull(rsFatt!Doc_Prefisso) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                            sSQL = sSQL & fnNotNullN(0) & ", "
                            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
                            
                                Case 1
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                                Case 3
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * rsFatt!RV_POQuantitaLiq) + (PrezzoMedio * fnNotNullN(rsFatt!RV_POQuantitaLiq))) & ", "
                            End Select
                            
                            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                            sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!RV_POQuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                            sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ")"
                            
                            CnDMT.Execute sSQL
        
                            InserisciNotaCredito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                            InserisciNotaDebito fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO
                        End If
                    End If
                End If
            End If
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione S.N.F. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione S.N.F. n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!RV_POSocio)
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
                
            DoEvents
                            
                            
        rsFatt.MoveNext
        Wend
    End If
    rsFatt.Close
    Set rsFatt = Nothing


If (Unita_Progresso + prg.Value) >= prg.Max Then
    prg.Value = prg.Max
Else
    prg.Value = prg.Value + Unita_Progresso
End If
DoEvents

rsOgg.MoveNext
Wend
rsOgg.Close
Set rsOgg = Nothing
End Sub

Private Function GET_ESISTENZA_CAMPIONATURA(IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCampionatura FROM RV_POCampionatura "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_CAMPIONATURA = False
Else
    GET_ESISTENZA_CAMPIONATURA = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_RIGA_PER_PREZZO_MEDIO(rsTmp As ADODB.Recordset, IDRiga As Long, IDTipoOggetto As Long, IDOggetto As Long, _
IDValoriOggettoDettaglio As Long, IDArticolo As Long, IDCategoriaMerceologica As Long, Quantita As Double, ImportoNettoVendita As Double, _
ImportoSconti As Double, ImportoVarImballo As Double, ImportoCommissioni As Double, ImportoLiquidazione As Double, _
DataVendita As String, DataConferimento As String, DataLavorazione As String, QuantitaDoc As Double)

rsTmp.AddNew
    rsTmp!IDRiga = IDRiga
    rsTmp!IDTipoOggetto = IDTipoOggetto
    rsTmp!IDOggetto = IDOggetto
    rsTmp!IDValoriOggettoDettaglio = IDValoriOggettoDettaglio
    rsTmp!IDArticolo = IDArticolo
    rsTmp!IDCategoriaMerceologica = IDCategoriaMerceologica
    rsTmp!Quantita = Quantita
    rsTmp!QuantitaDoc = QuantitaDoc
    rsTmp!ImportoNettoVendita = ImportoNettoVendita
    rsTmp!ImportoSconti = ImportoSconti
    rsTmp!ImportoVarImballo = ImportoVarImballo
    rsTmp!ImportoCommissioni = ImportoCommissioni
    rsTmp!ImportoLiquidazione = ImportoLiquidazione
    If Len(DataVendita) > 0 Then
        rsTmp!DataVendita = DataVendita
    End If
    If Len(DataConferimento) > 0 Then
        rsTmp!DataConferimento = DataConferimento
    End If
    If Len(DataLavorazione) > 0 Then
        rsTmp!DataLavorazione = DataLavorazione
    End If
rsTmp.Update

IDRiga = IDRiga + 1

End Sub
Private Sub SALVA_PREZZO_MEDIO_IN_CAMPIONATURA(IDRigaCampionatura As Long, Prezzo As Double, AliquotaIva As Double)
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCampionaturaRighe=" & IDRigaCampionatura

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rsNew.EOF Then
    rsNew!ImportoUnitario = Prezzo
    rsNew!ImportoNettoRiga = Prezzo * fnNotNullN(rsNew!QuantitaDefinitiva)
    rsNew!ImportoImpostaRiga = (rsNew!ImportoNettoRiga / 100) * AliquotaIva
    rsNew!ImportoLordoRiga = fnNotNullN(rsNew!ImportoNettoRiga) + fnNotNullN(rsNew!ImportoImpostaRiga)
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing

End Sub
Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POMoltiplicatore FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_MOLTIPLICATORE_ARTICOLO = 1
Else
    If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
        GET_MOLTIPLICATORE_ARTICOLO = 1
    Else
        GET_MOLTIPLICATORE_ARTICOLO = fnNotNullN(rs!RV_POMoltiplicatore)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_IMPORTO_DA_LISTINO(IDArticolo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDListino=" & IDListino

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_DA_LISTINO = 0
Else
    GET_IMPORTO_DA_LISTINO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_VENDITA_LIQ(IDValoriOggettoDettaglio As Long, IDOggetto As Long, IDTipoOggetto As Long, Optional Tipo As Long = 1, Optional IDRigaProcesso As Long = 0) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_VENDITA_LIQ = False

sSQL = "SELECT IDRV_POLiquidazioneRigheEla "
sSQL = sSQL & "FROM RV_POIELiquidazioneControlloVenditaLiq "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglioArticolo=" & IDValoriOggettoDettaglio
sSQL = sSQL & " AND IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND TipoRiga=" & Tipo
sSQL = sSQL & " AND Ufficiale=" & fnNormBoolean(1)
If (IDRigaProcesso > 0) Then
    sSQL = sSQL & " AND IDRV_POProcessoIVGammaRighe=" & IDRigaProcesso
End If
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_VENDITA_LIQ = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_UM_COOP(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_COOP = 0
Else
    GET_LINK_UM_COOP = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub AVVIA_CALCOLO_QTA_DA_ABBATTERE(TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim rsConf As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

FrmNuovoPeriodo.List1.Clear
FrmNuovoPeriodo.List1.AddItem "RECUPERO DATI PER CALCOLO QUANTITA' DA ABBATTERE..."
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
DoEvents

If TipoLiquidazione = 1 Then
    Set rsConf = New ADODB.Recordset
    rsConf.CursorLocation = adUseClient
    
    rsConf.Fields.Append "IDRigaConferimento", adInteger, , adFldIsNullable
    
    rsConf.Open , , adOpenKeyset, adLockBatchOptimistic
    
    ADD_CONF_PER_QTA_ABB rsConf, "RV_POIELiquidazioneControlloConfQtaAbbDDT", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB rsConf, "RV_POIELiquidazioneControlloConfQtaAbbFA", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB rsConf, "RV_POIELiquidazioneControlloConfQtaAbbSNF", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB rsConf, "RV_POIELiquidazioneControlloConfQtaAbbNC", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB rsConf, "RV_POIELiquidazioneControlloConfQtaAbbND", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_MIX rsConf, "RV_POIELiquidazioneControlloConfQtaAbbDDTMix", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_MIX rsConf, "RV_POIELiquidazioneControlloConfQtaAbbFAMix", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_MIX rsConf, "RV_POIELiquidazioneControlloConfQtaAbbSNFMix", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_MIX rsConf, "RV_POIELiquidazioneControlloConfQtaAbbNCMix", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_MIX rsConf, "RV_POIELiquidazioneControlloConfQtaAbbNDMix", TipoLiquidazione, DataInizio, DataFine
    ADD_CONF_PER_QTA_ABB_SCARTO rsConf, "RV_POIELiquidazioneControlloConfQtaAbbScarto", TipoLiquidazione, DataInizio, DataFine
    
    If ((rsConf.EOF) And (rsConf.BOF)) Then Exit Sub
    
    rsConf.MoveFirst
    
    While Not rsConf.EOF
        CALCOLO_CONF_QTA_ABB fnNotNullN(rsConf!IDRigaConferimento), TIPO_RAGGR_ABB
    rsConf.MoveNext
    Wend
    
    rsConf.Close
    Set rsConf = Nothing
End If
If TipoLiquidazione = 2 Then
    CALCOLO_CONF_QTA_ABB_VEND 2, "RV_POIELiquidazioneControlloConfQtaAbbDDT", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 114, "RV_POIELiquidazioneControlloConfQtaAbbFA", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 8, "RV_POIELiquidazioneControlloConfQtaAbbSNF", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 11, "RV_POIELiquidazioneControlloConfQtaAbbNC", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 107, "RV_POIELiquidazioneControlloConfQtaAbbND", TIPO_RAGGR_ABB, DataInizio, DataFine

    CALCOLO_CONF_QTA_ABB_VEND 2, "RV_POIELiquidazioneControlloConfQtaAbbDDTMix", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 114, "RV_POIELiquidazioneControlloConfQtaAbbFAMix", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 8, "RV_POIELiquidazioneControlloConfQtaAbbSNFMix", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 11, "RV_POIELiquidazioneControlloConfQtaAbbNCMix", TIPO_RAGGR_ABB, DataInizio, DataFine
    CALCOLO_CONF_QTA_ABB_VEND 107, "RV_POIELiquidazioneControlloConfQtaAbbNDMix", TIPO_RAGGR_ABB, DataInizio, DataFine
    
    CALCOLO_QTA_ABB_VEND_SCARTO TIPO_RAGGR_ABB, DataInizio, DataFine
End If
End Sub
Private Sub ADD_CONF_PER_QTA_ABB(rsConf As ADODB.Recordset, NomeTabella As String, TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT RV_POIDConferimentoRighe FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazione>0 "
'sSQL = sSQL & " AND IDTipoDocumentoCoop=1 "


If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
End If
If TipoLiquidazione = 2 Then
    sSQL = sSQL & " AND DataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataCompetenzaLiq<=" & fnNormDate(DataFine)
End If

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    If fnNotNullN(rs!RV_POIDConferimentoRighe) > 0 Then
        rsConf.Filter = "IDRigaConferimento=" & fnNotNullN(rs!RV_POIDConferimentoRighe)
        If rsConf.EOF Then
            rsConf.AddNew
                rsConf!IDRigaConferimento = fnNotNullN(rs!RV_POIDConferimentoRighe)
            rsConf.Update
        End If
        rsConf.Filter = vbNullString
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ADD_CONF_PER_QTA_ABB_MIX(rsConf As ADODB.Recordset, NomeTabella As String, TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazione>0 "
sSQL = sSQL & " AND RV_POIDTipoUtilizzoLinea=4 "
sSQL = sSQL & " AND RV_POTipoRiga=1 "
'sSQL = sSQL & " AND IDTipoDocumentoCoop=1"
If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND DataLiquidazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataLiquidazione<=" & fnNormDate(DataFine)
End If
If TipoLiquidazione = 2 Then
    sSQL = sSQL & " AND DataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataCompetenzaLiq<=" & fnNormDate(DataFine)
End If

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    If fnNotNullN(rs!IDRV_POCaricoMerceRighe) > 0 Then
        rsConf.Filter = "IDRigaConferimento=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
        If rsConf.EOF Then
            rsConf.AddNew
                rsConf!IDRigaConferimento = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
            rsConf.Update
        End If
        rsConf.Filter = vbNullString
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ADD_CONF_PER_QTA_ABB_SCARTO(rsConf As ADODB.Recordset, NomeTabella As String, TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazione>0 "
'sSQL = sSQL & " AND IDTipoDocumentoCoop=1"

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND DataDocumentoConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataDocumentoConferimento<=" & fnNormDate(DataFine)
End If
If TipoLiquidazione = 2 Then
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DataFine)
End If

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    If fnNotNullN(rs!IDRV_POCaricoMerceRighe) > 0 Then
        rsConf.Filter = "IDRigaConferimento=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
        If rsConf.EOF Then
            rsConf.AddNew
                rsConf!IDRigaConferimento = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
            rsConf.Update
        End If
        rsConf.Filter = vbNullString
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CALCOLO_CONF_QTA_ABB(IDRigaConferimento As Long, TipoRaggruppamento As Long)
On Error GoTo ERR_CALCOLO_CONF_QTA_ABB
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset

Dim PercAbbConferimento As Double
Dim PercAbbConferimentoScarto As Double

Dim TotaleVenduto As Double
Dim TotaleQuadratura As Double
Dim TotaleConferimentoElaborato As Double
Dim TotaleQuadraturaAbbattuta As Double
Dim TotaleConferimentoVenduto As Double
Dim DescrizioneConferimento As String
Dim TotaleProcesso As Double
Dim QuantitaSingoloProcesso As Double
Dim IDUMCoopProcesso As Long
Dim AvviaLiquidazioneRiga As Long
Dim IDRiferimentoRaggruppamento As Long
Dim IDAnagraficaSocio As Long

TotaleQuadratura = 0
TotaleVenduto = 0
TotaleConferimentoElaborato = 0
TotaleQuadraturaAbbattuta = 0
TotaleConferimentoVenduto = 0
PercAbbConferimento = 0
PercAbbConferimentoScarto = 0
DescrizioneConferimento = ""
Dim AvviaCalcolo As Boolean

'ELIMINAZIONE DATI QUANTITA ABBATTUTA DEL CONFERIMENTO
EliminaQtaAbbConf IDRigaConferimento

sSQL = "SELECT RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe, RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta, RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, "
sSQL = sSQL & "Articolo.RV_POPercentualeAbbattimentoLiquidazione, Articolo.RV_POPercentualeAbbattimentoLiquidazioneScarto, RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceTesta.DataDocumento,"
sSQL = sSQL & "RV_POCaricoMerceTesta.IDAnagrafica , RV_POCaricoMerceTesta.Anagrafica, RV_POCaricoMerceTesta.Nome, "
sSQL = sSQL & "Articolo.IDCategoriaFiscale, Articolo.IDClassificazioneArticolo, Articolo.IDCategoriaMerceologica, Articolo.IDRepartoFiscale, "
sSQL = sSQL & "Articolo.RV_POIDGruppoArticoloPerEvasioneMix, Articolo.RV_PO01_IDFamigliaProdotti, Articolo.RV_PO01_IDVarieta, Articolo.RV_POIDCategoriaLiquidazione, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDTipoDocumentoCoop "
sSQL = sSQL & " FROM RV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & " Articolo ON RV_POCaricoMerceRighe.IDArticolo = Articolo.IDArticolo INNER JOIN "
sSQL = sSQL & " RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = CnDMT.OpenResultset(sSQL)
AvviaCalcolo = True

If Not rs.EOF Then
    
    If AvviaCalcolo = False Then Exit Sub
    
    PercAbbConferimento = fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazione)
    PercAbbConferimentoScarto = fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto)
    IDAnagraficaSocio = fnNotNullN(rs!IDAnagrafica)
    
    DescrizioneConferimento = "Elaborazione conferimento del socio/fornitore " & fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome) & " N° " & fnNotNullN(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento)
    
    Select Case TipoRaggruppamento
        
        Case 1
            IDRiferimentoRaggruppamento = fnNotNullN(rs!IDCategoriaFiscale)
        Case 2
            IDRiferimentoRaggruppamento = fnNotNullN(rs!IDClassificazioneArticolo)
        Case 3
            IDRiferimentoRaggruppamento = fnNotNullN(rs!IDCategoriaMerceologica)
        Case 4
            IDRiferimentoRaggruppamento = fnNotNullN(rs!IDRepartoFiscale)
        Case 5
            IDRiferimentoRaggruppamento = fnNotNullN(rs!RV_POIDGruppoArticoloPerEvasioneMix)
        Case 6
            IDRiferimentoRaggruppamento = fnNotNullN(RV_PO01_IDFamigliaProdotti)
        Case 7
            IDRiferimentoRaggruppamento = fnNotNullN(rs!RV_PO01_IDVarieta)
        Case 8
            IDRiferimentoRaggruppamento = fnNotNullN(rs!RV_POIDCategoriaLiquidazione)
        Case Else
            IDRiferimentoRaggruppamento = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    End Select
End If
'If fnNotNullN(rs!RV_POIDCategoriaLiquidazione) = 7 Then
'    MsgBox "STOP"
'End If
rs.CloseResultset
Set rs = Nothing

FrmNuovoPeriodo.List1.AddItem DescrizioneConferimento
FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
DoEvents

TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimento(2, "RV_POIELiquidazioneControlloConfQtaAbbDDT", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimento(114, "RV_POIELiquidazioneControlloConfQtaAbbFA", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimento(8, "RV_POIELiquidazioneControlloConfQtaAbbSNF", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimento(11, "RV_POIELiquidazioneControlloConfQtaAbbNC", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimento(107, "RV_POIELiquidazioneControlloConfQtaAbbND", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimentoMix(2, "RV_POIELiquidazioneArticoliMixDDT", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimentoMix(114, "RV_POIELiquidazioneArticoliMixFA", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimentoMix(8, "RV_POIELiquidazioneArticoliMixSNF", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimentoMix(11, "RV_POIELiquidazioneArticoliMixNC", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)
TotaleVenduto = TotaleVenduto + GetTotaleVendutoConferimentoMix(107, "RV_POIELiquidazioneArticoliMixND", IDAnagraficaSocio, IDRiferimentoRaggruppamento, TipoRaggruppamento)

TotaleQuadratura = GetTotaleScartoConferimento(IDRiferimentoRaggruppamento, IDAnagraficaSocio, TotaleQuadraturaAbbattuta, TipoRaggruppamento)

TotaleConferimentoElaborato = (TotaleVenduto + TotaleQuadratura) - (((TotaleVenduto + TotaleQuadratura) / 100) * PercAbbConferimento)

If PercAbbConferimentoScarto > 0 Then
    TotaleQuadraturaAbbattuta = (TotaleQuadratura) - (((TotaleQuadratura) / 100) * PercAbbConferimentoScarto)
Else
    TotaleQuadraturaAbbattuta = TotaleQuadraturaAbbattuta
End If


TotaleConferimentoVenduto = TotaleConferimentoElaborato - TotaleQuadraturaAbbattuta

If TotaleVenduto = 0 Then Exit Sub

sSQL = "SELECT * FROM RV_POLiquidazioneConfQtaAbb "
sSQL = sSQL & "WHERE ID=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

'DOCUMENTO DI TRASPORTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbDDT"
sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'FATTURA ACCOMPAGNATORIA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbFA"
sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CORRISPETTIVO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbSNF"
sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)


While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NOTA DI CREDITO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbNC"
sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!Quantita = rsNew!Quantita * -1
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq) * -1
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NOTA DI DEBITO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbND"
sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)


While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'SCARTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbScarto"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDRV_POLavorazione), 0, 0, 2) = True) Then
        AvviaLiquidazioneRiga = 0
        End If
    End If
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POLavorazione)
            rsNew!IDOggetto = 0
            rsNew!IDTipoOggetto = 0
            rsNew!TipoRigaLiquidazione = 2
            
            Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
                Case 1
                    rsNew!Quantita = fnNotNullN(rs!Colli)
                Case 2
                    rsNew!Quantita = fnNotNullN(rs!PesoLordo)
                Case 3
                    rsNew!Quantita = fnNotNullN(rs!PesoNetto)
                Case 4
                    rsNew!Quantita = fnNotNullN(rs!Tara)
                Case 5
                    rsNew!Quantita = fnNotNullN(rs!Pezzi)
                Case Else
                    rsNew!Quantita = fnNotNullN(rs!Qta_UM)
            End Select
            rsNew!QuantitaOriginale = rsNew!Quantita
            If PercAbbConferimentoScarto > 0 Then
                rsNew!Quantita = rsNew!Quantita - ((rsNew!Quantita / 100) * PercAbbConferimentoScarto)
            Else
                rsNew!Quantita = rsNew!Quantita - ((rsNew!Quantita / 100) * fnNotNullN(rs!PercentualeAbbattimentoScarto))
            End If
        
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'DOCUMENTO DI TRASPORTO ARTICOLO MIX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbDDTMix"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 1, fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If IDUMCoopProcesso = 0 Then
            IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        End If
        TotaleProcesso = 0
        QuantitaSingoloProcesso = 0
        
        Select Case IDUMCoopProcesso
            Case 1
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleProcesso > 0) Then QuantitaSingoloProcesso = (fnNotNullN(rs!Colli) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 2
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoLordo) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 3
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoNetto) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 4
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Tara) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 5
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Pezzi) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
        End Select
        
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 4
            rsNew!Quantita = (QuantitaSingoloProcesso / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = QuantitaSingoloProcesso
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'FATTURA ACCOMPAGNATORIA ARTICOLO MIX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbFAMix"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 1, fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If IDUMCoopProcesso = 0 Then
            IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        End If
        TotaleProcesso = 0
        QuantitaSingoloProcesso = 0
        
        Select Case IDUMCoopProcesso
            Case 1
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleProcesso > 0) Then QuantitaSingoloProcesso = (fnNotNullN(rs!Colli) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 2
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoLordo) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 3
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoNetto) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 4
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Tara) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 5
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Pezzi) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
        End Select
        
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 4
            rsNew!Quantita = (QuantitaSingoloProcesso / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = QuantitaSingoloProcesso
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CORRISPETTIVI ARTICOLO MIX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbSNFMix"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 1, fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If IDUMCoopProcesso = 0 Then
            IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        End If
        TotaleProcesso = 0
        QuantitaSingoloProcesso = 0
        
        Select Case IDUMCoopProcesso
            Case 1
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleProcesso > 0) Then QuantitaSingoloProcesso = (fnNotNullN(rs!Colli) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 2
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoLordo) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 3
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoNetto) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 4
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Tara) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 5
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Pezzi) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
        End Select
        
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 4
            rsNew!Quantita = (QuantitaSingoloProcesso / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = QuantitaSingoloProcesso
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NOTA DI CREDITO ARTICOLO MIX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbNCMix"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 1, fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If IDUMCoopProcesso = 0 Then
            IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        End If
        TotaleProcesso = 0
        QuantitaSingoloProcesso = 0
        
        Select Case IDUMCoopProcesso
            Case 1
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleProcesso > 0) Then QuantitaSingoloProcesso = (fnNotNullN(rs!Colli) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 2
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoLordo) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 3
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoNetto) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 4
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Tara) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 5
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Pezzi) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
        End Select
        
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 4
            rsNew!Quantita = (QuantitaSingoloProcesso / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!Quantita = rsNew!Quantita * -1
            rsNew!QuantitaOriginale = QuantitaSingoloProcesso * -1
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NOTA DI DEBITO ARTICOLO MIX'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbNDMix"
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND RV_POTipoRIga=1"
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 1, fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If IDUMCoopProcesso = 0 Then
            IDUMCoopProcesso = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        End If
        TotaleProcesso = 0
        QuantitaSingoloProcesso = 0
        
        Select Case IDUMCoopProcesso
            Case 1
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleProcesso > 0) Then QuantitaSingoloProcesso = (fnNotNullN(rs!Colli) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 2
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoLordo) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 3
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!PesoNetto) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 4
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Tara) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
            Case 5
                TotaleProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If TotaleProcesso > 0 Then QuantitaSingoloProcesso = (fnNotNullN(rs!Pezzi) / TotaleProcesso) * fnNotNullN(rs!RV_POQuantitaLiq)
        End Select
        
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 4
            rsNew!Quantita = (QuantitaSingoloProcesso / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = QuantitaSingoloProcesso
        rsNew.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_CALCOLO_CONF_QTA_ABB:
    MsgBox Err.Description, vbCritical, "CALCOLO_CONF_QTA_ABB (" & IDRigaConferimento & ")"
End Sub
Private Function GetTotaleVendutoConferimentoMix(IDTipoOggetto As Long, NomeTabella As String, IDAnagraficaSocio As Long, IDRiferimentoGruppo As Long, Optional TipoRaggruppamento As Long = 0) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TotaleQtaProcesso As Double
Dim TotaleQtaRigaVendita As Double
Dim QtaSingolaRiga As Double
Dim IDUMCoop As Long
Dim AvviaLiquidazioneRiga As Long
Dim IDSocio As Long

GetTotaleVendutoConferimentoMix = 0
IDSocio = 0

sSQL = "SELECT * "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND RV_POIDTipoUtilizzoLinea=4"
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"

Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND RV_POIDConferimentoRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If

If ((IDTipoOggetto = 11) Or (IDTipoOggetto = 107)) Then
    sSQL = sSQL & " AND ((RV_POIDTipoVariazione=2) OR (RV_POIDTipoVariazione=3))"
End If
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND doc_data>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND doc_data<=" & fnNormDate(DATA_FINE)
End If
Set rs = CnDMT.OpenResultset(sSQL)

IDUMCoop = 3

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    
    If AvviaLiquidazioneRiga = 1 Then
        TotaleQtaProcesso = 0
        QtaSingolaRiga = 0
        TotaleQtaRigaVendita = fnNotNullN(rs!RV_POQuantitaLiq)
        
        IDUMCoop = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If (IDUMCoop = 0) Then
            IDUMCoop = fnNotNullN(rs!IDUnitaDiMisuraCoopVendita)
        End If
        
        
        Select Case IDUMCoop
            Case 1
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Colli) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 2
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!PesoLordo) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 3
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!PesoNetto) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 4
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Tara) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 5
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Pezzi) / TotaleQtaProcesso) * TotaleQtaRigaVendita
        End Select
        
        GetTotaleVendutoConferimentoMix = GetTotaleVendutoConferimentoMix + QtaSingolaRiga
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If GetTotaleVendutoConferimentoMix > 0 Then
    If IDTipoOggetto = 11 Then GetTotaleVendutoConferimentoMix = GetTotaleVendutoConferimentoMix * -1
End If
End Function
Private Function GetTotaleProcesso(IDProcesso As Long, NomeCampoSomma As String) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GetTotaleProcesso = 0

sSQL = "SELECT SUM(" & NomeCampoSomma & ") AS Totale "
sSQL = sSQL & " FROM RV_POProcessoIVGammaRighe "
sSQL = sSQL & " WHERE IDRV_POProcessoIVGamma=" & IDProcesso

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GetTotaleProcesso = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GetTotaleVendutoConferimento(IDTipoOggetto As Long, NomeTabella As String, IDAnagraficaSocio As Long, IDRiferimentoGruppo As Long, Optional TipoRaggruppamento As Long = 0) As Double
On Error GoTo ERR_GetTotaleVendutoConferimento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaLiquidazioneRiga As Long
Dim IDSocio As Long

GetTotaleVendutoConferimento = 0
IDSocio = 0
'sSQL = "SELECT SUM(RV_POQuantitaLiq) as Totale "
sSQL = "SELECT * "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"
Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND RV_POIDConferimentoRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If

If ((IDTipoOggetto = 11) Or (IDTipoOggetto = 107)) Then
    sSQL = sSQL & " AND ((RV_POIDTipoVariazione=2) OR (RV_POIDTipoVariazione=3))"
End If
If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate(DATA_FINE)
End If

Set rs = CnDMT.OpenResultset(sSQL)

'If Not rs.EOF Then
'    GetTotaleVendutoConferimento = fnNotNullN(rs!Totale)
'End If

While Not rs.EOF
        AvviaLiquidazioneRiga = 1
        If NO_LIQ_VEND_UFF = 1 Then
            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
                AvviaLiquidazioneRiga = 0
            End If
        End If
        If AvviaLiquidazioneRiga = 1 Then
            GetTotaleVendutoConferimento = GetTotaleVendutoConferimento + fnNotNullN(rs!RV_POQuantitaLiq)
        End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If GetTotaleVendutoConferimento > 0 Then
    If IDTipoOggetto = 11 Then GetTotaleVendutoConferimento = GetTotaleVendutoConferimento * -1
End If
Exit Function
ERR_GetTotaleVendutoConferimento:
    MsgBox Err.Description, vbCritical, "GetTotaleVendutoConferimento"
End Function
Private Function GetTotaleScartoConferimento(IDRiferimentoGruppo As Long, IDAnagraficaSocio As Long, TotaleQuadraturaAbbattuta As Double, Optional TipoRaggruppamento As Long = 0) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaTotale As Double
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaSingolo As Double
Dim IDSocio As Long

QuantitaTotale = 0
TotaleQuadraturaAbbattuta = 0
IDSocio = 0

sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbScarto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND PreConferimento=0"
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"

Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If

If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
End If

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDRV_POLavorazione), 0, 0, 2) = True) Then
        AvviaLiquidazioneRiga = 0
        End If
    End If

    If AvviaLiquidazioneRiga = 1 Then
        Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
            Case 1
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Colli)
                QuantitaSingolo = fnNotNullN(rs!Colli)
            Case 2
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!PesoLordo)
                QuantitaSingolo = fnNotNullN(rs!PesoLordo)
            Case 3
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!PesoNetto)
                QuantitaSingolo = fnNotNullN(rs!PesoNetto)
            Case 4
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Tara)
                QuantitaSingolo = fnNotNullN(rs!Tara)
            Case 5
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Pezzi)
                QuantitaSingolo = fnNotNullN(rs!Pezzi)
            Case Else
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Qta_UM)
                QuantitaSingolo = fnNotNullN(rs!Qta_UM)
        End Select
        QuantitaTotale = QuantitaTotale + QuantitaSingolo
        
'        If (fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto) > 0) Then
'            QuantitaSingolo = QuantitaSingolo - ((QuantitaSingolo / 100) * fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto))
'        Else
            If fnNotNullN(rs!PercentualeAbbattimentoScarto) > 0 Then
                QuantitaSingolo = QuantitaSingolo - ((QuantitaSingolo / 100) * fnNotNullN(rs!PercentualeAbbattimentoScarto))
            End If
'        End If
        TotaleQuadraturaAbbattuta = TotaleQuadraturaAbbattuta + QuantitaSingolo
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

GetTotaleScartoConferimento = QuantitaTotale

End Function
Private Function GET_QUANTITA_ABBATTUTA(IDRigaConferimento As Long, IDValoriOggettoDettaglio As Long, IDOggetto As Long, IDTipoOggetto As Long, TipoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_QUANTITA_ABBATTUTA = 0

sSQL = "SELECT ID, Quantita "
sSQL = sSQL & "FROM RV_POLiquidazioneConfQtaAbb "
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
sSQL = sSQL & " AND IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND TipoRigaLiquidazione=" & TipoRiga

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_QUANTITA_ABBATTUTA = fnNotNullN(rs!Quantita)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_VENDITA_MIX(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long, DescrizioneTipoOggetto As String, NomeTabella As String)
Dim sSQL As String
Dim rsFatt As ADODB.Recordset
Dim Link_CategoriaMerceologica As Long

Dim PrezzoMedio As Double
Dim PrezzoDaReg As Double

Dim ImportoScontiPM As Double
Dim ImportoScontiDaReg As Double

Dim ImportoVarImballoPM As Double
Dim ImportoVarImballoDaReg As Double

Dim ImportoCommissioniPM As Double
Dim ImportoCommissioniDaReg As Double

Dim ImportoNettoIvaPM As Double
Dim ImportoNettoIvaDaReg As Double

Dim Link_TMP_Prezzo_Medio As Long

Dim Quantita_Venduta_Per_Variazione As Double
Dim Unita_Progresso As Double
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiq As Double
Dim IDUMCoopVendita As Long
Dim TotaleProcesso As Double
Dim QuantitaLiqAbb As Double
Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double


FrmNuovoPeriodo.List1.AddItem "ELABORAZIONE " & DescrizioneTipoOggetto & " PER PRODOTTI MIX..."
DoEvents

sSQL = "SELECT * FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1 "
sSQL = sSQL & " AND RV_POIDTipoUtilizzoLinea=4 "
If LINK_CAT_MERCE > 0 Then
    sSQL = sSQL & " AND IDCategoriaLiquidazioneArticoloMix=" & LINK_CAT_MERCE
End If

Select Case TipoLiquidazione
    Case 1
        sSQL = sSQL & " AND DataLiquidazione>=" & fnNormDate(DATA_INIZIO)
        sSQL = sSQL & " AND DataLiquidazione<=" & fnNormDate(DATA_FINE)
        If LINK_SOCIO > 0 Then
            sSQL = sSQL & " AND IDAnagraficaSocio=" & LINK_SOCIO
        End If
    Case 2
        sSQL = sSQL & " AND DataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO)
        sSQL = sSQL & " AND DataCompetenzaLiq<=" & fnNormDate(DATA_FINE)
        If LINK_SOCIO > 0 Then
            sSQL = sSQL & " AND IDAnagraficaSocio=" & LINK_SOCIO
        End If
    Case 3
End Select


sSQL = sSQL & " ORDER BY IDAnagraficaSocio, doc_data"

Set rsFatt = New ADODB.Recordset
rsFatt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsFatt.EOF = False Then
    prg.Value = 0
    prg.Max = 1000
    Unita_Progresso = prg.Max / rsFatt.RecordCount
    
    While Not rsFatt.EOF
        FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione " & fnNotNull(rsFatt!Oggetto) & " n° " & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data) & " - Socio: " & fnNotNull(rsFatt!AnagraficaSocio) & " " & fnNotNull(rsFatt!NomeAnagraficaSocio)
        FrmNuovoPeriodo.List1.AddItem "- " & FrmNuovoPeriodo.lblInfoStatus.Caption
        FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
        AvviaLiquidazioneRiga = 1
        If NO_LIQ_VEND_UFF = 1 Then
            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 1, fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe)) = True) Then
                AvviaLiquidazioneRiga = 0
            End If
        End If
    
        If AvviaLiquidazioneRiga = 1 Then
            If GET_ESISTENZA_CAMPIONATURA(fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe)) = False Then
                If GET_FLAG_NON_LIQUIDARE(fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe)) = False Then
                    
                    If fnNotNullN(rsFatt!IDArticolo) > 0 Then
                        
                        ImportoScontiPM = 0
                        ImportoVarImballoPM = 0
                        ImportoCommissioniPM = 0
                        ImportoNettoIvaPM = 0
                        Link_TMP_Prezzo_Medio = 0
                        
                        If LINK_LISTINO = 0 Then
                            PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!IDAnagraficaSocio), fnNotNullN(rsFatt!IDArticolo), fnNotNull(rsFatt!DataLiquidazione), fnNotNull(rsFatt!RV_PODataCompetenzaLiq), fnNotNull(rsFatt!DataLavorazione), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio)
                        End If
                                            
                        If LINK_LISTINO > 0 Then
                            PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!Link_Art_Articolo), LINK_LISTINO)
                            ImportoScontiDaReg = 0
                            PrezzoMedio = 0
                            ImportoVarImballoDaReg = 0
                            ImportoCommissioniDaReg = 0
                            ImportoNettoIvaDaReg = PrezzoDaReg
                            Link_TMP_Prezzo_Medio = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Else
                            If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                                Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                                    Case 1
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        Link_TMP_Prezzo_Medio = 0
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                    Case 2
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                    Case Else
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                End Select
                            Else
                                If TIPO_IMPORTO_ARTICOLO = 1 Then 'TIPO PREZZO DI VENDITA
                                    If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then 'PREZZO DI VENDITA
                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                                        Link_TMP_Prezzo_Medio = 0
                                    Else
                                        'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                            PrezzoDaReg = PrezzoMedio
                                            ImportoScontiDaReg = ImportoScontiPM
                                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                            ImportoCommissioniDaReg = ImportoCommissioniPM
                                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                            Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
'                                        Else
'                                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'                                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                            Link_TMP_Prezzo_Medio = 0
'                                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        End If
                                    End If
                                Else
                                    'If fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq) = 1 Then
                                        PrezzoDaReg = PrezzoMedio
                                        ImportoScontiDaReg = ImportoScontiPM
                                        'ImportoVarImballoDaReg = ImportoVarImballoPM
                                        ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                        ImportoCommissioniDaReg = ImportoCommissioniPM
                                        ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                        Link_TMP_Prezzo_Medio = Link_TMP_Prezzo_Medio
                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 3
'                                    Else
'                                        PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                        TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                        ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                                        ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                        'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                                        ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                                        ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                                        Link_TMP_Prezzo_Medio = 0
                                    'End If
                                End If
                            End If
                        End If
                        
                        TrattenutePerLavorazione = 0
                        TrattenuteGenerali = 0
                        TrattenuteTotali = 0
    
                        TrattValGen1 = 0
                        TrattValGen2 = 0
                        TrattPercGen1 = 0
                        TrattPercGen2 = 0
                        
                        TrattValLav1 = 0
                        TrattValLav2 = 0
                        TrattPercLav1 = 0
                        TrattPercLav2 = 0
                        
                        TrattValPreLiq1 = 0
                        TrattValPreLiq2 = 0
                        
                        IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopArticoloMix)
                        If (IDUMCoop = 0) Then
                            IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopVendita)
                        End If
                        Select Case IDUMCoop
                            Case 1
                                TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Colli")
                                If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Colli) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            Case 2
                                TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoLordo")
                                If (TotaleQtaProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoLordo) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            Case 3
                                TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoNetto")
                                If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoNetto) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            Case 4
                                TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Tara")
                                If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Tara) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            Case 5
                                TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Pezzi")
                                If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Pezzi) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                            
                        End Select
                        
                        If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                            QuantitaLiqAbb = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe), fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 4)
                            If QuantitaLiqAbb > 0 Then
                                QuantitaLiq = QuantitaLiqAbb
                            End If
                        End If
                        
                        If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
                        If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

                        
                        TrattenuteArticolo fnNotNullN(rsFatt!IDAnagraficaSocio), fnNotNullN(rsFatt!IDArticolo), fnNotNullN(rsFatt!IDCategoriaMerceologicaArticoloLavorato), fnNotNullN(rsFatt!IDTipoLavorazioneArticoloLavorato), False, QuantitaLiq, fnNotNullN(PrezzoDaReg), fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe), fnNotNullN(rsFatt!IDTipoOggetto), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe), fnNotNullN(rsFatt!RV_POInvenduto)

                        sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                        sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
                        sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
                        sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                        sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                        sSQL = sSQL & "IDTipoOggetto, IDRV_POTipoOggettoVariante, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                        sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                        sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                        sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                        sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                        sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali, "
                        sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato, "
                        sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, ImportoScontiDaReg, "
                        sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, ImportoVarImpImballoDaReg, "
                        sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, ImportoCommissioniDaReg, "
                        sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, ImpUniVendDocNettoIvaVenditaDaReg, "
                        sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, "
                        sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                        sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2, "
                        sSQL = sSQL & "TrattenutaValorePreLiq1, TrattenutaValorePreLiq2, IDRV_POProcessoIVGammaRighe, "
                        sSQL = sSQL & "IDArticoloMixVenduto, CodiceArticoloMixVenduto, DescrizioneArticoloMixVenduto"
                        sSQL = sSQL & ") "
                        
                        sSQL = sSQL & "VALUES ("
                        If fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe) > 0 Then
                            sSQL = sSQL & 1 & ", "
                        Else
                            sSQL = sSQL & 4 & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!IDArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Articolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDTipoLavorazioneArticoloLavorato) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDRV_POTipoCategoriaArticoloLavorato) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDRV_POCalibroArticoloLavorato) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!DataLiquidazione) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDAnagraficaSocio) & ", "
                        sSQL = sSQL & LINK_PERIODO & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                        sSQL = sSQL & fnNormString(fnNotNull(rsFatt!Oggetto) & " n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                        sSQL = sSQL & fnNotNullN(0) & ", "
                        sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                        
                        TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 5)
                        TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 5)
                        TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
                        
                        If LINK_LISTINO > 0 Then
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                        Else
'                            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
'
'                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
'                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
'                            End Select
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq, 2)) & ", "
'                            sSQL = sSQL & fnNormNumber(FormatNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq), 2)) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                             
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologicaArticoloLavorato) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ", "
                        sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
                        sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ", "
                        If LINK_LISTINO = 0 Then
                            'IMPORTO SCONTI
                            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO VARIAZIONE IMBALLO
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO COMMISSIONI
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            'IMPORTO UNITARIO NETTO IVA
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
                            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
                            sSQL = sSQL & Link_TMP_Prezzo_Medio & ", "
                            sSQL = sSQL & LINK_TIPO_PREZZO_MEDIO_ARTICOLO & ", "
                        End If
                        sSQL = sSQL & fnNotNullN(rsFatt!RV_POInvenduto) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattPercLav2) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq1) & ", "
                        sSQL = sSQL & fnNormNumber(TrattValPreLiq2) & ","
                        sSQL = sSQL & fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe) & ", "
                        sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
                        sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione)
                        sSQL = sSQL & ")"
                        
                        CnDMT.Execute sSQL
                        
                        InserisciNotaCredito_MIX fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, TotaleProcesso, fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe)
                        InserisciNotaDebito_MIX fnNotNullN(rsFatt!IDValoriOggettoDettaglio), fnNotNullN(rsFatt!IDTipoOggetto), PrezzoMedio, fnNotNullN(rsFatt!RV_POPrezzoMedioInLiq), fnNotNull(rsFatt!RV_POCodiceLotto), ImportoScontiPM, ImportoVarImballoPM, ImportoCommissioniPM, ImportoNettoIvaPM, Link_TMP_Prezzo_Medio, LINK_TIPO_PREZZO_MEDIO_ARTICOLO, TotaleProcesso, fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe)
                        
                    End If
                End If
            End If
        End If

        If (Unita_Progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If

        DoEvents
                      
    rsFatt.MoveNext
    Wend
End If
rsFatt.Close
Set rsFatt = Nothing

End Sub


Private Sub InserisciNotaCredito_MIX(ValoriOggettoDettaglio As Long, IDTipoOggetto As Long, PrezzoMedio As Double, PrezzoMedioInFattura As Long, CodiceLottoVendita As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, ImportoNettoIvaPM As Double, Link_TMP_Prezzo_Medio As Long, IDTipoPrezzoMedio As Long, TotaleProcesso As Double, IDProcessoIVGammaRighe As Long)
Dim LINK_TIPO_CALCOLO_NC As Long
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim PrezzoDaReg As Double
Dim ImportoScontiDaReg As Double
Dim ImportoVarImballoDaReg As Double
Dim ImportoCommissioniDaReg As Double
Dim ImportoNettoIvaDaReg As Double
Dim Link_Riga_Prezzo_Medio_TMP As Long
Dim QuantitaLiq As Double

Dim IDUMCoopVendita As Long
'Dim TotaleProcesso As Double
Dim QuantitaLiqAbb As Double
Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double
Dim LiquidaAPrezzoMedio As Boolean

'NOTA DI CREDITO
sSQL = "SELECT * FROM RV_POIELiquidazioneArticoliMixNC "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND IDRV_POProcessoIVGammaRighe=" & IDProcessoIVGammaRighe
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio 'Tolto per incongruenza con la riga del documento
End If

Set rsFatt = CnDMT.OpenResultset(sSQL)

While Not rsFatt.EOF
    'If (rsFatt!IDOggetto = 188224) And (rsFatt!IDValoriOggettoDettaglio = 7971) Then
    '    MsgBox "STOP"
    'End If

    If fnNotNullN(rsFatt!IDArticolo) > 0 Then
            LiquidaAPrezzoMedio = False
            If LINK_LISTINO > 0 Then
                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!IDArticolo), LINK_LISTINO)
                ImportoScontiDaReg = 0
                'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                ImportoVarImballoDaReg = 0
                
                ImportoCommissioniDaReg = 0
                ImportoNettoIvaDaReg = 0
                Link_Riga_Prezzo_Medio_TMP = 0
                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
            
            Else
            
                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                        Case 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Case 2
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                        Case Else
                            PrezzoDaReg = PrezzoMedio
                            
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            
                            
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                    End Select
                Else
                    If TIPO_IMPORTO_ARTICOLO = 1 Then
                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                        Else
                            'If PrezzoMedioInFattura = 1 Then
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                PrezzoDaReg = PrezzoMedio
                                ImportoScontiDaReg = ImportoScontiPM
                                'ImportoVarImballoDaReg = ImportoVarImballoPM
                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                
                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                                LiquidaAPrezzoMedio = True
'                            Else
'                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                ImportoScontiDaReg = ImportoScontiPM
'                                'ImportoVarImballoDaReg = ImportoVarImballoPM
'                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                                ImportoCommissioniDaReg = ImportoCommissioniPM
'                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
'                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
'
'                            End If
                        End If
                    Else
                        'If PrezzoMedioInFattura = 1 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            LiquidaAPrezzoMedio = True
'                        Else
'                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                            Link_Riga_Prezzo_Medio_TMP = 0
'                        End If
                    End If
                End If
            End If
            
            IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopArticoloMix)
            If (IDUMCoop = 0) Then
                IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopVendita)
            End If
            Select Case IDUMCoop
                Case 1
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Colli")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Colli) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 2
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoLordo")
                    If (TotaleQtaProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoLordo) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 3
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoNetto")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoNetto) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 4
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Tara")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Tara) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 5
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Pezzi")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Pezzi) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
            End Select
            
            
            If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                QuantitaLiqAbb = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe), fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 4)
                If Abs(QuantitaLiqAbb) > 0 Then
                    QuantitaLiq = Abs(QuantitaLiqAbb)
                End If
            End If
            
            QuantitaLiq = QuantitaLiq * -1
            If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
                QuantitaLiq = QuantitaLiq
            Else
                If LiquidaAPrezzoMedio = True Then
                    QuantitaLiq = 0
                Else
                    QuantitaLiq = QuantitaLiq
                End If
            End If
            If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
            If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

            
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
            sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, "
            ''''PEZZO NUOVO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiDaReg, ImportoVarImpImballoDaReg, ImportoCommissioniDaReg, ImpUniVendDocNettoIvaVenditaDaReg, "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            sSQL = sSQL & "IDRV_POTipoOggettoVariante, IDCategoriaMerceologica, "
            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali,  "
            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato,  "
            
            '''PEZZO NUOVO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, "
            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, "
            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, "
            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, "
            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, IDRV_POProcessoIVGammaRighe, "
            sSQL = sSQL & "IDArticoloMixVenduto, CodiceArticoloMixVenduto, DescrizioneArticoloMixVenduto"
            sSQL = sSQL & ") "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & 3 & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Articolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDTipoLavorazioneArticoloLavorato) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDRV_POTipoCategoriaArticoloLavorato) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDRV_POCalibroArticoloLavorato) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!DataLiquidazione) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDAnagraficaSocio) & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
            sSQL = sSQL & fnNormString("N.C. n° " & Trim(fnNotNull(rsFatt!Doc_Prefisso)) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) + (PrezzoMedio * fnNotNullN(QuantitaLiq))) & ", "
            
            
            TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)
            TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 2)
            TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
            
            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
            
                Case 1
                    If LINK_LISTINO = 0 Then
'                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                        
                        sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                    Else
                    
                        If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
'                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        End If
                    End If
                Case 3
                    LINK_TIPO_CALCOLO_NC = GET_TIPO_CALCOLO_PREZZO_MEDIO_NC
                    
                    Select Case LINK_TIPO_CALCOLO_NC
                        Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case Else
                                
                            End Select
                        Case 2 'INCLUSO VARIAZIONE PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                    
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 4  'Variazione di prezzo totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select
                                
                                                        
                        Case 3 'INCLUSO VARAZIONE PREZZO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                
                                Case Else
                                
                            End Select

                        
                        Case 4 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fQuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select

                        
                        Case Else 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1  'Variazione parziale di prezzo
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2  'Variazione parziale di peso
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3  'Variazione di peso totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4  'Variazione di prezzo totale
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case Else
                                
                            End Select

                    End Select
            End Select
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNotNullN(1) & ", "
                Case 2
                    sSQL = sSQL & fnNotNullN(2) & ", "
                Case 3
                    sSQL = sSQL & fnNotNullN(2) & ", "
                Case 4
                    sSQL = sSQL & fnNotNullN(1) & ", "
            End Select
            
            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologicaArticoloLavorato) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case 2
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 3
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 4
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case Else
                    
            End Select
            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ","
            'IMPORTO SCONTI
            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
            'IMPORTO VARIAZIONE IMBALLO
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
            'IMPORTO COMMISSIONI
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
            'IMPORTO UNITARIO NETTO IVA
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
            sSQL = sSQL & Link_Riga_Prezzo_Medio_TMP & ", "
            sSQL = sSQL & IDTipoPrezzoMedio & ", "
            sSQL = sSQL & fnNotNullN(0) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione)
            sSQL = sSQL & ")"

    CnDMT.Execute sSQL
End If
    
        
rsFatt.MoveNext
Wend

rsFatt.CloseResultset
Set rsFatt = Nothing
End Sub

Private Sub InserisciNotaDebito_MIX(ValoriOggettoDettaglio As Long, IDTipoOggetto As Long, PrezzoMedio As Double, PrezzoMedioInFattura As Long, CodiceLottoVendita As String, ImportoScontiPM As Double, ImportoVarImballoPM As Double, ImportoCommissioniPM As Double, ImportoNettoIvaPM As Double, Link_TMP_Prezzo_Medio As Long, IDTipoPrezzoMedio As Long, TotaleProcesso As Double, IDProcessoIVGammaRighe As Long)
Dim LINK_TIPO_CALCOLO_ND As Long
Dim TIPO_IMPORTO_ARTICOLO_LOCAL As Long
Dim PrezzoDaReg As Double
Dim ImportoScontiDaReg As Double
Dim ImportoVarImballoDaReg As Double
Dim ImportoCommissioniDaReg As Double
Dim ImportoNettoIvaDaReg As Double
Dim Link_Riga_Prezzo_Medio_TMP As Long
Dim QuantitaLiq As Double
Dim IDUMCoopVendita As Long
'Dim TotaleProcesso As Double
Dim QuantitaLiqAbb As Double

Dim TotaleImponibileDaReg As Double
Dim TotaleImpostaDaReg As Double
Dim TotaleLordoDaReg As Double
Dim LiquidaAPrezzoMedio As Boolean

'NOTA DI DEBITO
sSQL = "SELECT * FROM RV_POIELiquidazioneArticoliMixND "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND IDRV_POProcessoIVGammaRighe=" & IDProcessoIVGammaRighe
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND RV_POIDValoriOggettoDettaglio=" & ValoriOggettoDettaglio
End If

Set rsFatt = CnDMT.OpenResultset(sSQL)

While Not rsFatt.EOF
    If fnNotNullN(rsFatt!IDArticolo) > 0 Then
             LiquidaAPrezzoMedio = False
             If LINK_LISTINO > 0 Then
                PrezzoDaReg = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsFatt!IDArticolo), LINK_LISTINO)
                ImportoScontiDaReg = 0
                'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                ImportoVarImballoDaReg = 0
                
                ImportoCommissioniDaReg = 0
                ImportoNettoIvaDaReg = 0
                Link_Riga_Prezzo_Medio_TMP = 0
                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
            
            Else
                If fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq) > 0 Then
                    Select Case fnNotNullN(rsFatt!RV_POIDTipoImportoVenditaLiq)
                        Case 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                        Case 2
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                        Case Else
                            PrezzoDaReg = PrezzoMedio
                            
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                            
                            
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            LiquidaAPrezzoMedio = True
                    End Select
                Else
                    If TIPO_IMPORTO_ARTICOLO = 1 Then
                        If LINK_TIPO_PREZZO_MEDIO_ARTICOLO = 0 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
                            Link_Riga_Prezzo_Medio_TMP = 0
                        Else
                            'If PrezzoMedioInFattura = 1 Then
                                TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                                PrezzoDaReg = PrezzoMedio
                                ImportoScontiDaReg = ImportoScontiPM
                                'ImportoVarImballoDaReg = ImportoVarImballoPM
                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
                                
                                ImportoCommissioniDaReg = ImportoCommissioniPM
                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                                LiquidaAPrezzoMedio = True
'                            Else
'                                TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                                ImportoScontiDaReg = ImportoScontiPM
'                                'ImportoVarImballoDaReg = ImportoVarImballoPM
'                                'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                                ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                                ImportoCommissioniDaReg = ImportoCommissioniPM
'                                ImportoNettoIvaDaReg = ImportoNettoIvaPM
'                                Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
'
'                            End If
                        End If
                    Else
                        'If PrezzoMedioInFattura = 1 Then
                            TIPO_IMPORTO_ARTICOLO_LOCAL = 3
                            PrezzoDaReg = PrezzoMedio
                            ImportoScontiDaReg = ImportoScontiPM
                            'ImportoVarImballoDaReg = ImportoVarImballoPM
                            'ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
                            ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
    
                            ImportoCommissioniDaReg = ImportoCommissioniPM
                            ImportoNettoIvaDaReg = ImportoNettoIvaPM
                            Link_Riga_Prezzo_Medio_TMP = Link_TMP_Prezzo_Medio
                            LiquidaAPrezzoMedio = True
'                        Else
'                            TIPO_IMPORTO_ARTICOLO_LOCAL = 1
'                            PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
'                            ImportoScontiDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA)
'                            'ImportoVarImballoDaReg = fnNotNullN(rsFatt!RV_POImportoDaLiq)
'                            ImportoVarImballoDaReg = IIf((fnNotNullN(rsFatt!RV_POImportoDaLiq) < 0), Abs(fnNotNullN(rsFatt!RV_POImportoDaLiq)), fnNotNullN(rsFatt!RV_POImportoDaLiq))
'                            'ImportoVarImballoDaReg = IIf((ImportoVarImballoPM < 0), Abs(ImportoVarImballoPM), ImportoVarImballoPM)
'
'                            ImportoCommissioniDaReg = fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)
'                            ImportoNettoIvaDaReg = fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)
'                            Link_Riga_Prezzo_Medio_TMP = 0
'                        End If
                    End If
                End If
            End If
            
            IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopArticoloMix)
            If (IDUMCoop = 0) Then
                IDUMCoop = fnNotNullN(rsFatt!IDUnitaDiMisuraCoopVendita)
            End If
            Select Case IDUMCoop
                Case 1
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Colli")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Colli) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 2
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoLordo")
                    If (TotaleQtaProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoLordo) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 3
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "PesoNetto")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!PesoNetto) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 4
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Tara")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Tara) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
                Case 5
                    'TotaleProcesso = GetTotaleProcesso(rsFatt!IDRV_POProcessoIVGamma, "Pezzi")
                    If (TotaleProcesso > 0) Then QuantitaLiq = (fnNotNullN(rsFatt!Pezzi) / TotaleProcesso) * fnNotNullN(rsFatt!RV_POQuantitaLiq)
            End Select
            
            If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                QuantitaLiqAbb = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsFatt!IDRV_POCaricoMerceRighe), fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe), fnNotNullN(rsFatt!IDOggetto), fnNotNullN(rsFatt!IDTipoOggetto), 4)
                If QuantitaLiqAbb > 0 Then
                    QuantitaLiq = QuantitaLiqAbb
                End If
            End If
            
            If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
                QuantitaLiq = QuantitaLiq
            Else
                If LiquidaAPrezzoMedio = True Then
                    QuantitaLiq = 0
                Else
                    QuantitaLiq = QuantitaLiq
                End If
            End If
            
            
            If DEC_QTA_LIQ > 0 Then QuantitaLiq = FormatNumber(QuantitaLiq, DEC_QTA_LIQ)
            If DEC_IMP_UNI_LIQ > 0 Then PrezzoDaReg = FormatNumber(PrezzoDaReg, DEC_IMP_UNI_LIQ)

            
            sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
            sSQL = sSQL & "TipoRiga, IDArticolo, CodiceArticolo, Articolo, IDTipoLavorazione, IDTipoCategoria, IDCalibro, "
            sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, "
            sSQL = sSQL & "IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
            sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
            sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
            sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
            sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
            sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
            sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, "
            ''''PEZZO NUOVO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiDaReg, ImportoVarImpImballoDaReg, ImportoCommissioniDaReg, ImpUniVendDocNettoIvaVenditaDaReg, "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            sSQL = sSQL & "IDRV_POTipoOggettoVariante, IDCategoriaMerceologica, "
            sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, Quantita_Per_totali,  "
            sSQL = sSQL & "ImpUniVendDocLordo, ImpUniVendDocLordoScontato,  "
            
            '''PEZZO NUOVO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "ImportoScontiVendita, ImportoScontiPM, "
            sSQL = sSQL & "ImportoVarImpImballoVendita, ImportoVarImpImballoPM, "
            sSQL = sSQL & "ImportoCommissioniVendita, ImportoCommissioniPM, "
            sSQL = sSQL & "ImpUniVendDocNettoIvaVendita, ImpUniVendDocNettoIvaVenditaPM, "
            sSQL = sSQL & "IDRV_POTMPLiquidazionePrezzoMedio, IDRV_POTipoPrezzoMedio, Invenduto, IDRV_POProcessoIVGammaRighe, "
            sSQL = sSQL & "IDArticoloMixVenduto, CodiceArticoloMixVenduto, DescrizioneArticoloMixVenduto"
            sSQL = sSQL & ") "
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & 3 & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Articolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoLavorazione) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDTipoCategoria) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POIDCalibro) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!art_numero_colli) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_peso) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_peso) - fnNotNullN(rsFatt!Art_tara)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_tara) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_pezzi) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
            sSQL = sSQL & LINK_PERIODO & ", "
            sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
            sSQL = sSQL & fnNormString("N.D. n° " & fnNotNull(rsFatt!Doc_Prefisso) & "/" & fnNotNull(rsFatt!doc_numero) & " del " & fnNotNull(rsFatt!doc_data)) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
            sSQL = sSQL & fnNormDate(rsFatt!doc_data) & ", "
            sSQL = sSQL & fnNormString(rsFatt!doc_numero) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
            sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!IDIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!CodiceIvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!AliquotaIvaArticolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!IvaArticolo) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo))) & ","
            sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * fnNotNullN(QuantitaLiq)) + (PrezzoMedio * fnNotNullN(QuantitaLiq))) & ", "
            
            TotaleImponibileDaReg = FormatNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq, 2)
            TotaleImpostaDaReg = FormatNumber((TotaleImponibileDaReg / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo), 2)
            TotaleLordoDaReg = TotaleImponibileDaReg + TotaleImpostaDaReg
            
            
            Select Case TIPO_IMPORTO_ARTICOLO_LOCAL
            
                Case 1
                    If LINK_LISTINO = 0 Then
'                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq)) & ", "
'                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) & ", "
'                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * (QuantitaLiq)) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(QuantitaLiq))) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                        sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                        sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                    Else
                        If ((fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rsFatt!RV_POIDTipoVariazione) = 3)) Then
'                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
'                            sSQL = sSQL & fnNormNumber(fnNotNullN(PrezzoDaReg) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                            sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoDaReg) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(PrezzoDaReg) * QuantitaLiq)) & ", "
                            sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        Else
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                            sSQL = sSQL & fnNormNumber(0) & ", "
                        End If
                    End If
                Case 3
                    LINK_TIPO_CALCOLO_ND = GET_TIPO_CALCOLO_PREZZO_MEDIO_ND
                    
                    
                    Select Case LINK_TIPO_CALCOLO_ND
                        Case 1 'INCLUSO VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                
                                Case 1
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                    
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                Case 4
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "

                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                            
                            End Select
                        Case 2 'INCLUSO VARIAZIONE PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
        
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
'                                    sSQL = sSQL & fnNormNumber(PrezzoMedio * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (PrezzoMedio * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
        
                                    sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
                                    sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                    
                                    
                            End Select
                                                        
                        Case 3 'INCLUSO VARAZIONE PREZZO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "

                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0 * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(0) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (0 * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                                    sSQL = sSQL & fnNormNumber(0) & ", "
                            End Select
                        
                        Case 4 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            
                            End Select
                        
                        Case Else 'ESCLUDI VARIAZIONE PREZZO E PESO
                            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                                Case 1
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 2
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 3
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                                
                                Case 4
'                                    sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) & ", "
'                                    sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * fnNotNullN(rsFatt!AliquotaIvaArticolo)) * QuantitaLiq) + (fnNotNullN(rsFatt!RV_POImportoLiq) * QuantitaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(PrezzoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImponibileDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleImpostaDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber(TotaleLordoDaReg) & ", "
                                    sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
                                    sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
                            
                            End Select
                    End Select
            End Select
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNotNullN(3) & ", "
                Case 2
                    sSQL = sSQL & fnNotNullN(4) & ", "
                Case 3
                    sSQL = sSQL & fnNotNullN(4) & ", "
                Case 4
                    sSQL = sSQL & fnNotNullN(3) & ", "
            End Select
            
            sSQL = sSQL & fnNotNullN(rsFatt!IDCategoriaMerceologica) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            sSQL = sSQL & fnNormNumber(0) & ", "
            
            Select Case fnNotNullN(rsFatt!RV_POIDTipoVariazione)
                Case 1
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case 2
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 3
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
                Case 4
                    sSQL = sSQL & fnNormNumber(0) & ", "
                Case Else
                    sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
            End Select
            sSQL = sSQL & fnNormNumber(rsFatt!Art_prezzo_unitario_lordo_IVA) & ", "
            sSQL = sSQL & fnNormNumber(rsFatt!Art_importo_net_sconto_lor_IVA) & ","
            'IMPORTO SCONTI
            sSQL = sSQL & fnNormNumber((fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA) - fnNotNullN(rsFatt!Art_pre_uni_net_sco_net_IVA))) & ", "
            sSQL = sSQL & fnNormNumber(ImportoScontiPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoScontiDaReg) & ", "
            'IMPORTO VARIAZIONE IMBALLO
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoDaLiq)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoVarImballoPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoVarImballoDaReg) & ", "
            'IMPORTO COMMISSIONI
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoRigaCommissioni)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoCommissioniPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoCommissioniDaReg) & ", "
            'IMPORTO UNITARIO NETTO IVA
            sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!Art_prezzo_unitario_netto_IVA)) & ", "
            sSQL = sSQL & fnNormNumber(ImportoNettoIvaPM) & ", "
            'sSQL = sSQL & fnNormNumber(ImportoNettoIvaDaReg) & ", "
            'LINK DEL TIPO PREZZO MEDIO TEMPORANEO CALCOLATO
            sSQL = sSQL & Link_Riga_Prezzo_Medio_TMP & ", "
            sSQL = sSQL & IDTipoPrezzoMedio & ", "
            sSQL = sSQL & fnNotNullN(0) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!IDRV_POProcessoIVGammaRighe) & ", "
            sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_codice) & ", "
            sSQL = sSQL & fnNormString(rsFatt!Art_Descrizione)
            sSQL = sSQL & ")"
            
        CnDMT.Execute sSQL
    End If
    
        
rsFatt.MoveNext
Wend

rsFatt.CloseResultset
Set rsFatt = Nothing
End Sub

Private Sub EliminaQtaAbbConf(IDConferimentoRiga As Long)
On Error GoTo ERR_EliminaQtaAbbConf
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaLiquidazioneRiga As Long

sSQL = "SELECT * FROM RV_POLiquidazioneConfQtaAbb "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!TipoRigaLiquidazione)) = True) Then
            AvviaLiquidazioneRiga = 0
        End If
    End If
    If AvviaLiquidazioneRiga = 1 Then
        sSQL = "DELETE FROM RV_POLiquidazioneConfQtaAbb "
        sSQL = sSQL & "WHERE ID=" & fnNotNullN(rs!ID)
        CnDMT.Execute sSQL
    End If
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
Exit Sub
ERR_EliminaQtaAbbConf:
    MsgBox Err.Description, vbCritical, "EliminaQtaAbbConf"
End Sub
Private Function GET_CONTROLLO_CONF_VEND(IDRigaConferimento As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_CONF_VEND
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_CONF_VEND = True

If TIPO_LIQUIDAZIONE = 2 Then
    sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POTMPLiquidazioneVendita "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLO_CONF_VEND = False
    End If
    rs.CloseResultset
    Set rs = Nothing
End If

Exit Function
ERR_GET_CONTROLLO_CONF_VEND:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_CONF_VEND"
End Function
Private Sub GET_SCARTI_VENDITA(prg As ProgressBar, DataInizio As String, DataFine As String, TipoLiquidazione As Long)
Dim sSQL As String
Dim rsScarti As DmtOleDbLib.adoResultset
Dim Link_CategoriaMerceologica As Long
Dim Link_UM_Coop As Long
Dim QuantitaLiq As Double
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaLiqAbbattuta As Double
Dim IDSocio As Long
Dim IDConferimentoRiga As Long

sSQL = "SELECT RV_POLavorazione.IDRV_POLavorazione, RV_POLavorazione.IDRV_POCaricoMerceRighe, RV_POLavorazione.IDTipoLavorazione, "
sSQL = sSQL & "RV_POLavorazione.IDArticolo, RV_POLavorazione.CodiceArticolo, RV_POLavorazione.Articolo, RV_POLavorazione.Colli, "
sSQL = sSQL & "RV_POLavorazione.PesoLordo, RV_POLavorazione.PesoNetto, RV_POLavorazione.Tara, RV_POLavorazione.Pezzi, RV_POLavorazione.Qta_UM, "
sSQL = sSQL & "RV_POLavorazione.IDImballoVendita, RV_POLavorazione.CodiceImballoVendita, RV_POLavorazione.ImballoVendita, "
sSQL = sSQL & "RV_POLavorazione.IDRV_POCalibro, RV_POLavorazione.IDRV_POTipoCategoria, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceTesta.DataDocumento , RV_POCaricoMerceTesta.IDAnagrafica, RV_POLavorazione.DataDocumento AS DataLavorazioneScarto, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDAnagrafica, RV_POCaricoMerceTesta.PreConferimento "
sSQL = sSQL & "FROM RV_POLavorazione INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POLavorazione.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & " WHERE RV_POCaricoMerceTesta.IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND RV_POCaricoMerceTesta.PreConferimento=0"

If TIPO_LIQUIDAZIONE <> 1 Then
    sSQL = sSQL & " AND RV_POLavorazione.DataDocumento>=" & fnNormDate(DATA_INIZIO)
    sSQL = sSQL & " AND RV_POLavorazione.DataDocumento<=" & fnNormDate(DATA_FINE)
End If
If LIQUIDA_FORNITORE = 1 Then
    sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDAnagrafica=" & LINK_SOCIO
End If


Set rsScarti = CnDMT.OpenResultset(sSQL)
While Not rsScarti.EOF
    If ARTICOLI_DI_QUAD = False Then
        IDConferimentoRiga = fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe)
        IDSocio = fnNotNullN(rsScarti!IDAnagrafica)
        TrattenutePerLavorazione = 0
        TrattenuteGenerali = 0
        TrattenuteTotali = 0
        TrattValGen1 = 0
        TrattValGen2 = 0
        TrattPercGen1 = 0
        TrattPercGen2 = 0
        
        TrattValLav1 = 0
        TrattValLav2 = 0
        TrattPercLav1 = 0
        TrattPercLav2 = 0
            
        If fnNotNullN(rsScarti!IDArticolo) > 0 Then
            AvviaLiquidazioneRiga = 1
            If NO_LIQ_VEND_UFF = 1 Then
                If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rsScarti!IDRV_POLavorazione), 0, 0, 2) = True) Then
                AvviaLiquidazioneRiga = 0
                End If
            End If
'            If NO_RIP_SCARTI_IN_LIQ = 1 Then
'                AvviaLiquidazioneRiga = 0
'            End If
            If AvviaLiquidazioneRiga = 1 Then
                Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsScarti!IDArticolo))
                QuantitaLiq = fnNotNullN(rsScarti!Qta_UM)
                Link_UM_Coop = GET_LINK_UM_COOP(fnNotNullN(rsScarti!IDArticolo))
                Select Case Link_UM_Coop
                    Case 1
                        QuantitaLiq = fnNotNullN(rsScarti!Colli)
                    Case 2
                        QuantitaLiq = fnNotNullN(rsScarti!PesoLordo)
                    Case 3
                        QuantitaLiq = fnNotNullN(rsScarti!PesoNetto)
                    Case 4
                        QuantitaLiq = fnNotNullN(rsScarti!Tara)
                    Case 5
                        QuantitaLiq = fnNotNullN(rsScarti!Pezzi)
                End Select
                If LINK_TIPO_AUMENTO_PESO = GET_TIPO_SCARTO_ARTICOLO(fnNotNullN(rsScarti!IDArticolo)) Then
                    QuantitaLiq = QuantitaLiq * -1
                End If
                
                
                If (ATTIVA_CALCOLO_QTA_DA_ABB = 1) Then
                    QuantitaLiqAbbattuta = GET_QUANTITA_ABBATTUTA(fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), fnNotNullN(rsScarti!IDRV_POLavorazione), 0, 0, 2)
                    If QuantitaLiqAbbattuta = 0 Then
                        QuantitaLiq = QuantitaLiq
                    Else
                        QuantitaLiq = QuantitaLiqAbbattuta
                    End If
                End If
                
                'TrattenuteArticolo IDSocio, fnNotNullN(rsScarti!IDArticolo), Link_CategoriaMerceologica, fnNotNullN(rsScarti!IDTipoLavorazione), False, fnNotNullN(rsScarti!Qta_UM), 0, fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), 0, 0, fnNotNullN(rsScarti!IDRV_POLavorazione), 0
                TrattenuteArticolo IDSocio, fnNotNullN(rsScarti!IDArticolo), Link_CategoriaMerceologica, fnNotNullN(rsScarti!IDTipoLavorazione), False, QuantitaLiq, 0, fnNotNullN(rsScarti!IDRV_POCaricoMerceRighe), 0, 0, fnNotNullN(rsScarti!IDRV_POLavorazione), 0
                            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "TipoRiga, IDRV_POCaricoMerceRighe, IDArticolo, CodiceArticolo, Articolo, Quantita, "
                sSQL = sSQL & "DataConferimento, IDSocio, IDValoreOggettoDettaglio, IDRV_POPeriodoLiquidazione, "
                sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali, "
                sSQL = sSQL & "Colli, PesoLordo, PesoNetto, Tara, Pezzi, IDTipoLavorazione, IDCalibro, IDTipoCategoria, "
                sSQL = sSQL & "IDImballo, CodiceImballo, Imballo, "
                sSQL = sSQL & "TrattenutaValoreGen1, TrattenutaValoreGen2, TrattenutaPercGen1, TrattenutaPercGen2, "
                sSQL = sSQL & "TrattenutaValoreLav1, TrattenutaValoreLav2, TrattenutaPercLav1, TrattenutaPercLav2) "
    
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnNormNumber(2) & ", "
                sSQL = sSQL & IDConferimentoRiga & ", "
                sSQL = sSQL & fnNotNullN(rsScarti!IDArticolo) & ", "
                sSQL = sSQL & fnNormString(rsScarti!CodiceArticolo) & ", "
                sSQL = sSQL & fnNormString(rsScarti!Articolo) & ", "
                'sSQL = sSQL & fnNormNumber(rsScarti!Qta_UM) & ", "
                sSQL = sSQL & fnNormNumber(QuantitaLiq) & ", "
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
                sSQL = sSQL & fnNormString(rsScarti!ImballoVendita) & ", "
                sSQL = sSQL & fnNormNumber(TrattValGen1) & ", "
                sSQL = sSQL & fnNormNumber(TrattValGen2) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercGen1) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercGen2) & ", "
                sSQL = sSQL & fnNormNumber(TrattValLav1) & ", "
                sSQL = sSQL & fnNormNumber(TrattValLav2) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercLav1) & ", "
                sSQL = sSQL & fnNormNumber(TrattPercLav2) & ")"
                
                CnDMT.Execute sSQL
            End If
        End If
    End If

rsScarti.MoveNext
Wend


rsScarti.CloseResultset
Set rsScarti = Nothing

End Sub
Private Sub CALCOLO_CONF_QTA_ABB_VEND(IDTipoOggetto As Long, NomeTabella As String, TipoRaggruppamento As Long, DataInizio As String, DataFine As String)
On Error GoTo ERR_CALCOLO_CONF_QTA_ABB_VEND
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

Dim PercAbbConferimento As Double
Dim PercAbbConferimentoScarto As Double

Dim TotaleVenduto As Double
Dim TotaleQuadratura As Double
Dim TotaleConferimentoElaborato As Double
Dim TotaleQuadraturaAbbattuta As Double
Dim TotaleConferimentoVenduto As Double
Dim DescrizioneConferimento As String
Dim TotaleProcesso As Double
Dim QuantitaSingoloProcesso As Double
Dim IDUMCoopProcesso As Long
Dim AvviaLiquidazioneRiga As Long
Dim IDRiferimentoRaggruppamento As Long
Dim IDAnagraficaSocio As Long
Dim AvviaCalcolo As Boolean

FrmNuovoPeriodo.List1.Clear

TotaleQuadratura = 0
TotaleVenduto = 0
TotaleConferimentoElaborato = 0
TotaleQuadraturaAbbattuta = 0
TotaleConferimentoVenduto = 0
PercAbbConferimento = 0
PercAbbConferimentoScarto = 0
DescrizioneConferimento = ""

sSQL = "SELECT * FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1 "
sSQL = sSQL & " AND DataCompetenzaLiq>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND DataCompetenzaLiq<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazione>0"
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1))) "
If ((IDTipoOggetto = 11) Or (IDTipoOggetto = 107)) Then
    sSQL = sSQL & " AND ((RV_POIDTipoVariazione = 2) OR (RV_POIDTipoVariazione=3))"
End If

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    
'    If fnNotNull(rs!Numero) = "2323" Then
'        MsgBox "STOP"
'    End If
    
    FrmNuovoPeriodo.List1.AddItem "Calcolo quantità da abbattere: " & fnNotNull(rs!Oggetto) & " n° " & fnNotNull(rs!Numero) & " del " & fnNotNull(rs!DataCompetenzaLiq)
    FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
    FrmNuovoPeriodo.lblInfoStatus.Caption = "Calcolo quantità da abbattere: " & fnNotNull(rs!Oggetto) & " n° " & fnNotNull(rs!Numero) & " del " & fnNotNull(rs!DataCompetenzaLiq)
    
    DoEvents

    TotaleQuadratura = 0
    TotaleVenduto = 0
    TotaleConferimentoElaborato = 0
    TotaleQuadraturaAbbattuta = 0
    TotaleConferimentoVenduto = 0
    PercAbbConferimento = 0
    PercAbbConferimentoScarto = 0
    DescrizioneConferimento = ""


    TotaleVenduto = TotaleVenduto + GetTotaleVenduto(2, "RV_POIELiquidazioneControlloConfQtaAbbDDT", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVenduto(114, "RV_POIELiquidazioneControlloConfQtaAbbFA", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVenduto(8, "RV_POIELiquidazioneControlloConfQtaAbbSNF", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVenduto(11, "RV_POIELiquidazioneControlloConfQtaAbbNC", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVenduto(107, "RV_POIELiquidazioneControlloConfQtaAbbND", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVendutoMix(2, "RV_POIELiquidazioneArticoliMixDDT", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVendutoMix(114, "RV_POIELiquidazioneArticoliMixFA", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVendutoMix(8, "RV_POIELiquidazioneArticoliMixSNF", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVendutoMix(11, "RV_POIELiquidazioneArticoliMixNC", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    TotaleVenduto = TotaleVenduto + GetTotaleVendutoMix(107, "RV_POIELiquidazioneArticoliMixND", fnNotNullN(rs!IDAnagraficaSocio), fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    
    TotaleQuadratura = GetTotaleScartoConferimentoVend(fnNotNullN(rs!IDCategoriaLiquidazione), fnNotNullN(rs!IDAnagraficaSocio), TotaleQuadraturaAbbattuta, fnNotNullN(rs!DataCompetenzaLiq), TipoRaggruppamento)
    
    TotaleConferimentoElaborato = (TotaleVenduto + TotaleQuadratura) - (((TotaleVenduto + TotaleQuadratura) / 100) * fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazione))
    
    If fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto) > 0 Then
        TotaleQuadraturaAbbattuta = (TotaleQuadratura) - (((TotaleQuadratura) / 100) * fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto))
    Else
        TotaleQuadraturaAbbattuta = TotaleQuadraturaAbbattuta
    End If
    
    
    TotaleConferimentoVenduto = TotaleConferimentoElaborato - TotaleQuadraturaAbbattuta
    
    If TotaleVenduto = 0 Then Exit Sub
    
    
    sSQL = "DELETE FROM RV_POLiquidazioneConfQtaAbb "
    sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rs!IDValoriOggettoDettaglio)
    sSQL = sSQL & " AND IDOggetto=" & fnNotNullN(rs!IDOggetto)
    sSQL = sSQL & " AND IDTipoOggetto=" & fnNotNullN(rs!IDTipoOggetto)
    CnDMT.Execute sSQL
    
    sSQL = "SELECT * FROM RV_POLiquidazioneConfQtaAbb "
    sSQL = sSQL & "WHERE ID=0"
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

    AvviaLiquidazioneRiga = 1
'    If NO_LIQ_VEND_UFF = 1 Then
'        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
'            AvviaLiquidazioneRiga = 0
'        End If
'    End If

    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = fnNotNullN(rs!RV_POIDConferimentoRighe)
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
            rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
            rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
            rsNew!TipoRigaLiquidazione = 1
            rsNew!Quantita = (fnNotNullN(rs!RV_POQuantitaLiq) / TotaleVenduto) * TotaleConferimentoVenduto
            rsNew!QuantitaOriginale = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
    
    rsNew.Close
    Set rsNew = Nothing
        
    rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CALCOLO_CONF_QTA_ABB_VEND:
    MsgBox Err.Description, vbCritical, "CALCOLO_CONF_QTA_ABB_VEND"
End Sub
Private Function GetTotaleVenduto(IDTipoOggetto As Long, NomeTabella As String, IDAnagraficaSocio As Long, IDRiferimentoGruppo As Long, DataDocumento As String, Optional TipoRaggruppamento As Long = 0) As Double
On Error GoTo ERR_GetTotaleVendutoConferimento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaLiquidazioneRiga As Long
Dim IDSocio As Long

GetTotaleVenduto = 0
IDSocio = 0

sSQL = "SELECT * "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND DataCompetenzaLiq=" & fnNormDate(DataDocumento)
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"
Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND RV_POIDConferimentoRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If

If ((IDTipoOggetto = 11) Or (IDTipoOggetto = 107)) Then
    sSQL = sSQL & " AND ((RV_POIDTipoVariazione=2) OR (RV_POIDTipoVariazione=3))"
End If

Set rs = CnDMT.OpenResultset(sSQL)


While Not rs.EOF
        AvviaLiquidazioneRiga = 1
'        If NO_LIQ_VEND_UFF = 1 Then
'            If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)) = True) Then
'                AvviaLiquidazioneRiga = 0
'            End If
'        End If
        If AvviaLiquidazioneRiga = 1 Then
            GetTotaleVenduto = GetTotaleVenduto + fnNotNullN(rs!RV_POQuantitaLiq)
        End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If GetTotaleVenduto > 0 Then
    If IDTipoOggetto = 11 Then GetTotaleVenduto = GetTotaleVenduto * -1
End If
Exit Function
ERR_GetTotaleVendutoConferimento:
    MsgBox Err.Description, vbCritical, "GetTotaleVenduto"
End Function
Private Function GetTotaleScartoConferimentoVend(IDRiferimentoGruppo As Long, IDAnagraficaSocio As Long, TotaleQuadraturaAbbattuta As Double, DataDocumento As String, Optional TipoRaggruppamento As Long = 0) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaTotale As Double
Dim AvviaLiquidazioneRiga As Long
Dim QuantitaSingolo As Double
Dim IDSocio As Long

GetTotaleScartoConferimentoVend = 0
QuantitaTotale = 0
TotaleQuadraturaAbbattuta = 0
IDSocio = 0

sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbScarto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND PreConferimento=0"
sSQL = sSQL & " AND DataDocumento=" & fnNormDate(DataDocumento)
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"

Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If

'If TIPO_LIQUIDAZIONE <> 1 Then
'    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DATA_INIZIO)
'    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
'End If

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
    If NO_LIQ_VEND_UFF = 1 Then
        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDRV_POLavorazione), 0, 0, 2) = True) Then
        AvviaLiquidazioneRiga = 0
        End If
    End If

    If AvviaLiquidazioneRiga = 1 Then
        Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
            Case 1
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Colli)
                QuantitaSingolo = fnNotNullN(rs!Colli)
            Case 2
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!PesoLordo)
                QuantitaSingolo = fnNotNullN(rs!PesoLordo)
            Case 3
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!PesoNetto)
                QuantitaSingolo = fnNotNullN(rs!PesoNetto)
            Case 4
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Tara)
                QuantitaSingolo = fnNotNullN(rs!Tara)
            Case 5
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Pezzi)
                QuantitaSingolo = fnNotNullN(rs!Pezzi)
            Case Else
                'QuantitaTotale = QuantitaTotale + fnNotNullN(rs!Qta_UM)
                QuantitaSingolo = fnNotNullN(rs!Qta_UM)
        End Select
        
        If (LINK_TIPO_AUMENTO_PESO = fnNotNullN(rs!IDTipoProdotto)) Then
            QuantitaSingolo = QuantitaSingolo * -1
        End If
        
        QuantitaTotale = QuantitaTotale + QuantitaSingolo
        
'        If (fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto) > 0) Then
'            QuantitaSingolo = QuantitaSingolo - ((QuantitaSingolo / 100) * fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto))
'        Else
            If fnNotNullN(rs!PercentualeAbbattimentoScarto) > 0 Then
                QuantitaSingolo = QuantitaSingolo - ((QuantitaSingolo / 100) * fnNotNullN(rs!PercentualeAbbattimentoScarto))
            End If
'        End If
        TotaleQuadraturaAbbattuta = TotaleQuadraturaAbbattuta + QuantitaSingolo
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

GetTotaleScartoConferimentoVend = QuantitaTotale

End Function
Private Function GetTotaleVendutoMix(IDTipoOggetto As Long, NomeTabella As String, IDAnagraficaSocio As Long, IDRiferimentoGruppo As Long, DataDocumento As String, Optional TipoRaggruppamento As Long = 0) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TotaleQtaProcesso As Double
Dim TotaleQtaRigaVendita As Double
Dim QtaSingolaRiga As Double
Dim IDUMCoop As Long
Dim AvviaLiquidazioneRiga As Long
Dim IDSocio As Long

GetTotaleVendutoMix = 0
IDSocio = 0

sSQL = "SELECT * "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND RV_POIDTipoUtilizzoLinea=4"
sSQL = sSQL & " AND DataCompetenzaLiq=" & fnNormDate(DataDocumento)
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1)))"

Select Case TipoRaggruppamento
    Case 1
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 2
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 3
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 4
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 5
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 6
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 7
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 8
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    Case 9
        sSQL = sSQL & " AND IDCategoriaFiscale=" & IDRiferimentoGruppo
        
    Case 10
        sSQL = sSQL & " AND IDClassificazioneArticolo=" & IDRiferimentoGruppo
        
    Case 11
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & IDRiferimentoGruppo
        
    Case 12
        sSQL = sSQL & " AND IDRepartoFiscale=" & IDRiferimentoGruppo
        
    Case 13
        sSQL = sSQL & " AND IDGruppoArticoloPerEvasioneMix=" & IDRiferimentoGruppo
        
    Case 14
        sSQL = sSQL & " AND IDFamigliaProdotti =" & IDRiferimentoGruppo
        
    Case 15
        sSQL = sSQL & " AND IDVarieta=" & IDRiferimentoGruppo
        
    Case 16
        sSQL = sSQL & " AND IDCategoriaLiquidazione=" & IDRiferimentoGruppo
        
    Case Else
        sSQL = sSQL & " AND RV_POIDConferimentoRighe=" & IDRiferimentoGruppo
        IDSocio = IDAnagraficaSocio
    
End Select
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaSocio=" & IDSocio
End If
If ((IDTipoOggetto = 11) Or (IDTipoOggetto = 107)) Then
    sSQL = sSQL & " AND ((RV_POIDTipoVariazione=2) OR (RV_POIDTipoVariazione=3))"
End If

Set rs = CnDMT.OpenResultset(sSQL)

IDUMCoop = 3

While Not rs.EOF
    AvviaLiquidazioneRiga = 1
'    If NO_LIQ_VEND_UFF = 1 Then
'        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDRV_POProcessoIVGammaRighe)) = True) Then
'            AvviaLiquidazioneRiga = 0
'        End If
'    End If
    
    If AvviaLiquidazioneRiga = 1 Then
        TotaleQtaProcesso = 0
        QtaSingolaRiga = 0
        TotaleQtaRigaVendita = fnNotNullN(rs!RV_POQuantitaLiq)
        
        IDUMCoop = fnNotNullN(rs!IDUnitaDiMisuraCoopArticoloMix)
        If (IDUMCoop = 0) Then
            IDUMCoop = fnNotNullN(rs!IDUnitaDiMisuraCoopVendita)
        End If
        
        
        Select Case IDUMCoop
            Case 1
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Colli")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Colli) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 2
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoLordo")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!PesoLordo) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 3
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "PesoNetto")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!PesoNetto) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 4
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Tara")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Tara) / TotaleQtaProcesso) * TotaleQtaRigaVendita
            Case 5
                TotaleQtaProcesso = GetTotaleProcesso(fnNotNullN(rs!IDRV_POProcessoIVGamma), "Pezzi")
                If (TotaleQtaProcesso > 0) Then QtaSingolaRiga = (fnNotNullN(rs!Pezzi) / TotaleQtaProcesso) * TotaleQtaRigaVendita
        End Select
        
        GetTotaleVendutoMix = GetTotaleVendutoMix + QtaSingolaRiga
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If GetTotaleVendutoMix > 0 Then
    If IDTipoOggetto = 11 Then GetTotaleVendutoMix = GetTotaleVendutoMix * -1
End If
End Function

Private Sub CALCOLO_QTA_ABB_VEND_SCARTO(TipoRaggruppamento As Long, DataInizio As String, DataFine As String)
On Error GoTo ERR_CALCOLO_QTA_ABB_VEND_SCARTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

Dim PercAbbConferimento As Double
Dim PercAbbConferimentoScarto As Double

Dim TotaleVenduto As Double
Dim TotaleQuadratura As Double
Dim TotaleConferimentoElaborato As Double
Dim TotaleQuadraturaAbbattuta As Double
Dim TotaleConferimentoVenduto As Double
Dim DescrizioneConferimento As String
Dim TotaleProcesso As Double
Dim QuantitaSingoloProcesso As Double
Dim IDUMCoopProcesso As Long
Dim AvviaLiquidazioneRiga As Long
Dim IDRiferimentoRaggruppamento As Long
Dim IDAnagraficaSocio As Long
Dim AvviaCalcolo As Boolean

FrmNuovoPeriodo.List1.Clear

TotaleQuadratura = 0
TotaleVenduto = 0
TotaleConferimentoElaborato = 0
TotaleQuadraturaAbbattuta = 0
TotaleConferimentoVenduto = 0
PercAbbConferimento = 0
PercAbbConferimentoScarto = 0
DescrizioneConferimento = ""

sSQL = "SELECT * FROM RV_POIELiquidazioneControlloConfQtaAbbScarto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazione>0"
sSQL = sSQL & " AND RV_POPercentualeAbbattimentoLiquidazioneScarto>0"
sSQL = sSQL & " AND PreConferimento=0"
sSQL = sSQL & " AND ((IDTipoDocumentoCoop=1) OR ((IDTipoDocumentoCoop=2) AND (TrattaComeConferimento=1))) "

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    
    FrmNuovoPeriodo.List1.AddItem "Calcolo quantità da abbattere: Scarto " & fnNotNull(rs!DataDocumento)
    FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
    FrmNuovoPeriodo.lblInfoStatus.Caption = "Calcolo quantità da abbattere: Scarto " & fnNotNull(rs!DataDocumento)
    
    DoEvents
    
    
    sSQL = "DELETE FROM RV_POLiquidazioneConfQtaAbb "
    sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rs!IDRV_POLavorazione)
    sSQL = sSQL & " AND IDOggetto=0 "
    sSQL = sSQL & " AND IDTipoOggetto=0"
    CnDMT.Execute sSQL
    
    
    sSQL = "SELECT * FROM RV_POLiquidazioneConfQtaAbb "
    sSQL = sSQL & "WHERE ID=0"
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

    
    
    AvviaLiquidazioneRiga = 1
'    If NO_LIQ_VEND_UFF = 1 Then
'        If (GET_CONTROLLO_VENDITA_LIQ(fnNotNullN(rs!IDRV_POLavorazione), 0, 0, 2) = True) Then
'        AvviaLiquidazioneRiga = 0
'        End If
'    End If
    If (AvviaLiquidazioneRiga = 1) Then
        rsNew.AddNew
            rsNew!IDRV_POCaricoMerceRighe = fnNotNullN(rs!IDRV_POCaricoMerceRighe)
            rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDRV_POLavorazione)
            rsNew!IDOggetto = 0
            rsNew!IDTipoOggetto = 0
            rsNew!TipoRigaLiquidazione = 2
            
            Select Case fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
                Case 1
                    rsNew!Quantita = fnNotNullN(rs!Colli)
                Case 2
                    rsNew!Quantita = fnNotNullN(rs!PesoLordo)
                Case 3
                    rsNew!Quantita = fnNotNullN(rs!PesoNetto)
                Case 4
                    rsNew!Quantita = fnNotNullN(rs!Tara)
                Case 5
                    rsNew!Quantita = fnNotNullN(rs!Pezzi)
                Case Else
                    rsNew!Quantita = fnNotNullN(rs!Qta_UM)
            End Select
            rsNew!QuantitaOriginale = rsNew!Quantita
            If (fnNotNullN(rs!IDTipoProdotto) = LINK_TIPO_AUMENTO_PESO) Then
                rsNew!QuantitaOriginale = rsNew!Quantita * -1
            End If

            If fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto) > 0 Then
                rsNew!Quantita = rsNew!Quantita - ((rsNew!Quantita / 100) * fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto))
            End If
            If (fnNotNullN(rs!IDTipoProdotto) = LINK_TIPO_AUMENTO_PESO) Then
                rsNew!Quantita = rsNew!Quantita * -1
            End If
        rsNew.Update
    End If
    
    rsNew.Close
    Set rsNew = Nothing
    
    
    rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CALCOLO_QTA_ABB_VEND_SCARTO:
    MsgBox Err.Description, vbCritical, "CALCOLO_QTA_ABB_VEND_SCARTO"
End Sub
Private Function GET_TIPO_SCARTO_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TIPO_SCARTO_ARTICOLO = 0

sSQL = "SELECT IDArticolo, IDTipoProdotto "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_TIPO_SCARTO_ARTICOLO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
