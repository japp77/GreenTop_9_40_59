Attribute VB_Name = "ModCalcoloLiqVenduto"


Dim rsLav As ADODB.Recordset
Dim rsVend As ADODB.Recordset
Private ArrayQuad(3, 3) As String
Private ArrayTratt(9, 4) As String
Private Link_AddebitoImballo As Integer
Private Link_ListinoImballo As Long
Private Link_Iva As Long
Private Codice_Iva As String
Private Aliquota_Iva As Double
Private Iva As String
Private Link_Iva_Medio As Long
Private Codice_Iva_Medio As String
Private Aliquota_Iva_Medio As Double
Private Iva_Medio As String

Private Data_Documento_Conferimento As String

'Variabile per le trattenute
Private TrattenutePerLavorazione As Double
Private TrattenuteGenerali As Double
Private TrattenuteTotali As Double

Private TotaleRighe As Long


Public Sub VenditeArticolo(IDRigaVendita As Long, IDLiquidazione As Long, IDPeriodoLiquidazione As Long, IDTipoOggetto As Long, IDOggetto As Long)
Dim sSQL As String
Dim rsFatt As DmtADOLib.adoResultset
Dim PrezzoMedio As Double
Dim Link_CategoriaMerceologica As Long
Dim PrezzoDaReg As Double
Dim Unita_progresso As Double


TIPO_QUADRATURA = ParametroTipoQuadratura

Screen.MousePointer = 11
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneRigheConf"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneLavorazione"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneVendita"

Select Case IDTipoOggetto

Case 2 'DOCUMENTO DI TRASPORTO
    sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_Data, ValoriOggettoPerTipo0002.Doc_Numero "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0004.IDValoriOggettoDettaglio=" & IDRigaVendita
    Set rsFatt = Cn.OpenResultset(sSQL)
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
        
            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                If EsistenzaArticoloConferitoElaborato(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), IDPeriodoLiquidazione) = False Then
                    ConferimentoRighe fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!RV_POIDSocio), IDPeriodoLiquidazione
                End If
            End If
            
            GET_IVAArticoloPerPrezzoMedio fnNotNullN(rsFatt!Link_Art_Articolo)
            
            GET_IVAArticolo fnNotNullN(rsFatt!IDTipoOggetto), rsFatt!IDOggetto, rsFatt!IDValoriOggettoDettaglio
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!RV_PODataConferimento))
            End If
            
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsFatt!Link_Art_Articolo))
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), Link_CategoriaMerceologica, fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "IDArticolo, IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & IDPeriodoLiquidazione & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("D.D.T. n° " & rsFatt!Doc_numero & " del " & rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & Link_Iva & ", "
                sSQL = sSQL & fnNormString(Codice_Iva) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(Iva) & ", "
                sSQL = sSQL & Link_Iva_Medio & ", "
                sSQL = sSQL & fnNormString(Codice_Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva_Medio) & ", "
                sSQL = sSQL & fnNormString(Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio)) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(Link_CategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ")"
                
                Cn.Execute sSQL
        End If
        
    rsFatt.MoveNext
    Wend
    
    rsFatt.CloseResultset
    Set rsFatt = Nothing
    
Case 114 'FATTURA ACCOMPAGNATORIA
    sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_Data, ValoriOggettoPerTipo0072.Doc_Numero "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0001.IDValoriOggettoDettaglio=" & IDRigaVendita
    
    Set rsFatt = Cn.OpenResultset(sSQL)
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
        
            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                If EsistenzaArticoloConferitoElaborato(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), IDPeriodoLiquidazione) = False Then
                    ConferimentoRighe fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!RV_POIDSocio), IDPeriodoLiquidazione
                End If
            End If
            
            GET_IVAArticoloPerPrezzoMedio fnNotNullN(rsFatt!Link_Art_Articolo)
            
            GET_IVAArticolo fnNotNullN(rsFatt!IDTipoOggetto), rsFatt!IDOggetto, rsFatt!IDValoriOggettoDettaglio
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(DataDocumento))
            End If
            
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsFatt!Link_Art_Articolo))
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), Link_CategoriaMerceologica, fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "IDArticolo, IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & IDPeriodoLiquidazione & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("D.D.T. n° " & rsFatt!Doc_numero & " del " & rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & Link_Iva & ", "
                sSQL = sSQL & fnNormString(Codice_Iva) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(Iva) & ", "
                sSQL = sSQL & Link_Iva_Medio & ", "
                sSQL = sSQL & fnNormString(Codice_Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva_Medio) & ", "
                sSQL = sSQL & fnNormString(Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio)) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(Link_CategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ")"
                
                Cn.Execute sSQL
        End If
        
        If (Unita_progresso + prg.Value) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_progresso
        End If
    rsFatt.MoveNext
    Wend
    
    rsFatt.CloseResultset
    Set rsFatt = Nothing
    
    
Case 8 'SCONTRINO NON FISCALE
    sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_Data, ValoriOggettoPerTipo0008.Doc_Numero "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0034.IDValoriOggettoDettaglio=" & IDRigaVendita
    
    Set rsFatt = Cn.OpenResultset(sSQL)
    
    While Not rsFatt.EOF
        If fnNotNullN(rsFatt!Link_Art_Articolo) > 0 Then
        
            If fnNotNullN(rsFatt!RV_POIDConferimentoRighe) > 0 Then
                If EsistenzaArticoloConferitoElaborato(fnNotNullN(rsFatt!RV_POIDConferimentoRighe), IDPeriodoLiquidazione) = False Then
                    ConferimentoRighe fnNotNullN(rsFatt!RV_POIDConferimentoRighe), fnNotNullN(rsFatt!RV_POIDSocio), IDPeriodoLiquidazione
                End If
            End If
            
            GET_IVAArticoloPerPrezzoMedio fnNotNullN(rsFatt!Link_Art_Articolo)
            
            GET_IVAArticolo fnNotNullN(rsFatt!IDTipoOggetto), rsFatt!IDOggetto, rsFatt!IDValoriOggettoDettaglio
            
            If TIPO_CALCOLO_PREZZO_MEDIO = 2 Then
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(rsFatt!Doc_data))
            Else
                PrezzoMedio = GetPrezzoUnitarioPeriodo(fnNotNullN(rsFatt!Link_Art_Articolo), fnNotNull(DataDocumento))
            End If
            
            If TIPO_IMPORTO_ARTICOLO = 1 Then
                PrezzoDaReg = fnNotNullN(rsFatt!RV_POImportoLiq)
            Else
                PrezzoDaReg = PrezzoMedio
            End If
            
            Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsFatt!Link_Art_Articolo))
            
            TrattenuteArticolo fnNotNullN(rsFatt!RV_POIDSocio), fnNotNullN(rsFatt!Link_Art_Articolo), Link_CategoriaMerceologica, fnNotNullN(rsFatt!RV_POIDTipoLavorazione), False, fnNotNullN(rsFatt!Art_quantita_totale), fnNotNullN(PrezzoDaReg)
            
                sSQL = "INSERT INTO RV_POTMPLiquidazioneVendita ("
                sSQL = sSQL & "IDArticolo, IDRV_POCaricoMerceRighe, DataConferimento, IDSocio, "
                sSQL = sSQL & "IDRV_POPeriodoLiquidazione, Quantita, IDOggetto, IDValoreOggettoDettaglio, Oggetto, "
                sSQL = sSQL & "IDTipoOggetto, DataDocumentoVendita, NumeroDocumento, ImportoUnitario, ImportoNettoTotale, ImportoMedioPeriodo, ImportoTotaleSuPeriodo, "
                sSQL = sSQL & "IDIva_Vend, CodiceIva_Vend, AliquotaIva_Vend, Iva_Vend, IDIva_Medio, CodiceIva_Medio, AliquotaIva_Medio, Iva_Medio, "
                sSQL = sSQL & "ImpostaImportoUnitario, ImpostaImportoMedio, "
                sSQL = sSQL & "ImpostaTotaleIva, ImpostaTotaleMedioIva, ImportoTotaleLordoIva, ImportoTotaleMedioLordoIva, "
                sSQL = sSQL & "PrezzoUnitarioDaReg, TotaleNettoRigaDaReg, TotaleImpostaDaReg, TotaleLordoRigaDaReg, IDCategoriaMerceologica, "
                sSQL = sSQL & "TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & fnNotNullN(rsFatt!Link_Art_Articolo) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDConferimentoRighe) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!RV_PODataConferimento) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!RV_POIDSocio) & ", "
                sSQL = sSQL & IDPeriodoLiquidazione & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!Art_quantita_totale) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDOggetto) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDValoriOggettoDettaglio) & ", "
                sSQL = sSQL & fnNormString("D.D.T. n° " & rsFatt!Doc_numero & " del " & rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNotNullN(rsFatt!IDTipoOggetto) & ", "
                sSQL = sSQL & fnNormDate(rsFatt!Doc_data) & ", "
                sSQL = sSQL & fnNormString(rsFatt!Doc_numero) & ", "
                sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & Link_Iva & ", "
                sSQL = sSQL & fnNormString(Codice_Iva) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva) & ", "
                sSQL = sSQL & fnNormString(Iva) & ", "
                sSQL = sSQL & Link_Iva_Medio & ", "
                sSQL = sSQL & fnNormString(Codice_Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(Aliquota_Iva_Medio) & ", "
                sSQL = sSQL & fnNormString(Iva_Medio) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio)) & ","
                sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * fnNotNullN(rsFatt!Art_quantita_totale)) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                Select Case TIPO_IMPORTO_ARTICOLO
                
                    Case 1
                        sSQL = sSQL & fnNormNumber(rsFatt!RV_POImportoLiq) & ", "
                        sSQL = sSQL & fnNormNumber(fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(rsFatt!RV_POImportoLiq) / 100) * Aliquota_Iva) * rsFatt!Art_quantita_totale) + (fnNotNullN(rsFatt!RV_POImportoLiq) * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                    Case 3
                        sSQL = sSQL & fnNormNumber(PrezzoMedio) & ", "
                        sSQL = sSQL & fnNormNumber(PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale)) & ", "
                        sSQL = sSQL & fnNormNumber(((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) & ", "
                        sSQL = sSQL & fnNormNumber((((fnNotNullN(PrezzoMedio) / 100) * Aliquota_Iva_Medio) * rsFatt!Art_quantita_totale) + (PrezzoMedio * fnNotNullN(rsFatt!Art_quantita_totale))) & ", "
                End Select
                sSQL = sSQL & fnNotNullN(Link_CategoriaMerceologica) & ", "
                sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
                sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ")"
                
                Cn.Execute sSQL
        End If
        
    rsFatt.MoveNext
    Wend
    
    rsFatt.CloseResultset
    Set rsFatt = Nothing

End Select
    
    START_LIQUIDAZIONE IDLiquidazione, IDPeriodoLiquidazione
Screen.MousePointer = 0
End Sub
Private Function GET_IVAArticolo(IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim sTabellaTesta As String
Dim sTabellaCorpo As String

Select Case IDTipoOggetto

    Case 114
        sTabellaTesta = "ValoriOggettoPerTipo0072"
        sTabellaCorpo = "ValoriOggettoDettaglio0001"
    Case 2
        sTabellaTesta = "ValoriOggettoPerTipo0002"
        sTabellaCorpo = "ValoriOggettoDettaglio0004"
    
    Case 107
        sTabellaTesta = "ValoriOggettoPerTipo006B"
        sTabellaCorpo = "ValoriOggettoDettaglio0007"
    Case 11
        sTabellaTesta = "ValoriOggettoPerTipo000B"
        sTabellaCorpo = "ValoriOggettoDettaglio0016"
    Case Default
        Link_Iva = 0
        Codice_Iva = ""
        Aliquota_Iva = 0
        Exit Function
End Select
If sTabellaTesta = "" Then
    Link_Iva = 0
    Codice_Iva = ""
    Aliquota_Iva = 0
    Iva = ""
    Exit Function
End If
sSQL = "SELECT Iva.IDIva, Iva.AliquotaIva, Iva.Codice, Iva_1.AliquotaIva AS AliquotaIvaCliente, Iva_1.Codice AS CodiceIvaCliente, Iva_1.IDIva AS IDIvaCliente "
sSQL = sSQL & "FROM " & sTabellaTesta & " INNER JOIN "
sSQL = sSQL & sTabellaCorpo & " ON " & sTabellaTesta & ".IDOggetto = " & sTabellaCorpo & ".IDOggetto LEFT OUTER JOIN "
sSQL = sSQL & "Iva Iva_1 ON " & sTabellaTesta & ".Link_Nom_IVA = Iva_1.IDIva LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON " & sTabellaCorpo & ".Link_Art_Iva = Iva.IDIva "
sSQL = sSQL & "WHERE " & sTabellaTesta & ".IDOggetto = " & IDOggetto & " AND "
sSQL = sSQL & sTabellaCorpo & ".IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

If rs.EOF Then
    Link_Iva = 0
    Codice_Iva = ""
    Aliquota_Iva = 0
    Iva = ""
Else
    If fnNotNullN(rs!IDIvaCliente) = 0 Then
        Link_Iva = fnNotNullN(rs!IDIva)
        Codice_Iva = fnNotNull(rs!Codice)
        Aliquota_Iva = fnNotNullN(rs!AliquotaIva)
        Iva = GetIva(Link_Iva)
    Else
        Link_Iva = fnNotNullN(rs!IDIvaCliente)
        Codice_Iva = fnNotNull(rs!CodiceIvaCliente)
        Aliquota_Iva = fnNotNullN(rs!AliquotaIvaCliente)
        Iva = GetIva(Link_Iva)
    End If
End If

rs.Close
Set rs = Nothing


End Function
Private Function GetIva(IDIva) As String
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset


sSQL = "SELECT Iva FROM Iva WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetIva = ""
Else
    GetIva = fnNotNull(rs!Iva)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Sub GET_IVAArticoloPerPrezzoMedio(IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT Iva.IDIva, Iva.Iva, Iva.AliquotaIva, Iva.Codice "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE Articolo.IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    Link_Iva_Medio = 0
    Codice_Iva_Medio = ""
    Aliquota_Iva_Medio = 0
    Iva_Medio = ""
Else
    Link_Iva_Medio = fnNotNullN(rs!IDIva)
    Codice_Iva_Medio = fnNotNull(rs!Codice)
    Aliquota_Iva_Medio = fnNotNullN(rs!AliquotaIva)
    Iva_Medio = fnNotNull(rs!Iva)

End If

rs.CloseResultset
Set rs = Nothing

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

If TotaleQuantita > 0 Then
    GetPrezzoUnitarioPeriodo = TotaleVendita / TotaleQuantita
End If


End Function
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
Private Sub ConferimentoRighe(IDRigaConferimento As Long, IDSocio As Long, IDPeriodoLiquidazione As Long)
Dim sSQL As String
Dim rsRighe As ADODB.Recordset
Dim Link_CategoriaMerceologica As Long


    sSQL = "SELECT RV_POCaricoMerceRighe.*, RV_POCaricoMerceTesta.IDAnagrafica, RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceTesta.DataDocumento, "
    sSQL = sSQL & "RV_POCaricoMerceTesta.Anagrafica, RV_POCaricoMerceTesta.Nome "
    sSQL = sSQL & "FROM RV_POCaricoMerceRighe INNER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe = " & IDRigaConferimento
    
    Set rsRighe = New ADODB.Recordset
    
    rsRighe.Open sSQL, Cn.InternalConnection
    While Not rsRighe.EOF
    
        
    Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(fnNotNullN(rsRighe!IDArticolo))
    
    GET_SCARTI rsRighe!IDAnagrafica, rsRighe!IDRV_POCaricoMerceRighe
    
    
    If ((rsRighe!DataDocumento >= DATA_INIZIO) And (rsRighe!DataDocumento <= DATA_FINE)) Then
        If TIPO_QUANTITA = 3 Then
            TrattenuteArticolo fnNotNullN(rsRighe!IDAnagrafica), rsRighe!IDArticolo, Link_CategoriaMerceologica, 0, True, fnNotNullN(rsRighe!Qta_UM), 0
        End If
    End If
        
        
        sSQL = "INSERT INTO RV_POTMPLiquidazioneRigheConf ("
        sSQL = sSQL & "IDRV_POCaricoMerceRighe, IDRV_POCaricoMerceTesta, IDRV_POPeriodoLiquidazione, IDArticolo,  "
        sSQL = sSQL & "IDImballo, Quantita,  Colli, PesoLordo, Tara, PesoNetto, Pezzi, "
        sSQL = sSQL & "Trattenuta, IDCategoriaMerceologica, IDAnagrafica, Anagrafica, Nome, "
        sSQL = sSQL & "NumeroDocumento, DataDocumento) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & rsRighe!IDRV_POCaricoMerceRighe & ", "
        sSQL = sSQL & rsRighe!IDRV_POcaricoMercetesta & ", "
        sSQL = sSQL & IDPeriodoLiquidazione & ", "
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
    
    
    
    rsRighe.Close
    Set rsRighe = Nothing


End Sub
Private Function GetPrezzoImballo(IDImballo As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsList As ADODB.Recordset

sSQL = "SELECT Articolo.IDArticolo, RV_POTipoImballo.Rendere, Articolo.RV_POImballoPerAddebito "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoImballo ON Articolo.RV_POIDTipoImballo = RV_POTipoImballo.IDRV_POTipoImballo "
sSQL = sSQL & "WHERE IDArticolo=" & IDImballo

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

If rs.EOF Then
    GetPrezzoImballo = 0
Else
    If rs!Rendere = True Then
        If rs!RV_POImballoPerAddebito = 1 Then
            sSQL = "SELECT PrezzoNettoIva FROM ListinoPerArticolo "
            sSQL = sSQL & "WHERE IDListino=" & Link_ListinoImballo & " AND "
            sSQL = sSQL & "IDArticolo=" & IDImballo
            
            Set rsList = New ADODB.Recordset
            rsList.Open sSQL, Cn.InternalConnection
            
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
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDTipoQuadratura FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroTipoQuadratura = fnNotNullN(rs!IDTipoQuadratura)
Else
    ParametroTipoQuadratura = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function EsistenzaArticoloConferitoElaborato(IDRigaConferimento As Long, IDPeriodoLiquidazione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POTMPLiquidazioneRigheConf "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDRV_POPeriodoLiquidazione=" & IDPeriodoLiquidazione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    EsistenzaArticoloConferitoElaborato = False
Else
    EsistenzaArticoloConferitoElaborato = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_SCARTI(IDSocio As Long, IDConferimentoRiga As Long)
Dim sSQL As String
Dim rsScarti As DmtADOLib.adoResultset
Dim Link_CategoriaMerceologica As Long

sSQL = "SELECT * FROM RV_POLavorazione WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rsScarti = Cn.OpenResultset(sSQL)
While Not rsScarti.EOF
    If ARTICOLI_DI_QUAD = False Then
        TrattenuteArticolo IDSocio, rsScarti!IDArticolo, 0, rsScarti!IDTipoLavorazione, True, rsScarti!Qta_UM, 0
        Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA(rsScarti!IDArticolo)
        
        sSQL = "INSERT INTO RV_POTMPLiquidazioneLavorazione ("
        sSQL = sSQL & "IDRV_POLavorazione, IDRV_POCaricoMerceRighe, IDArticolo, Quantita, "
        sSQL = sSQL & "IDCategoriaMerceologica, TrattenutePerLavorazione, TrattenuteGenerali, TrattenuteTotali) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNotNullN(rsScarti!IDRV_POLavorazione) & ", "
        sSQL = sSQL & IDConferimentoRiga & ", "
        sSQL = sSQL & fnNotNullN(rsScarti!IDArticolo) & ", "
        sSQL = sSQL & fnNormNumber(rsScarti!Qta_UM) & ", "
        sSQL = sSQL & Link_CategoriaMerceologica & ", "
        sSQL = sSQL & fnNormNumber(TrattenutePerLavorazione) & ", "
        sSQL = sSQL & fnNormNumber(TrattenuteGenerali) & ", "
        sSQL = sSQL & fnNormNumber(TrattenuteTotali) & ")"
        
        Cn.Execute sSQL
    End If

rsScarti.MoveNext
Wend


rsScarti.CloseResultset
Set rsScarti = Nothing

End Sub
Private Function GET_CATEGORIA_MERCEOLOGICA(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

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
Private Function ElaborazionePrezzoUnitarioArticoloDaLiquidare(prg As ProgressBar)


End Function
Private Sub fnCalcolaPrezzoUnitario(prg As ProgressBar)
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset
Dim rsComm As DmtADOLib.adoResultset
Dim Prezzo As Double
Dim rsOgg As DmtADOLib.adoResultset
Dim Link_Oggetto As Long
Dim Unita_progresso As Double

prg.Value = 0
prg.Max = 1000

Unita_progresso = prg.Max / TotaleRighe


''''''''''''''''''DOCUMENTO DI TRASPORTO''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0004.IDOggetto "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo0002.Doc_Data>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND ValoriOggettoPerTipo0002.Doc_Data<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1"


Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF
    
    If fnControlloEsistenzaDocTMP(fnNotNullN(rs!IDOggetto)) = False Then
         Link_Oggetto = fnNotNullN(rs!IDOggetto)
        'Prende in considerazione le righe dell'oggetto
        sSQL = "SELECT IDOggetto, IDValoriOggettoDettaglio, art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
                
        Set rsOgg = Cn.OpenResultset(sSQL)
        While Not rsOgg.EOF
            'CALCOLO DELLE COMMISSIONI SULL'IMPORTO UNITARIO
            sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
            Set rsComm = Cn.OpenResultset(sSQL)
                Prezzo = fnNotNullN(rsOgg!Art_pre_uni_net_sco_net_IVA)
                While Not rsComm.EOF
                    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
                rsComm.MoveNext
                Wend
                rsComm.CloseResultset
                Set rsComm = Nothing

                'AGGIORNAMENTO DELLA RIGA CON L'IMPORTO CALCOLATO
                sSQL = "UPDATE ValoriOggettoDettaglio0004 SET "
                sSQL = sSQL & "RV_POImportoLiq=" & fnNormNumber(Prezzo)
                sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rsOgg!IDValoriOggettoDettaglio)
                Cn.Execute sSQL

            
        rsOgg.MoveNext
        Wend
        rsOgg.CloseResultset
        Set rsOgg = Nothing
        'INSERIMENTO DELL'IDOGGETTO NELLA TABELLA TEMPORANEA
        sSQL = "INSERT INTO RV_POTMPDocEla ("
        sSQL = sSQL & "IDOggetto) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNotNullN(Link_Oggetto) & ")"
        Cn.Execute sSQL
        End If
        
    If (Unita_progresso + prg.Value) >= prg.Max Then
        prg.Value = prg.Max
    Else
        prg.Value = prg.Value + Unita_progresso
    End If
        
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''FATTURA ACCOMPAGNATORIA''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0001.IDOggetto "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo0072.Doc_Data>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND ValoriOggettoPerTipo0072.Doc_Data<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1"


Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF
    
    If fnControlloEsistenzaDocTMP(fnNotNullN(rs!IDOggetto)) = False Then
         Link_Oggetto = fnNotNullN(rs!IDOggetto)
        'Prende in considerazione le righe dell'oggetto
        sSQL = "SELECT IDOggetto, IDValoriOggettoDettaglio, art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
                
        Set rsOgg = Cn.OpenResultset(sSQL)
        While Not rsOgg.EOF
            'CALCOLO DELLE COMMISSIONI SULL'IMPORTO UNITARIO
            sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
            Set rsComm = Cn.OpenResultset(sSQL)
                Prezzo = fnNotNullN(rsOgg!Art_pre_uni_net_sco_net_IVA)
                While Not rsComm.EOF
                    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
                rsComm.MoveNext
                Wend
                rsComm.CloseResultset
                Set rsComm = Nothing

                'AGGIORNAMENTO DELLA RIGA CON L'IMPORTO CALCOLATO
                sSQL = "UPDATE ValoriOggettoDettaglio0001 SET "
                sSQL = sSQL & "RV_POImportoLiq=" & fnNormNumber(Prezzo)
                sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rsOgg!IDValoriOggettoDettaglio)
                Cn.Execute sSQL

            
        rsOgg.MoveNext
        Wend
        rsOgg.CloseResultset
        Set rsOgg = Nothing
        'INSERIMENTO DELL'IDOGGETTO NELLA TABELLA TEMPORANEA
        sSQL = "INSERT INTO RV_POTMPDocEla ("
        sSQL = sSQL & "IDOggetto) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNotNullN(Link_Oggetto) & ")"
        Cn.Execute sSQL
    End If

    If (Unita_progresso + prg.Value) >= prg.Max Then
        prg.Value = prg.Max
    Else
        prg.Value = prg.Value + Unita_progresso
    End If

rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
''''''''''''''''''SCONTRINO NON FISCALE''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0034.IDOggetto "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo0008.Doc_Data>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND ValoriOggettoPerTipo0008.Doc_Data<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1"


Set rs = Cn.OpenResultset(sSQL)
While Not rs.EOF
    
    If fnControlloEsistenzaDocTMP(fnNotNullN(rs!IDOggetto)) = False Then
         Link_Oggetto = fnNotNullN(rs!IDOggetto)
        'Prende in considerazione le righe dell'oggetto
        sSQL = "SELECT IDOggetto, IDValoriOggettoDettaglio, art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
                
        Set rsOgg = Cn.OpenResultset(sSQL)
        While Not rsOgg.EOF
            'CALCOLO DELLE COMMISSIONI SULL'IMPORTO UNITARIO
            sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rsOgg!IDOggetto)
            Set rsComm = Cn.OpenResultset(sSQL)
                Prezzo = fnNotNullN(rsOgg!Art_pre_uni_net_sco_net_IVA)
                While Not rsComm.EOF
                    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
                rsComm.MoveNext
                Wend
                rsComm.CloseResultset
                Set rsComm = Nothing

                'AGGIORNAMENTO DELLA RIGA CON L'IMPORTO CALCOLATO
                sSQL = "UPDATE ValoriOggettoDettaglio0034 SET "
                sSQL = sSQL & "RV_POImportoLiq=" & fnNormNumber(Prezzo)
                sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rsOgg!IDValoriOggettoDettaglio)
                Cn.Execute sSQL

            
        rsOgg.MoveNext
        Wend
        rsOgg.CloseResultset
        Set rsOgg = Nothing
        'INSERIMENTO DELL'IDOGGETTO NELLA TABELLA TEMPORANEA
        sSQL = "INSERT INTO RV_POTMPDocEla ("
        sSQL = sSQL & "IDOggetto) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNotNullN(Link_Oggetto) & ")"
        Cn.Execute sSQL
    End If
    
    If (Unita_progresso + prg.Value) >= prg.Max Then
        prg.Value = prg.Max
    Else
        prg.Value = prg.Value + Unita_progresso
    End If


rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing

End Sub
Private Function fnControlloEsistenzaDocTMP(IDOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDOggetto FROM RV_POTMPDocEla WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    fnControlloEsistenzaDocTMP = False
Else
    fnControlloEsistenzaDocTMP = True
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_TOTALERIGHE() As Long
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT COUNT(Movimento.IDAzienda) AS TotaleRighe "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & " WHERE DataDocumento>=" & fnNormDate(DATA_INIZIO)
sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(DATA_FINE)
sSQL = sSQL & " AND IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND IDTipoOggetto<10000 AND IDTipoOggetto<>4"



Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_TOTALERIGHE = 0
Else
    GET_TOTALERIGHE = fnNotNullN(rs!TotaleRighe)
End If

rs.CloseResultset
Set rs = Nothing

End Function
