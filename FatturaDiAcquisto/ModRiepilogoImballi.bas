Attribute VB_Name = "ModRiepilogoImballi"
Private QUANTITA_ENTRATA_PREC As Double
Private QUANTITA_USCITA_PREC As Double
Private PrimaElaborazione As Boolean

Private rsMov As ADODB.Recordset
Private rsMovPrec As ADODB.Recordset
Private rsTMP As ADODB.Recordset

Private IDTipoOggettoLav As Long
Private IDTipoOggettoScarto As Long

Private Link_TipoImballo As Long
Private DataInizioEsercizio As String
Private DataFineEsercizio As String

Private CnDMT As DmtOleDbLib.adoConnection

Private IDListinoImballo As Long


Public Sub AVVIA_PROCEDURA(fraRiepilogoImballo As Frame, cnDMTLocal As DmtOleDbLib.adoConnection, IDAnagrafica As Long, IDEsercizio As Long, ProgressBar1 As ProgressBar, lblInfoStatus As Label)
On Error GoTo ERR_AVVIA_PROCEDURA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_progresso As Double
Dim I As Integer

    If IDEsercizio = 0 Then Exit Sub
    
    fraRiepilogoImballo.Visible = True
    
    
    Set CnDMT = cnDMTLocal
    
    IDTipoOggettoLav = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    IDTipoOggettoScarto = fnGetTipoOggetto("RV_POLavorazioneL")
    Link_TipoImballo = ParametroImballo
    
    INIT_ESERCIZIO IDEsercizio

    
    sSQL = "SELECT Movimento.IDMovimento, ContatoreArticoloPerMagazzino.PartecipaGiacenza, Movimento.IDAzienda, Movimento.IDOggetto, Movimento.IDTipoOggetto, "
    sSQL = sSQL & "Movimento.IDTipoMovimento, Movimento.IDMagazzino, Movimento.IDAnagrafica, Movimento.IDTipoAnagrafica, Movimento.IDFunzione,"
    sSQL = sSQL & "Movimento.IDProcessoPerFunzione, Movimento.IDArticolo, Movimento.Oggetto, Movimento.DataMovimento, Movimento.DataDocumento,"
    sSQL = sSQL & "Movimento.QuantitaTotale , Movimento.PrezzoUnitario, Movimento.NumeroDocumento, Articolo.RV_POIDTipoImballo "
    sSQL = sSQL & "FROM Movimento LEFT OUTER JOIN "
    sSQL = sSQL & "ContatorePerProcesso ON Movimento.IDProcessoPerFunzione = ContatorePerProcesso.IDProcessoPerFunzione LEFT OUTER JOIN "
    sSQL = sSQL & "ContatoreArticoloPerMagazzino ON ContatorePerProcesso.IDContatoreArticolo = ContatoreArticoloPerMagazzino.IDContatoreArticolo AND "
    sSQL = sSQL & "Movimento.IDMagazzino = ContatoreArticoloPerMagazzino.IDMagazzino LEFT OUTER JOIN"
    sSQL = sSQL & " Articolo ON Movimento.IDArticolo = Articolo.IDArticolo"
    sSQL = sSQL & " WHERE Movimento.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Movimento.DataMovimento>=" & fnNormDate(DataInizioEsercizio)
    sSQL = sSQL & " AND Movimento.DataMovimento<=" & fnNormDate(DataFineEsercizio)
    sSQL = sSQL & " AND Movimento.IDTipoOggetto<>4"
    sSQL = sSQL & " AND Movimento.IDTipoOggetto<>303"
    sSQL = sSQL & " AND Movimento.IDTipoOggetto<>" & IDTipoOggettoLav
    sSQL = sSQL & " AND Movimento.IDTipoOggetto<>" & IDTipoOggettoScarto
    sSQL = sSQL & " AND Movimento.IDAnagrafica=" & IDAnagrafica
    sSQL = sSQL & " ORDER BY Movimento.IDMovimento"
    
    
    ProgressBar1.Value = 0
    lblInfoStatus.Caption = ""
    lblInfoStatus.Caption = "START IN CORSO...................."
    
    CnDMT.CursorLocation = adUseClient
    
    DoEvents
    
    CREA_TABELLA_TEMPORANEA
    
    Set rs = CnDMT.OpenResultset(sSQL)
    rsMov.Open , , adOpenKeyset, adLockBatchOptimistic
    rsMovPrec.Open , , adOpenKeyset, adLockBatchOptimistic
    
    While Not rs.EOF
        rsMov.AddNew
        rsMovPrec.AddNew
            For I = 0 To rsMov.Fields.Count - 1
                If rsMov.Fields(I).Name = "Segno" Then
                    If Len(Trim(fnNotNull(rs!PartecipaGiacenza))) = 0 Then
                        rsMov.Fields(I).Value = "-"
                    Else
                        rsMov.Fields(I).Value = fnNotNull(rs!PartecipaGiacenza)
                    End If
                    rsMovPrec.Fields(I).Value = rsMov.Fields(I).Value
                Else
                    rsMov.Fields(I).Value = rs.adoColumns(rsMov.Fields(I).Name).Value
                    rsMovPrec.Fields(I).Value = rs.adoColumns(rsMov.Fields(I).Name).Value
                End If
            Next
        rsMov.Update
        rsMovPrec.Update
        lblInfoStatus.Caption = "PREPARAZIONE DATI " & fnNotNullN(rs!IDMovimento)
        DoEvents
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    EliminazioneDati
    
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = 100
    
   
    
    ''''''''''''''''''''''''''CALCOLO PROGRESS PRIMARIA'''''''''''''''''''''''''''
    sSQL = "SELECT Count(IDArticolo) AS NUMERO FROM Articolo "
    sSQL = sSQL & "WHERE IDTipoProdotto=" & Link_TipoImballo
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    
    Set rsCount = CnDMT.OpenResultset(sSQL)
    
    If rsCount.EOF Then
        Unita_progresso = 0
    Else
        If fnNotNullN(rsCount!Numero) > 0 Then
            Unita_progresso = FormatNumber((ProgressBar1.Max / fnNotNullN(rsCount!Numero)), 4)
        Else
            Unita_progresso = 0
        End If
    End If
    
    rsCount.CloseResultset
    Set rsCount = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If Unita_progresso = 0 Then
        MsgBox "Non ci sono dati da elaborare", vbInformation, "Elaborazione"
        fraRiepilogoImballo.Visible = False
        Exit Sub
    End If

    'AVVIO PROCEDURA DI INSERIMENTO DATI
    sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo FROM Articolo  "
    sSQL = sSQL & "WHERE IDTipoProdotto=" & Link_TipoImballo
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

    Set rs = CnDMT.OpenResultset(sSQL)
    
    Set rsTMP = New ADODB.Recordset
    
    rsTMP.Open "SELECT * FROM RV_POTMPSaldoImballo WHERE IdUtente = " & TheApp.IDUser, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    'rsTMP.Delete
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        lblInfoStatus.Caption = "INSERIMENTO DATI DELL'IMBALLO " & UCase(fnNotNull(rs!CodiceArticolo)) & "-" & UCase(fnNotNull(rs!Articolo))
        
        DoEvents
        
        MovimentiClientiPerImballo IDEsercizio, IDAnagrafica, fnNotNullN(rs!IDArticolo), 2
        
        MovimentiFornitoriPerImballo IDEsercizio, IDAnagrafica, fnNotNullN(rs!IDArticolo)
        
        If (Unita_progresso + ProgressBar1.Value) >= ProgressBar1.Max Then
            ProgressBar1.Value = ProgressBar1.Max
        Else
            ProgressBar1.Value = ProgressBar1.Value + Unita_progresso
        End If
    DoEvents
    rs.MoveNext
    Wend
        
    rs.CloseResultset
    Set rs = Nothing
    
    rsTMP.Close
    Set rsTMP = Nothing
    
    rsMov.Close
    Set rsMov = Nothing
    
    rsMovPrec.Close
    Set rsMovPrec = Nothing
       
    CnDMT.CursorLocation = adUseServer
    
    
    lblInfoStatus.Caption = "OPERAZIONE COMPLETATA"


    fraRiepilogoImballo.Visible = False
    
Exit Sub
ERR_AVVIA_PROCEDURA:
    MsgBox Err.Description, vbCritical, "Riepilogo imballi"
    fraRiepilogoImballo.Visible = False
End Sub
Private Sub EliminazioneDati()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "DELETE FROM RV_POTMPSaldoImballo WHERE IDUtente=" & TheApp.IDUser

CnDMT.Execute sSQL
    
End Sub
Private Sub INIT_ESERCIZIO(IDEsercizio As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Esercizio "
sSQL = sSQL & "WHERE IDEsercizio=" & IDEsercizio

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    DataInizioEsercizio = fnNotNull(rs!DataInizio)
    DataFineEsercizio = fnNotNull(rs!DataFine)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CREA_TABELLA_TEMPORANEA()
    Set rsMov = New ADODB.Recordset
    
    rsMov.CursorLocation = adUseClient

    rsMov.Fields.Append "IDMovimento", adInteger
    rsMov.Fields.Append "IDTipoOggetto", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDAnagrafica", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDTipoAnagrafica", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDAzienda", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDFunzione", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDProcessoPerFunzione", adInteger, , adFldIsNullable
    rsMov.Fields.Append "IDMagazzino", adInteger, , adFldIsNullable
    rsMov.Fields.Append "DataMovimento", adDBTimeStamp, , adFldIsNullable
    rsMov.Fields.Append "QuantitaTotale", adDouble, , adFldIsNullable
    rsMov.Fields.Append "Segno", adChar, 1, adFldIsNullable
    rsMov.Fields.Append "IDTipoMovimento", adInteger, , adFldIsNullable
    rsMov.Fields.Append "Oggetto", adVarChar, 250, adFldIsNullable
    rsMov.Fields.Append "DataDocumento", adDBTimeStamp, , adFldIsNullable
    rsMov.Fields.Append "NumeroDocumento", adChar, 40, adFldIsNullable
    rsMov.Fields.Append "PrezzoUnitario", adDouble, , adFldIsNullable
    



    Set rsMovPrec = New ADODB.Recordset
    
    rsMovPrec.CursorLocation = adUseClient

    rsMovPrec.Fields.Append "IDMovimento", adInteger
    rsMovPrec.Fields.Append "IDTipoOggetto", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDAnagrafica", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDTipoAnagrafica", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDAzienda", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDFunzione", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDProcessoPerFunzione", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "IDMagazzino", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "DataMovimento", adDBTimeStamp, , adFldIsNullable
    rsMovPrec.Fields.Append "QuantitaTotale", adDouble, , adFldIsNullable
    rsMovPrec.Fields.Append "Segno", adChar, 1, adFldIsNullable
    rsMovPrec.Fields.Append "IDTipoMovimento", adInteger, , adFldIsNullable
    rsMovPrec.Fields.Append "Oggetto", adVarChar, 250, adFldIsNullable
    rsMovPrec.Fields.Append "DataDocumento", adDBTimeStamp, , adFldIsNullable
    rsMovPrec.Fields.Append "NumeroDocumento", adChar, 40, adFldIsNullable
    rsMovPrec.Fields.Append "PrezzoUnitario", adDouble, , adFldIsNullable


End Sub


Public Sub MovimentiClientiPerImballo(IDEsercizio As Long, IDAnagraficaSocio As Long, IDImballo As Long, IDTipoAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim SaldoInizialeEsercizio As Double
Dim SaldoPeriodoPrec As Double
Dim IDAnagrafica As Long

Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_progresso As Double


sSQL = "IDArticolo=" & IDImballo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoAnagrafica=2"


sSQL = sSQL & " AND DataMovimento>=#" & DataInizioEsercizio & "#"
sSQL = sSQL & " AND DataMovimento<=#" & Date & "#"
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaSocio

IDAnagrafica = 0

rsMov.Filter = sSQL



While Not rsMov.EOF

DoEvents

    If IDAnagrafica <> fnNotNullN(rsMov!IDAnagrafica) Then
        IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
        SaldoInizialeEsercizio = GET_SALDO_INIZIALE(IDImballo, IDEsercizio, fnNotNullN(rsMov!IDAnagrafica), 2)
        QUANTITA_ENTRATA_PREC = 0
        QUANTITA_USCITA_PREC = 0
        SaldoPeriodoPrec = GET_SALDO_PERIODO_PRECEDENTE(IDImballo, IDEsercizio, fnNotNullN(rsMov!IDAnagrafica), 2)
        

            rsTMP.AddNew
                rsTMP!IDArticoloImballo = IDImballo
                rsTMP!IDTipoMovimento = 0
                rsTMP!Segno = "+"
                rsTMP!QuantitaUscita = 0
                rsTMP!QuantitaEntrata = 0
                'rsTMP!DataMovimento = rsMov!DataMovimento
                rsTMP!IDTipoOggetto = 0
                rsTMP!IDOggetto = 0
                rsTMP!Oggetto = 0
                
                '    rsTMP!DataDocumento = rsMov!DataMovimento
                rsTMP!NumeroDocumento = " "
                rsTMP!IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
                rsTMP!IDTipoAnagrafica = fnNotNullN(rsMov!IDTipoAnagrafica)
                rsTMP!SaldoEsercizio = SaldoInizialeEsercizio
                rsTMP!QuantitaEntrataPrecedente = QUANTITA_ENTRATA_PREC
                rsTMP!QuantitaUscitaPrecedente = QUANTITA_USCITA_PREC
                rsTMP!IDAnagraficaClienteFiltro = IDAnagraficaSocio
                rsTMP!IDImballOFiltro = 0
                rsTMP!IDEsercizioRiferimentoFiltro = IDEsercizio
                rsTMP!DaDataFiltro = DataInizioEsercizio
                rsTMP!ADataFiltro = Date
                rsTMP!IDUtente = TheApp.IDUser
                rsTMP!ImportoUnitario = fnNotNullN(rsMov!PrezzoUnitario)
                rsTMP!IDAzienda = TheApp.IDFirm
                rsTMP!ImportoDiListino = GET_IMPORTO_LISTINO(fnNotNullN(rsMov!IDArticolo), IDListinoImballo)
                If (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoLav) Or (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoScarto) Then
                    rsTMP!MovimentoDaLavorazione = 1
                Else
                    rsTMP!MovimentoDaLavorazione = 0
                End If
                
             rsTMP.Update
        'End If
    End If
            
    If (CDate(rsMov!DataMovimento) >= CDate(DataInizioEsercizio)) And (CDate(rsMov!DataMovimento) <= CDate(Date)) Then
        rsTMP.AddNew
            rsTMP!IDArticoloImballo = IDImballo
            rsTMP!IDTipoMovimento = rsMov!IDTipoMovimento
            rsTMP!Segno = rsMov!Segno
            If rsTMP!Segno = "-" Then
                rsTMP!QuantitaUscita = fnNotNullN(rsMov!QuantitaTotale)
                rsTMP!QuantitaEntrata = 0
            Else
                rsTMP!QuantitaUscita = 0
                rsTMP!QuantitaEntrata = fnNotNullN(rsMov!QuantitaTotale)
            End If
            rsTMP!DataMovimento = rsMov!DataMovimento
            rsTMP!IDTipoOggetto = fnNotNullN(rsMov!IDTipoOggetto)
            rsTMP!IDOggetto = fnNotNullN(rsMov!IDOggetto)
            rsTMP!Oggetto = fnNotNull(rsMov!Oggetto)
            rsTMP!DataDocumento = rsMov!DataMovimento
            rsTMP!NumeroDocumento = fnNotNull(rsMov!NumeroDocumento)
            rsTMP!IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
            rsTMP!IDTipoAnagrafica = fnNotNullN(rsMov!IDTipoAnagrafica)
            rsTMP!SaldoEsercizio = SaldoInizialeEsercizio
            rsTMP!QuantitaEntrataPrecedente = QUANTITA_ENTRATA_PREC
            rsTMP!QuantitaUscitaPrecedente = QUANTITA_USCITA_PREC
            rsTMP!IDAnagraficaClienteFiltro = IDAnagraficaSocio
            rsTMP!IDImballOFiltro = 0
            rsTMP!IDEsercizioRiferimentoFiltro = IDEsercizio
            rsTMP!DaDataFiltro = DataInizioEsercizio
            rsTMP!ADataFiltro = Date
            rsTMP!IDUtente = TheApp.IDUser
            rsTMP!ImportoUnitario = fnNotNullN(rsMov!PrezzoUnitario)
            rsTMP!IDAzienda = TheApp.IDFirm
            rsTMP!ImportoDiListino = GET_IMPORTO_LISTINO(fnNotNullN(rsMov!IDArticolo), IDListinoImballo)
            If (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoLav) Or (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoScarto) Then
                rsTMP!MovimentoDaLavorazione = 1
            Else
                rsTMP!MovimentoDaLavorazione = 0
            End If
        rsTMP.Update
    End If



    
    DoEvents

rsMov.MoveNext
Wend
End Sub

Public Function GET_SALDO_INIZIALE(IDImballo As Long, IDEsercizio As Long, IDAnagrafica As Long, IDTipoAnagrafica As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT SaldoIniziale "
    sSQL = sSQL & "FROM RV_POSaldoImballoRighe RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_POSaldoImballo ON RV_POSaldoImballoRighe.IDRV_POSaldoImballo = RV_POSaldoImballo.IDRV_POSaldoImballo "
    sSQL = sSQL & "WHERE IDArticoloImballo=" & IDImballo
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica
    sSQL = sSQL & " AND IDTipoAnagrafica=" & IDTipoAnagrafica
    sSQL = sSQL & " AND IDEsercizio = " & IDEsercizio
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_SALDO_INIZIALE = 0
    Else
        GET_SALDO_INIZIALE = fnNotNullN(rs!SaldoIniziale)
    End If
    
rs.CloseResultset
Set rs = Nothing
End Function
Public Function GET_SALDO_PERIODO_PRECEDENTE(IDImballo As Long, IDEsercizio As Long, IDAnagrafica As Long, IDTipoAnagrafica As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Data_Fine_Periodo As String
Dim Segno As String

Data_Fine_Periodo = CDate(DataInizioEsercizio) - 1

'Dim test As Integer
'test = DateDiff("d", Me.txtDataInizioEsercizio.Text, Data_Fine_Periodo)
If DateDiff("d", Data_Fine_Periodo, DataInizioEsercizio) >= 0 Then
    GET_SALDO_PERIODO_PRECEDENTE = 0
    Exit Function
End If

sSQL = "IDArticolo=" & IDImballo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoAnagrafica=" & IDTipoAnagrafica
sSQL = sSQL & " AND DataMovimento>=#" & DataInizioEsercizio & "#"
sSQL = sSQL & " AND DataMovimento<=#" & Data_Fine_Periodo & "#"
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

rsMovPrec.Filter = sSQL
rsMovPrec.Sort = "IDAnagrafica"


    While Not rsMovPrec.EOF
        Segno = rsMovPrec!Segno
        If Segno = "-" Then
            QUANTITA_ENTRATA_PREC = QUANTITA_ENTRATA_PREC
            QUANTITA_USCITA_PREC = QUANTITA_USCITA_PREC + fnNotNullN(rsMovPrec!QuantitaTotale)
        Else
            QUANTITA_ENTRATA_PREC = QUANTITA_ENTRATA_PREC + fnNotNullN(rsMovPrec!QuantitaTotale)
            QUANTITA_USCITA_PREC = QUANTITA_USCITA_PREC
        End If
    rsMovPrec.MoveNext
    Wend

End Function

Public Sub MovimentiFornitoriPerImballo(IDEsercizio As Long, IDAnagraficaSocio As Long, IDImballo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim SaldoInizialeEsercizio As Double
Dim SaldoPeriodoPrec As Double
Dim IDAnagrafica As Long

Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_progresso As Double


''''''''''''''''''''''''''CALCOLO PROGRESS PRIMARIA'''''''''''''''''''''''''''
sSQL = "IDArticolo=" & IDImballo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoAnagrafica=3"
sSQL = sSQL & " AND DataMovimento>=#" & DataInizioEsercizio & "#"
sSQL = sSQL & " AND DataMovimento<=#" & Date & "#"
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaSocio
rsMov.Filter = sSQL
rsMov.Sort = "IDAnagrafica"


'Set rs = CnDMT.OpenResultset(sSQL)

IDAnagrafica = 0

While Not rsMov.EOF
    If IDAnagrafica <> fnNotNullN(rsMov!IDAnagrafica) Then

        IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
        SaldoInizialeEsercizio = GET_SALDO_INIZIALE(IDImballo, IDEsercizio, fnNotNullN(rsMov!IDAnagrafica), 3)
        QUANTITA_ENTRATA_PREC = 0
        QUANTITA_USCITA_PREC = 0
        SaldoPeriodoPrec = GET_SALDO_PERIODO_PRECEDENTE(IDImballo, IDEsercizio, fnNotNullN(rsMov!IDAnagrafica), 3)
   
        'If (QUANTITA_ENTRATA_PREC = 0 And QUANTITA_USCITA_PREC = 0) Then
        'Else
            rsTMP.AddNew
                 rsTMP!IDArticoloImballo = IDImballo
                 rsTMP!IDTipoMovimento = 0
                 rsTMP!Segno = "+"
                 rsTMP!QuantitaUscita = 0
                 rsTMP!QuantitaEntrata = 0
                 'rsTMP!DataMovimento = rsMov!DataMovimento
                 rsTMP!IDTipoOggetto = 0
                 rsTMP!IDOggetto = 0
                 rsTMP!Oggetto = 0
                 
                 '    rsTMP!DataDocumento = rsMov!DataMovimento
                 rsTMP!NumeroDocumento = " "
                 rsTMP!IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
                 rsTMP!IDTipoAnagrafica = fnNotNullN(rsMov!IDTipoAnagrafica)
                 rsTMP!SaldoEsercizio = SaldoInizialeEsercizio
                 rsTMP!QuantitaEntrataPrecedente = QUANTITA_ENTRATA_PREC
                 rsTMP!QuantitaUscitaPrecedente = QUANTITA_USCITA_PREC
                 rsTMP!IDAnagraficaClienteFiltro = IDAnagraficaSocio
                 rsTMP!IDImballOFiltro = 0
                 rsTMP!IDEsercizioRiferimentoFiltro = IDEsercizio
                 rsTMP!DaDataFiltro = DataInizioEsercizio
                 rsTMP!ADataFiltro = Date
                 rsTMP!IDUtente = TheApp.IDUser
                 rsTMP!ImportoUnitario = fnNotNullN(rsMov!PrezzoUnitario)
                 rsTMP!IDAzienda = TheApp.IDFirm
                 rsTMP!ImportoDiListino = GET_IMPORTO_LISTINO(fnNotNullN(rsMov!IDArticolo), IDListinoImballo)
                If (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoLav) Or (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoScarto) Then
                    rsTMP!MovimentoDaLavorazione = 1
                Else
                    rsTMP!MovimentoDaLavorazione = 0
                End If
             rsTMP.Update
        'End If
    End If
    If (CDate(rsMov!DataMovimento) >= CDate(DataInizioEsercizio)) And (CDate(rsMov!DataMovimento) <= CDate(Date)) Then
       rsTMP.AddNew
            rsTMP!IDArticoloImballo = IDImballo
            rsTMP!IDTipoMovimento = rsMov!IDTipoMovimento
            rsTMP!Segno = rsMov!Segno
            If rsTMP!Segno = "-" Then
                rsTMP!QuantitaUscita = fnNotNullN(rsMov!QuantitaTotale)
                rsTMP!QuantitaEntrata = 0
            Else
                rsTMP!QuantitaUscita = 0
                rsTMP!QuantitaEntrata = fnNotNullN(rsMov!QuantitaTotale)
            End If
            rsTMP!DataMovimento = fnNotNull(rsMov!DataMovimento)
            rsTMP!IDTipoOggetto = fnNotNullN(rsMov!IDTipoOggetto)
            rsTMP!IDOggetto = fnNotNullN(rsMov!IDOggetto)
            rsTMP!Oggetto = fnNotNull(rsMov!Oggetto)
            rsTMP!DataDocumento = rsMov!DataMovimento
            rsTMP!NumeroDocumento = fnNotNull(rsMov!NumeroDocumento)
            rsTMP!IDAnagrafica = fnNotNullN(rsMov!IDAnagrafica)
            rsTMP!IDTipoAnagrafica = fnNotNullN(rsMov!IDTipoAnagrafica)
            rsTMP!SaldoEsercizio = SaldoInizialeEsercizio
            rsTMP!QuantitaEntrataPrecedente = QUANTITA_ENTRATA_PREC
            rsTMP!QuantitaUscitaPrecedente = QUANTITA_USCITA_PREC
            rsTMP!IDAnagraficaClienteFiltro = IDAnagraficaSocio
            rsTMP!IDImballOFiltro = 0
            rsTMP!IDEsercizioRiferimentoFiltro = IDEsercizio
            rsTMP!DaDataFiltro = DataInizioEsercizio
            rsTMP!ADataFiltro = Date
            rsTMP!IDUtente = TheApp.IDUser
            rsTMP!ImportoUnitario = fnNotNullN(rsMov!PrezzoUnitario)
            rsTMP!IDAzienda = TheApp.IDFirm
            rsTMP!ImportoDiListino = GET_IMPORTO_LISTINO(fnNotNullN(rsMov!IDArticolo), IDListinoImballo)
            If (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoLav) Or (fnNotNullN(rsTMP!IDTipoOggetto) = IDTipoOggettoScarto) Then
                rsTMP!MovimentoDaLavorazione = 1
            Else
                rsTMP!MovimentoDaLavorazione = 0
            End If
        rsTMP.Update
    End If
    DoEvents

rsMov.MoveNext
Wend

End Sub
Private Function GET_IMPORTO_LISTINO(IDArticolo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDListinoDefault As Long
    
sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDListino = " & IDListino
sSQL = sSQL & " AND IDArticolo = " & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_LISTINO = 0
Else
    GET_IMPORTO_LISTINO = fnNotNullN(rs!PrezzoNettoIVA)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function ParametroImballo() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    ParametroImballo = rs!IDTipoImballo
Else
    ParametroImballo = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function


