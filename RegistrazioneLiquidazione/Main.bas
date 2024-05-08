Attribute VB_Name = "ModMain"
Option Explicit

'**+
'Nome                   : Main
'Parametri              : Nessuno
'Valori di ritorno      : Nessuno
'Funzionalità           : In questa Sub vengono eseguite
'                           : le operazioni di Startup.
'**/
Sub Main()
    Dim Proc  As DMTRunAppLib.Process
    Dim ErrorMsg As String

    On Error GoTo ErrorHandler
    
    'L'Applicazione.
    Set TheApp = New Application
    
    'Il nome della applicazione.
    TheApp.Name = App.EXEName
    
    
    DmtRegistry2.EXEName = App.EXEName
    
    'L'oggetto che si occupa della lettura delle risorse.
    Set gResource = New Resource
    'FrmMain.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    'Carica il form del Wizard senza mostrarlo.
    Load FrmMain
    
    Set FrmMain.Application = TheApp
    'Esegue l'applicazione
    TheApp.Run FrmMain.hWnd
    
    
    'Inizializza l'oggetto Semaforo per la gestione dei conflitti di multiutenza.
    InitSemaphore
    
    
    'Viene individuato il nome della funzione.
    Application_Name = TheApp.FunctionName
    
    'Lettura file di help
    'App.HelpFile = TheApp.Path & "\Diamante.hlp"


    '----------------------------------------------------------
    'Ciclo sui processi della funzione
    '----------------------------------------------------------
    For Each Proc In TheApp.Processes
        
        'L'identificativo del Tipo Oggetto correntemente gestito.
        Current_Process_ID = Proc.IDocType.ID
        
        '..............................................................................................................................
        'Gestione della Semaforo
        '..............................................................................................................................
        
        ' Verifica se l'applicazione corrente è bloccata da altri gestori.
        ' (Il controllo avviene sul Tipo Oggetto correntemente trattato ovvero Current_Process_ID)
        If Not gSemaphore.IsActionAvailable(Current_Process_ID, SemAllObjects, SemAllActions) Then
            '-------------------------------------------------------------
            'Il programma è bloccato da un'altra manutenzione in esecuzione.
            'Sarà pertanto terminato.
            '-------------------------------------------------------------
    
            'Scarica il form
            Unload FrmMain
    
            'Termina il programma
            End
        End If
        
        
        '----------------------------------------------------
        'Il programma non è bloccato e prosegue normalmente.
        '----------------------------------------------------
        
        'Ripulisce la tabella semaforo.
        'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
        SemaphoreUnlock
        
        'Imposta gli eventuali blocchi (semaforo) su altre manutenzioni.
        SemaphoreLock
        
        '..............................................................................................................................
        '..............................................................................................................................
        
        
        
        
        '-------------------------------------------------------------------------------------
        'In funzione del processo da gestire la manutenzione si deve comportare di conseguenza
        '-------------------------------------------------------------------------------------
        Select Case Proc.Name
        
            '*
            'Inserire qui il codice per la gestione del processo (o dei processi)
            '*
            
            Case "Manutenzione"  ' <---Tipicamente è questo l'unico processo gestito
            
            '   For Each Parameter In Proc.Parameters
            '       Select Case Parameter.Name
            '       *
            '       Inserire il codice per la gestione del parametro
            '       *
            '       Case ParameterName??????
            '       End Select
            '   next
                  
                '-------- Di solito --------
                
                'Inizializzazioni preliminari
                'FrmMaintControlli
                FrmMain.ConnessioneADO
                'Viene mostrato il form.
                FrmMain.Show
    
                
                
            Case Else
                ErrorMsg = "No processes to execute" & vbCrLf
                ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
                '*
                '/////////////////////////////////////////////////////
                'Inserire i processi che l'applicazione sa eseguire
                '/////////////////////////////////////////////////////
                '*
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE & vbCrLf
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_EXTENDED_DATABASE & vbCrLf
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_DA_SHELL & vbCrLf
                Err.Raise ERR_NO_PROCESSES, , ErrorMsg
        End Select
    Next


    
    Exit Sub
ErrorHandler:

    'Ripulisce la tabella Semaforo
    SemaphoreUnlock
    
    'Scarica il form
    Unload FrmMain
    
    If Err.Number = 1 + vbObjectError Then
        'Questo programma può essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avrà TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        Err.Raise Err.Number
    End If
    
End Sub





'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitSemaphore
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializzazione del semaforo per la gestione
'                  dei conflitti in caso di multiutenza

'
'**/
Private Sub InitSemaphore()
    
    Set gSemaphore = New Semaforo.dmtSemaphore
    Set gSemaphore.Database = TheApp.Database.Connection
    Set gSemaphore.objRes = gResource
    gSemaphore.IDUser = TheApp.IDUser
    gSemaphore.IDBranch = TheApp.Branch
    gSemaphore.IDFunction = TheApp.FunctionID
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreLock
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                 ////////////////////////////////////////////////////////////////////////
'                     Impostare qui gli eventuali blocchi sulle altre manutenzioni
'                 ////////////////////////////////////////////////////////////////////////
'**/
Public Sub SemaphoreLock()

    If Not gSemaphore Is Nothing Then
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions

        'Decommentare questa riga se si deve impedire ad un altro utente di entrare nella manutenzione corrente.
'        gSemaphore.SetObjectAction Current_Process_ID, SemAllObjects, SemAllActions

    End If
    
End Sub

'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreUnlock
'
'Parametri:
'
'Valori di ritorno:

'Funzionalità:
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'                     Sbloccare qui le altre manutenzioni (bloccate precedentemente in SemaphoreLock)
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Public Sub SemaphoreUnlock()

    If Not gSemaphore Is Nothing Then
    
        'Ripulisce la tabella semaforo per quanto riguarda il Tipo Oggetto e l'utente correnti
        gSemaphore.ClearObjectAction Current_Process_ID, SemAllObjects, SemAllActions
        
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Sblocca le manutenzioni bloccate precedentemente
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions
    
        'Decommentare questa riga se in SemaphoreLock è stato fatto altrettanto.
'        gSemaphore.ClearObjectAction Current_Process_ID, SemAllObjects, SemAllActions
    
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    'Sblocca gli eventuali gestori bloccati da questa manutenzione
    SemaphoreUnlock
    
    '--------------------------------
    'Distruzione degli oggetti allocati.
    '--------------------------------
    
    Set gSemaphore = Nothing
    
    
End Sub
Public Sub PrelevaAzienda()

    
    
End Sub
Private Function fncEsercizio() As String
    fncEsercizio = 0
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDEsercizio, Esercizio"
    sSQL = sSQL & " FROM Esercizio"
    sSQL = sSQL & " WHERE (IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (IDTipoEsercizio = 1)"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        VarIDEsercizio = rs!IDEsercizio
        fncEsercizio = rs!Esercizio
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function SalvaPeriodo(Modality As Integer) As Long
On Error GoTo ERR_SalvaPeriodo
Dim sSQL As String
Dim Link As Long


If Modality = 1 Then
    
    sSQL = "INSERT INTO RV_POLiquidazionePeriodo ("
    sSQL = sSQL & "IDRV_POLiquidazionePeriodo, Periodo,  DataInizio, DataFine, NumeroLiquidazione, "
    sSQL = sSQL & "TrattenutaRigaImporto, TrattenutaRigaPercentuale, "
    sSQL = sSQL & "IDTipoImportoArticolo, IDSocio, IDTipoImportoDocumento, IDTipoQuantita, "
    sSQL = sSQL & "ArticoliDiQuadratura, IDTipoLiquidazione, NumeroProtInt, IDAzienda, IDFIliale, IDCategoriaLiquidazione) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & LINK_PERIODO & ", "
    sSQL = sSQL & fnNormString("Liquidazione soci  dal " & DATA_INIZIO & " al " & DATA_FINE) & ", "
    sSQL = sSQL & fnNormDate(DATA_INIZIO) & ", "
    sSQL = sSQL & fnNormDate(DATA_FINE) & ", "
    sSQL = sSQL & fnNormNumber(NUMERO_LIQUIDAZIONE) & ", "
    sSQL = sSQL & fnNormNumber(TRATTENUTA_PER_IMPORTO) & ", "
    sSQL = sSQL & fnNormNumber(TRATTENUTA_PER_PERCENTUALE) & ", "
    sSQL = sSQL & fnNormNumber(TIPO_IMPORTO_ARTICOLO) & ", "
    sSQL = sSQL & fnNormNumber(LINK_SOCIO) & ", "
    sSQL = sSQL & fnNormNumber(TIPO_IMPORTO_DOCUMENTO) & ", "
    sSQL = sSQL & fnNormNumber(TIPO_QUANTITA) & ", "
    sSQL = sSQL & fnNormBoolean(ARTICOLI_DI_QUAD) & ", "
    sSQL = sSQL & fnNormNumber(TIPO_LIQUIDAZIONE) & ", "
    sSQL = sSQL & fnNormNumber(NUMERO_PROTOCOLLO) & ", "
    sSQL = sSQL & TheApp.IDFirm & ", "
    sSQL = sSQL & TheApp.Branch & ", "
    sSQL = sSQL & LINK_CAT_MERCE & ")"
   
Else
    sSQL = "UPDATE RV_POLiquidazionePeriodo SET "
    sSQL = sSQL & "Periodo=" & fnNormString("Liquidazione soci  dal " & DATA_INIZIO & " al " & DATA_FINE) & ", "
    sSQL = sSQL & "DataInizio = " & fnNormDate(DATA_INIZIO) & ", "
    sSQL = sSQL & "DataFine =" & fnNormDate(DATA_FINE) & ", "
    sSQL = sSQL & "NumeroLiquidazione=" & fnNormNumber(NUMERO_LIQUIDAZIONE) & ", "
    sSQL = sSQL & "NumeroProtInt=" & fnNormNumber(NUMERO_PROTOCOLLO) & ", "
    sSQL = sSQL & "IDTipoImportoArticolo=" & fnNormNumber(TIPO_IMPORTO_ARTICOLO) & ", "
    sSQL = sSQL & "IDSocio=" & fnNormNumber(LINK_SOCIO) & ", "
    sSQL = sSQL & "IDTipoImportoDocumento=" & fnNormNumber(TIPO_IMPORTO_DOCUMENTO) & ", "
    sSQL = sSQL & "IDTipoQuantita=" & fnNormNumber(TIPO_QUANTITA) & ", "
    sSQL = sSQL & "IDTipoLiquidazione=" & fnNormNumber(TIPO_LIQUIDAZIONE) & ", "
    sSQL = sSQL & "TrattenutaRigaImporto=" & fnNormNumber(TRATTENUTA_PER_IMPORTO) & ", "
    sSQL = sSQL & "TrattenutaRigaPercentuale=" & fnNormNumber(TRATTENUTA_PER_PERCENTUALE) & ", "
    sSQL = sSQL & "ArticoliDiQuadratura=" & fnNormBoolean(ARTICOLI_DI_QUAD) & " "
    sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & LINK_PERIODO
    
   
End If


CnDMT.Execute sSQL
Exit Function
ERR_SalvaPeriodo:
    SalvaPeriodo = 0

End Function
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    Dim VarData As String
    
    
    
    
    'Monta la query SQL per trovare il massimo valore della chiave primaria
    sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & Tabella ' & " WHERE " & >=" & VarData
    
    'Apertura del recordset
    Set rs = CnDMT.OpenResultset(fnAnsi2Jet(sSQL))
    
    'Determina il primo progressivo disponibile
    fnGetNewKey = fnNotNullN(rs.adoColumns("MaxID")) + 1
    If fnGetNewKey <= 0 Then fnGetNewKey = 1

    'Chiude il recordset e distrugge l'oggetto.
    rs.CloseResultset
    Set rs = Nothing
    
End Function


Public Sub ParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSocio = fnNotNullN(rs!IDCategoriaAnagrafica)
Else
    Link_TipoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Public Sub fnCalcolaPrezzoUnitario(prg As ProgressBar, TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsComm As DmtOleDbLib.adoResultset
Dim Prezzo As Double
Dim rsOgg As DmtOleDbLib.adoResultset
Dim Link_Oggetto As Long
Dim Unita_Progresso As Double
Const GiorniDiCalcolo As Long = 80
Dim DataFineCalcolo As String
Dim Moltiplicatore As Double
Dim ImportoCommissioni As Double
Dim MerceNettaPerLiquidazione As Double
Dim IDUMCoop_ArtVenduto As Long
Dim PrezzoScontato As Double
Dim QtaLiqPrec As Double

'If (TRATTENUTA_PER_IMPORTO = 0) And (TRATTENUTA_PER_PERCENTUALE = 0) Then Exit Sub

DataFineCalcolo = CDate(DataFine) + GiorniDiCalcolo

FrmNuovoPeriodo.List1.AddItem "CALCOLO PREZZO UNITARIO DI LIQUIDAZIONE DEI DOCUMENTI DI VENDITA"
DoEvents

''''''''''''''''''DOCUMENTO DI TRASPORTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, "
sSQL = sSQL & "ValoriOggettoPerTipo0002.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON dbo.ValoriOggettoDettaglio0004.IDOggetto = dbo.ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0004.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0004.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0002.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
'If LINK_SOCIO > 0 Then
'    sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POIDSocio=" & LINK_SOCIO
'End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga = 1"
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0004.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100

    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    
    While Not rs.EOF
        If fnNotNullN(rs!IDValoriOggettoDettaglio) = 496676 Then
            MsgBox "STOP"
        End If
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il D.D.T. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il D.D.T. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        QtaLiqPrec = fnNotNullN(rs!RV_POQuantitaLiq)
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        IDUMCoop_ArtVenduto = 0
        
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            If (IDUMCoop_ArtVenduto > 0) Then
                
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select

                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                        rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                    End If
                End If
            End If
        End If
        
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        
        
        PrezzoScontato = Prezzo
        'MerceNettaPerLiquidazione = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
        
        Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))
        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
        
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0 Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        
        rs!RV_POImportoLiq = Prezzo
        
        rs.Update

        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''DOCUMENTO DI TRASPORTO PER IV GAMMA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, "
sSQL = sSQL & "ValoriOggettoPerTipo0002.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0002.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON dbo.ValoriOggettoDettaglio0004.IDOggetto = dbo.ValoriOggettoPerTipo0002.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0004.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0004.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0002.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce>0 "
sSQL = sSQL & " AND RV_POIDProcessoIVGamma>0 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1"
If (TipoLiquidazione = 1) Then
    sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0004.IDOggetto"

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    
    While Not rs.EOF
        
'            If fnNotNullN(rs!IDOggetto) = 199243 Then
'                MsgBox "STOP"
'            End If
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il D.D.T. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il D.D.T. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        IDUMCoop_ArtVenduto = 0
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            If (IDUMCoop_ArtVenduto > 0) Then
                
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select

                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
        End If
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        
        PrezzoScontato = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
        
        Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))
        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
        
'        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
'            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
'        End If
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        
        rs!RV_POImportoLiq = Prezzo
        
        rs.Update
        

        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''FATTURA ACCOMPAGNATORIA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, "
sSQL = sSQL & "ValoriOggettoPerTipo0072.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON dbo.ValoriOggettoDettaglio0001.IDOggetto = dbo.ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0001.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0001.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0072.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
'If LINK_SOCIO > 0 Then
'    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POIDSocio=" & LINK_SOCIO
'End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga = 1"

sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0001.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100

    Link_Oggetto = 0
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
     
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il F.A. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il F.A. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
     
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        IDUMCoop_ArtVenduto = 0
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        
            If IDUMCoop_ArtVenduto > 0 Then
                
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
                
                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
        End If
        
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        
        PrezzoScontato = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
        
        Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))

        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
        
'        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
'            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
'        End If
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

''''''''''''''''''FATTURA ACCOMPAGNATORIA CON IV GAMMA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero , "
 sSQL = sSQL & "ValoriOggettoPerTipo0072.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0072.RV_PODataCompetenzaLiq "
 sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
 sSQL = sSQL & "ValoriOggettoPerTipo0072 ON dbo.ValoriOggettoDettaglio0001.IDOggetto = dbo.ValoriOggettoPerTipo0072.IDOggetto INNER JOIN "
 sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0001.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0001.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
 sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0072.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
 sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
 sSQL = sSQL & " AND RV_POIDAssegnazioneMerce>0 "
 sSQL = sSQL & " AND RV_POIDProcessoIVGamma>0 "
 sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga = 1"
 If (TipoLiquidazione = 1) Then
    sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DataFineCalcolo)
 Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
 End If
 
 sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0001.IDOggetto"
 

 Set rs = New ADODB.Recordset
 rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
 
 If rs.EOF = False Then
     prg.Value = 0
     prg.Max = 100
 
     Link_Oggetto = 0
     Unita_Progresso = prg.Max / rs.RecordCount
     DoEvents
     While Not rs.EOF
      
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il F.A. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il F.A. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        IDUMCoop_ArtVenduto = 0
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            If IDUMCoop_ArtVenduto > 0 Then
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
    

                
                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
        End If
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        

    
        PrezzoScontato = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
        
        Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))

        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
        
'        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
'            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
'        End If
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        
        rs!RV_POImportoLiq = Prezzo
        rs.Update

        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
     rs.MoveNext
     
     Wend
 
 End If
 rs.Close
 Set rs = Nothing

''''''''''''''''''SCONTRINO NON FISCALE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, "
sSQL = sSQL & "ValoriOggettoPerTipo0008.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON dbo.ValoriOggettoDettaglio0034.IDOggetto = dbo.ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0034.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0034.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0008.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
'If LINK_SOCIO > 0 Then
'    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POIDSocio=" & LINK_SOCIO
'End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga = 1"

sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0034.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il S.N.F. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il S.N.F. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        IDUMCoop_ArtVenduto = 0
        
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            If IDUMCoop_ArtVenduto > 0 Then
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
                
                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
        End If
        
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        
        PrezzoScontato = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
                
        Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))
        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
        
'        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
'            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
'        End If
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        rs!RV_POImportoLiq = Prezzo
        rs.Update
 
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        
        DoEvents
        
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing
''''''''''''''''''SCONTRINO NON FISCALE IV GAMMA''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero , "
sSQL = sSQL & "ValoriOggettoPerTipo0008.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo0008.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON dbo.ValoriOggettoDettaglio0034.IDOggetto = dbo.ValoriOggettoPerTipo0008.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0034.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0034.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo0008.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce>0 "
sSQL = sSQL & " AND RV_POIDProcessoIVGamma>0 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga = 1"
If (TipoLiquidazione = 1) Then
    sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0034.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il S.N.F. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il S.N.F. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))

        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        IDUMCoop_ArtVenduto = 0
        
        If RICALCOLA_TUTTO = 1 Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            
            If IDUMCoop_ArtVenduto > 0 Then
            
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
    

                
                rs!RV_POImportoDaLiq = 0
                If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
                    rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
        End If
        
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        
        PrezzoScontato = Prezzo
        
        If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore, rs, PrezzoScontato)
                

                Prezzo = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaLiq))
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If
        
        If AGG_COSTO_CONFEZ_PRZ_LIQ = 1 Then
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoConfezioneImballoLiq)
        End If
'
'        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
'            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
'        End If
        If AGG_COSTO_KIT_PRZ_LIQ = 1 Then
            If RICALCOLA_TUTTO = 1 Then
                If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                    rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
                End If
            End If
            Prezzo = Prezzo - fnNotNullN(rs!RV_POCostoKitLiq)
        End If
        
        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
            Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
        End If
        
        rs!RV_POImportoLiq = Prezzo
        rs.Update

        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        
        DoEvents
        
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

''''''''''''''''''NOTA DI CREDITO''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0016.*, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero , "
sSQL = sSQL & "ValoriOggettoPerTipo000B.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo000B.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON dbo.ValoriOggettoDettaglio0016.IDOggetto = dbo.ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0016.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0016.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo000B.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0016.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il N.C. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il N.C. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
    
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        If (fnNotNullN(rs!RV_PORigaRiscontroPeso) = 1) Then
            Prezzo = Abs(Prezzo)
        End If
        IDUMCoop_ArtVenduto = 0
        If (RICALCOLA_TUTTO = 1) Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            
            If IDUMCoop_ArtVenduto > 0 Then
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select

            End If
        End If
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        

        PrezzoScontato = Prezzo
            
        rs!RV_POImportoRigaCommissioni = 0
        rs!RV_POImportoDaLiq = 0
        
        If (fnNotNullN(rs!RV_POIDTipoVariazione) = 3) Or (fnNotNullN(rs!RV_POIDTipoVariazione) = 2) Then
            If (NON_CALC_INCIDENZA_IMB = 0) Then Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNull(rs!RV_POCodiceLotto), rs)
        End If
        
        If (NON_CALC_COMM = 0) Then Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), Moltiplicatore, rs, PrezzoScontato, 11)
        'Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore)
        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If

        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)

        'Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        'Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        If NO_CALC_PRELIQ_NC = 0 Then
            If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
                Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
            End If
        End If
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

'''''''''''''''''NOTA DI CREDITO IV GAMMA''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0016.*, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero , "
sSQL = sSQL & "ValoriOggettoPerTipo000B.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo000B.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON dbo.ValoriOggettoDettaglio0016.IDOggetto = dbo.ValoriOggettoPerTipo000B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0016.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0016.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo000B.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce>0 "
sSQL = sSQL & " AND RV_POIDProcessoIVGamma>0 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0016.RV_POTipoRiga = 1"

If (TipoLiquidazione = 1) Then
    sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If

sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0016.IDOggetto"

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il N.C. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il N.C. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        If (fnNotNullN(rs!RV_PORigaRiscontroPeso) = 1) Then
            Prezzo = Abs(Prezzo)
        End If
        IDUMCoop_ArtVenduto = 0
        
        If (RICALCOLA_TUTTO = 1) Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            
            If IDUMCoop_ArtVenduto > 0 Then
            
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
                

            End If
        End If
        MerceNettaPerLiquidazione = Prezzo
        
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
        PrezzoScontato = Prezzo


        rs!RV_POImportoRigaCommissioni = 0
        rs!RV_POImportoDaLiq = 0
        
        If (fnNotNullN(rs!RV_POIDTipoVariazione) = 3) Or (fnNotNullN(rs!RV_POIDTipoVariazione) = 2) Then
            If (NON_CALC_INCIDENZA_IMB = 0) Then Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNull(rs!RV_POCodiceLotto), rs)
        End If
        
        If (NON_CALC_COMM = 0) Then Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), Moltiplicatore, rs, PrezzoScontato, 11)
        'Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore)

        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If

        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)

        'Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        'Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        If NO_CALC_PRELIQ_NC = 0 Then
            If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
                Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
            End If
        End If
        rs!RV_POImportoLiq = Prezzo
        
        rs.Update

        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    Wend
End If
rs.Close
Set rs = Nothing

''''''''''''''''''NOTA DI DEBITO''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0007.*, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero , "
sSQL = sSQL & "ValoriOggettoPerTipo006B.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo006B.RV_PODataCompetenzaLiq "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo006B ON dbo.ValoriOggettoDettaglio0007.IDOggetto = dbo.ValoriOggettoPerTipo006B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0007.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0007.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo006B.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm

If TipoLiquidazione = 1 Then
    sSQL = sSQL & " AND RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0007.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il N.D. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il N.D. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))

        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        If (fnNotNullN(rs!RV_PORigaRiscontroPeso) = 1) Then
            Prezzo = Abs(Prezzo)
        End If
        IDUMCoop_ArtVenduto = 0
        If (RICALCOLA_TUTTO = 1) Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            
            If IDUMCoop_ArtVenduto > 0 Then
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select

            End If
        End If
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If

        PrezzoScontato = Prezzo
        
        rs!RV_POImportoRigaCommissioni = 0
        rs!RV_POImportoDaLiq = 0
        
        If (fnNotNullN(rs!RV_POIDTipoVariazione) = 3) Or (fnNotNullN(rs!RV_POIDTipoVariazione) = 2) Then
            If (NON_CALC_INCIDENZA_IMB = 0) Then Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNull(rs!RV_POCodiceLotto), rs)
        End If
        
        If (NON_CALC_COMM = 0) Then Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), Moltiplicatore, rs, PrezzoScontato, 107)
        'Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore)

        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If

        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)

        'Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        'Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        If NO_CALC_PRELIQ_ND = 0 Then
            If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
                Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
            End If
        End If
        rs!RV_POImportoLiq = Prezzo
        rs.Update

        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    Wend
End If
rs.Close
Set rs = Nothing

''''''''''''''''''NOTA DI DEBITO IV GAMMA''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0007.*, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero , "
sSQL = sSQL & "ValoriOggettoPerTipo006B.Link_Nom_anagrafica, RV_POConfigurazioneCliente.NonCalcolareTrattPerLiq, ValoriOggettoPerTipo006B.RV_PODataCompetenzaLiq  "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo006B ON dbo.ValoriOggettoDettaglio0007.IDOggetto = dbo.ValoriOggettoPerTipo006B.IDOggetto INNER JOIN "
sSQL = sSQL & "Oggetto ON dbo.ValoriOggettoDettaglio0007.IDOggetto = dbo.Oggetto.IDOggetto AND dbo.ValoriOggettoDettaglio0007.IDTipoOggetto = dbo.Oggetto.IDTipoOggetto LEFT OUTER JOIN "
sSQL = sSQL & "RV_POConfigurazioneCliente ON dbo.ValoriOggettoPerTipo006B.Link_Nom_anagrafica = dbo.RV_POConfigurazioneCliente.IDAnagrafica AND dbo.Oggetto.IDAzienda = dbo.RV_POConfigurazioneCliente.IDAzienda "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce>0 "
sSQL = sSQL & " AND RV_POIDProcessoIVGamma>0 "
sSQL = sSQL & " AND ValoriOggettoDettaglio0007.RV_POTipoRiga = 1"
If (TipoLiquidazione = 1) Then
    sSQL = sSQL & " AND RV_PODataLavorazione>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataLavorazione<=" & fnNormDate(DataFineCalcolo)
Else
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(DataFineCalcolo)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0007.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    prg.Value = 0
    prg.Max = 100
    
    Link_Oggetto = 0
    
    Unita_Progresso = prg.Max / rs.RecordCount
    DoEvents
    While Not rs.EOF
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
            FrmNuovoPeriodo.lblInfoStatus.Caption = "Elaborazione importo unitario di liquidazione per il N.D. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.AddItem "- Elaborazione importo unitario di liquidazione per il N.D. n° " & rs!doc_numero & " del " & rs!doc_data
            FrmNuovoPeriodo.List1.ListIndex = FrmNuovoPeriodo.List1.ListCount - 1
            DoEvents
        End If
        
        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
'
'        Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
        Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
        
        If (fnNotNullN(rs!RV_PORigaRiscontroPeso) = 1) Then
            Prezzo = Abs(Prezzo)
        End If
        IDUMCoop_ArtVenduto = 0
        If (RICALCOLA_TUTTO = 1) Then
            IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_Articolo))
            
            If IDUMCoop_ArtVenduto > 0 Then
                rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
                
                Select Case IDUMCoop_ArtVenduto
                    Case 1
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!art_numero_colli) * Moltiplicatore
                    Case 2
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
                    Case 3
                        rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
                    Case 4
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
                    Case 5
                        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
                End Select
            End If
        End If
        If (IDUMCoop_ArtVenduto > 0) Then
            If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
                Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
            Else
                Prezzo = Prezzo / Moltiplicatore
            End If
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If

        PrezzoScontato = Prezzo
        
        rs!RV_POImportoRigaCommissioni = 0
        rs!RV_POImportoDaLiq = 0
        
        If (fnNotNullN(rs!RV_POIDTipoVariazione) = 3) Or (fnNotNullN(rs!RV_POIDTipoVariazione) = 2) Then
            If (NON_CALC_INCIDENZA_IMB = 0) Then Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNull(rs!RV_POCodiceLotto), rs)
        End If
        
        If (NON_CALC_COMM = 0) Then Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), Moltiplicatore, rs, PrezzoScontato, 107)
        'Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), Moltiplicatore)
        
        If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
            Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
        Else
            Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
        End If

        Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)

        'Prezzo = Prezzo - (TRATTENUTA_PER_IMPORTO / Moltiplicatore)
        'Prezzo = Prezzo - ((Prezzo / 100) * TRATTENUTA_PER_PERCENTUALE)
        
        If NO_CALC_PRELIQ_ND = 0 Then
            If (fnNotNullN(rs!NonCalcolareTrattPerLiq) = 0) Then
                Prezzo = GET_TRATT_PRE_LIQ_ART(Prezzo, fnNotNullN(rs!Link_Art_Articolo))
            End If
        End If
        rs!RV_POImportoLiq = Prezzo
        rs.Update

        
        If (prg.Value + Unita_Progresso) >= prg.Max Then
            prg.Value = prg.Max
        Else
            prg.Value = prg.Value + Unita_Progresso
        End If
        DoEvents
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

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
Private Function fnControlloEsistenzaDocTMP(IDOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto FROM RV_POTMPDocEla WHERE IDOggetto=" & IDOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    fnControlloEsistenzaDocTMP = False
Else
    fnControlloEsistenzaDocTMP = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_COMMISSIONI_DOCUMENTO(PrezzoLiquidazione As Double, IDOggetto As Long, IDTipoOggetto As Long, Moltiplicatore As Double, rsTmp As ADODB.Recordset, PrezzoLiquidazioneScontato As Double, Optional IDTipoOggettoND_NC As Long = 2) As Double
Dim sSQL As String
Dim rsComm As DmtOleDbLib.adoResultset

GET_COMMISSIONI_DOCUMENTO = 0

'sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
'sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
'sSQL = sSQL & " AND ((IDRV_POTipoPedana=0) OR (IDRV_POTipoPedana IS NULL))"
'sSQL = sSQL & " AND ((APercentuale=0) OR (APercentuale IS NULL))"

sSQL = "SELECT RV_POCommissioniPerDoc.IDRV_POCommissioniPerDoc, RV_POCommissioniPerDoc.IDOggetto, RV_POCommissioniPerDoc.IDRV_POTipoCommissione, RV_POCommissioniPerDoc.Percentuale, "
sSQL = sSQL & "RV_POCommissioniPerDoc.Importo, RV_POCommissioniPerDoc.ImportoRiga, RV_POCommissioniPerDoc.Quantita, RV_POCommissioniPerDoc.APercentuale, RV_POCommissioniPerDoc.ImportoTotale,"
sSQL = sSQL & "RV_POCommissioniPerDoc.IDRV_POTipoPedana , RV_POCommissioniPerDoc.PercentualeDaCommissione, RV_POCommissioniPerDoc.IDArticoloImballo, RV_POTipoCommissione.IDRV_POTipoValoreDocumento "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc INNER JOIN "
sSQL = sSQL & "RV_POTipoCommissione ON RV_POCommissioniPerDoc.IDRV_POTipoCommissione = RV_POTipoCommissione.IDRV_POTipoCommissione "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND ((IDRV_POTipoPedana=0) OR (IDRV_POTipoPedana IS NULL))"
sSQL = sSQL & " AND ((APercentuale=0) OR (APercentuale IS NULL))"

Set rsComm = CnDMT.OpenResultset(sSQL)

While Not rsComm.EOF
    If GET_CONTROLLO_TIPO_COMM_ND_NC(fnNotNullN(rsComm!IDRV_POTipoCommissione), IDTipoOggettoND_NC) = 0 Then
        If (fnNotNullN(rsComm!IDRV_POTipoValoreDocumento) < 5) Then
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazione / 100) * fnNotNullN(rsComm!Percentuale))
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rsComm!Importo))
        Else
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazioneScontato / 100) * fnNotNullN(rsComm!Percentuale))
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rsComm!Importo))
        End If
    End If
rsComm.MoveNext
Wend

rsComm.CloseResultset
Set rsComm = Nothing

If (IDTipoOggettoND_NC = 11) Or (IDTipoOggettoND_NC = 107) Then
    rsTmp!RV_POImportoRigaCommissioni = GET_COMMISSIONI_DOCUMENTO
End If

GET_COMMISSIONI_DOCUMENTO = PrezzoLiquidazione - GET_COMMISSIONI_DOCUMENTO

End Function
Private Function GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo As Double, IDOggetto As Long, IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, CodiceLottoVendita As String, rsTmp As ADODB.Recordset) As Double
On Error GoTo ERR_GET_VARIAZIONE_LIQ_DOC_ORI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoIncidenzaImballo As Double

sSQL = "SELECT IDMovimento, RV_POImportoInclusoImballo "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND RV_POCodiceLottoVendita=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
End If
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    ImportoIncidenzaImballo = fnNotNullN(rs!RV_POImportoInclusoImballo)
    If fnNotNullN(rs!RV_POImportoInclusoImballo) > 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoInclusoImballo)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoInclusoImballo))
    End If
End If

rs.CloseResultset
Set rs = Nothing

rsTmp!RV_POImportoDaLiq = ImportoIncidenzaImballo

GET_VARIAZIONE_LIQ_DOC_ORI = Prezzo
Exit Function

ERR_GET_VARIAZIONE_LIQ_DOC_ORI:
    GET_VARIAZIONE_LIQ_DOC_ORI = Prezzo
End Function
Private Function GET_TRATT_PRE_LIQ_ART(Prezzo As Double, IDArticolo As Long) As Double
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT IDRV_POTrattenutaPerLiquidazione, ValoreTrattenutaPreLiq1, ValoreTrattenutaPreLiq2 "
sSQL = sSQL & "FROM RV_POTrattenutaPerLiquidazione "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDSocio=0"
sSQL = sSQL & " AND IDCategoriaMerceologica=0"
sSQL = sSQL & " AND IDTipoLavorazione=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TRATT_PRE_LIQ_ART = Prezzo
Else
    Prezzo = Prezzo - fnNotNullN(rs!ValoreTrattenutaPreLiq1)
    Prezzo = Prezzo - fnNotNullN(rs!ValoreTrattenutaPreLiq2)
    GET_TRATT_PRE_LIQ_ART = Prezzo
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(PrezzoLiquidazione As Double, IDOggetto As Long, IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, QuantitaLiquidazione As Double) As Double
On Error GoTo ERR_AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPed As DmtOleDbLib.adoResultset
Dim rsRiga As ADODB.Recordset
'Dim Moltiplicatore As Double
Dim Prezzo As Double
Dim ImportoCommPerPedana As Double
'Dim rsNew As ADODB.Recordset
Dim ImportoCommArticoloPerPedana As Double

AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA = PrezzoLiquidazione

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND APercentuale=1"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POCommissioniPerDocRighe "
    sSQL = sSQL & " WHERE IDOggetto = " & IDOggetto
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
    sSQL = sSQL & " AND IDRV_POTipoCommissione=" & fnNotNullN(rs!IDRV_POTipoCommissione)
    sSQL = sSQL & " AND IDRV_POCommissioniPerDoc=" & fnNotNullN(rs!IDRV_POCommissioniPerDoc)
    Set rsPed = CnDMT.OpenResultset(sSQL)
    
    While Not rsPed.EOF
        If RICALCOLA_TUTTO = 1 Then
            AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA - ((fnNotNullN(rsPed!Importo) * fnNotNullN(rsPed!Quantita)) / QuantitaLiquidazione)
        Else
            AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA = AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA - fnNotNullN(rsPed!Importo)
        End If
    rsPed.MoveNext
    Wend
    
    rsPed.CloseResultset
    Set rsPed = Nothing

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA:
    MsgBox Err.Description, vbCritical, "AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA"
End Function

Private Function GET_CONTROLLO_TIPO_COMM_ND_NC(IDTipoCommissione As Long, IDTipoOggetto As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_TIPO_COMM_ND_NC
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_TIPO_COMM_ND_NC = 0
If IDTipoOggetto = 2 Then Exit Function
If IDTipoOggetto = 8 Then Exit Function
If IDTipoOggetto = 114 Then Exit Function

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_TIPO_COMM_ND_NC = fnNotNullN(rs!NonCalcolareCommissioniInNDeNC)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_TIPO_COMM_ND_NC:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_TIPO_COMM_ND_NC"
End Function
Private Function GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = 0

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
        
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_TOTALE_COSTO_KIT(IDLavorazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(CostoTotaleRiga) as TotaleCosto "
sSQL = sSQL & "FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_COSTO_KIT = 0
Else
    GET_TOTALE_COSTO_KIT = fnNotNullN(rs!TotaleCosto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
