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
    TheApp.Run FrmMain.hwnd
    
    
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
        
        FrmMain.InitControlli
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
On Error GoTo ERR_PrelevaAzienda
    Dim TmpFiliale As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    'TmpFiliale = GetSetting("Diamante", "MenuSettings", "LASTBRANCH")
    
    sSQL = "SELECT Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica) INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda) ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE (((Filiale.IDFiliale)=" & MenuOptions.LastBranch & "))"
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
        VarIDAzienda = rs!IDAzienda
        VarIDAttivitaAzienda = rs!IDAttivitaAzienda
        VarIDFiliale = rs!IDFiliale
        VarIDUtente = MenuOptions.LastUserID
        
    rs.CloseResultset
    Set rs = Nothing
    
Exit Sub
ERR_PrelevaAzienda:
    MsgBox Err.Description, vbCritical, "PrelevaAzienda"
End Sub
Public Function fncEsercizio(DataDocumento As String) As Long

    fncEsercizio = 0
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDEsercizio, Esercizio "
    sSQL = sSQL & " FROM Esercizio "
    sSQL = sSQL & " WHERE (IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataDocumento)
    sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataDocumento)
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fncEsercizio = fnNotNullN(rs!IDEsercizio)
    Else
        fncEsercizio = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
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

Public Function fnGetPeriodoIVA(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDPeriodoIVA FROM PeriodoIVA WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (Anno = " & DatePart("yyyy", dData) & "))"
    
    Set rsEse = CnDMT.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetPeriodoIVA = rsEse!IDPeriodoIva
    Else
        fnGetPeriodoIVA = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Function CONTROLLO_ORDINE_BLOCCATO(IDOggettoOrdine As Long, IDUtente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto, RV_POIDUtenteBlocco "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    CONTROLLO_ORDINE_BLOCCATO = True
Else
    If fnNotNullN(rs!RV_POIDUtenteBlocco) = 0 Then
        CONTROLLO_ORDINE_BLOCCATO = 0
    Else
        If fnNotNullN(rs!RV_POIDUtenteBlocco) = IDUtente Then
            CONTROLLO_ORDINE_BLOCCATO = 0
        Else
            CONTROLLO_ORDINE_BLOCCATO = fnNotNullN(rs!RV_POIDUtenteBlocco)
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Sub BLOCCA_ORDINE(IDOggettoOrdine As Long, IDUtente As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
sSQL = sSQL & "RV_POIDUtenteBlocco=" & IDUtente
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine

CnDMT.Execute sSQL

End Sub
Sub TimerProc()
    FrmMain.TimerProc1
End Sub
