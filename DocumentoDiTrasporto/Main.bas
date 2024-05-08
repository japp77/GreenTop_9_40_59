Attribute VB_Name = "ModMain"
Option Explicit

'**+
'Nome                   : Main
'Parametri              : Nessuno
'Valori di ritorno      : Nessuno
'Funzionalità           : In questa Sub vengono eseguite
'                       : le operazioni di Startup.
'                       : Questa procedura deve essere
'                       : sempre presente nel prog. gestore.
'**/
Sub Main()
    On Error GoTo ErrorHandler
    'L'Applicazione
    Set TheApp = New Application

    'Carica il form della applicazione senza mostrarlo
    Load frmMain
    
    'Abilita il form della applicazione alla ricezione
    'degli eventi da parte della applicazione Diamante
    Set frmMain.Application = TheApp

    'Il nome della applicazione
    TheApp.Name = App.EXEName
    
    
    'Viene inizializzata la componente DmtRegistry2 che si occupa della gestione
    'degli accessi al registry.
    DmtRegistry2.EXEName = App.EXEName
    
    
    Set gResource = New Resource
    REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
    
    'L'icona di Diamante
    'frmMain.Icon = gResource.GetIcon(IDI_DIAMANTE16)
        
    'Esegue l'applicazione
    TheApp.Run frmMain.hwnd
    
    
    'La vista alla partenza deve essere quella del Form
    frmMain.BrwMain.Visible = False
    '
    'alla fine del caricamento dell'applicazione
    'visualizza il layout del form.
    frmMain.Show
    
    
    '..............................................................................
    'NOTA: In Form_Activate() viene determinata la modalità iniziale con cui si   .
    'avvia il programma:                                                          .
    'Se il filtro predefinito restituisce dei record si va in modalità variazione .
    'sul primo record altrimenti si va in modalità inserimento.                   .
    '..............................................................................
    
    Exit Sub
ErrorHandler:
    Unload frmMain
    If Err.Number = 1 + vbObjectError Then
        'Questo programma può essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avrà TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        Err.Raise Err.Number
    End If
End Sub

Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    Dim VarData As String

    'Monta la query SQL per trovare il massimo valore della chiave primaria
    sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & Tabella ' & " WHERE " & >=" & VarData
    
    'Apertura del recordset
    Set rs = Cn.OpenResultset(fnAnsi2Jet(sSQL))
    
    'Determina il primo progressivo disponibile
    fnGetNewKey = fnNotNullN(rs.adoColumns("MaxID")) + 1
    If fnGetNewKey <= 0 Then fnGetNewKey = 1

    'Chiude il recordset e distrugge l'oggetto.
    rs.CloseResultset
    Set rs = Nothing
    
End Function

