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
Public Function fnGetEsercizio(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio, DataInizio, DataFine FROM Esercizio WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (DataInizio <= " & fnNormDate(dData) & ")"
    sSQL = sSQL & " AND (DataFine >= " & fnNormDate(dData) & "))"
   

    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = rsEse!IDEsercizio
        DataInizio_Esercizio = rsEse!DataInizio
        DataFine_Esercizio = rsEse!DataFine
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
        
        sSQL = "SELECT " & CampoKey & " FROM " & Tabella & " ORDER BY " & CampoKey & " DESC"
        
        Set rs = Cn.OpenResultset(fnAnsi2Jet(sSQL))
    
        If rs.EOF = True Then
        
            fnGetNewKey = 1
    
        Else
            
            fnGetNewKey = fnNotNullN(rs.adoColumns(CampoKey)) + 1
    
        End If

        rs.CloseResultset
        Set rs = Nothing
    

    
End Function
Public Function fnGetPeriodoIVA(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDPeriodoIVA FROM PeriodoIVA WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (Anno = " & DatePart("yyyy", dData) & "))"
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetPeriodoIVA = rsEse!IDPeriodoIVA
    Else
        fnGetPeriodoIVA = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    '------------------------------
    'APERTURA DELLA CONNESSIONE
    '------------------------------
    
    'Leggiamo il tipo di database utilizzato (Access o SQL Server)
    'Apriamo la connessione in base al tipo di database rilevato
    '(MenuOptions.DBType restituisce il valore del DBType)
    'Select Case MenuOptions.DBType
    '    Case 0 'CONNESSIONE_SQL_SERVER            'Microsoft SQL Server
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=sa;PWD=")
    '    Case 1 'CONNESSIONE_ACCESS               'Microsoft ACCESS
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=admin;PWD=dmt192981046")
    '    Case -1
            'Se la voce DBType non viene trovata nel file di registro
            'vuol dire che Diamante non è stato installato correttamente
    '        MsgBox "Impossibile avviare il programma. Diamante non è stato installatto correttamente!", vbCritical, "Aggiornamento scadenze"
    '        End
    'End Select
    
    Set Cn = TheApp.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub
Public Sub fnCalcolaPrezzoUnitario(IDConferimentoRiga As Long, TipoLiquidazione As Long, DataInizio As String, DataFine As String)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsComm As DmtOleDbLib.adoResultset
Dim Prezzo As Double
Dim rsOgg As DmtOleDbLib.adoResultset
Dim Link_Oggetto As Long
Dim Unita_progresso As Double


''''''''''''''''''DOCUMENTO DI TRASPORTO''''''''''''''''''''


sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & "WHERE RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
Else
    sSQL = sSQL & "WHERE Doc_data>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND Doc_data<=" & fnNormDate(DataFine)
End If
If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POIDSocio=" & LINK_SOCIO
End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga = 1"
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0004.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then

    While Not rs.EOF
            
        
    
        sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
        Set rsComm = Cn.OpenResultset(sSQL)
        
        Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        
        While Not rsComm.EOF
            Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
        rsComm.MoveNext
        Wend
        
        rsComm.CloseResultset
        Set rsComm = Nothing
    
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
        End If
        
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''FATTURA ACCOMPAGNATORIA''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & "WHERE RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
Else
    sSQL = sSQL & "WHERE Doc_data>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND Doc_data<=" & fnNormDate(DataFine)
End If
If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POIDSocio=" & LINK_SOCIO
End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga = 1"

sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0001.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    While Not rs.EOF
            
        
    
        sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
        Set rsComm = Cn.OpenResultset(sSQL)
        
        Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        
        While Not rsComm.EOF
            Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
        rsComm.MoveNext
        Wend
        
        rsComm.CloseResultset
        Set rsComm = Nothing
    
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
        End If
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

''''''''''''''''''SCONTRINO NON FISCALE''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & "WHERE RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
Else
    sSQL = sSQL & "WHERE Doc_data>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND Doc_data<=" & fnNormDate(DataFine)
End If
If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POIDSocio=" & LINK_SOCIO
End If
sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga = 1"

sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0034.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    While Not rs.EOF
            
        
    
        sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
        Set rsComm = Cn.OpenResultset(sSQL)
        
        Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        
        While Not rsComm.EOF
            Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
        rsComm.MoveNext
        Wend
        
        rsComm.CloseResultset
        Set rsComm = Nothing
    
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
        End If
        
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

''''''''''''''''''NOTA DI CREDITO''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0016.*, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & "WHERE RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
Else
    sSQL = sSQL & "WHERE Doc_data>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND Doc_data<=" & fnNormDate(DataFine)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0016.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then
    
    While Not rs.EOF
            
        
    
        sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
        Set rsComm = Cn.OpenResultset(sSQL)
        
        Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        
        While Not rsComm.EOF
            Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
        rsComm.MoveNext
        Wend
        
        rsComm.CloseResultset
        Set rsComm = Nothing
    
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
        End If
        
    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing

''''''''''''''''''NOTA DI DEBITO''''''''''''''''''''
sSQL = "SELECT ValoriOggettoDettaglio0007.*, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto "

If TipoLiquidazione = 1 Then
    sSQL = sSQL & "WHERE RV_PODataConferimento>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_PODataConferimento<=" & fnNormDate(DataFine)
Else
    sSQL = sSQL & "WHERE Doc_data>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND Doc_data<=" & fnNormDate(DataFine)
End If
sSQL = sSQL & " ORDER BY ValoriOggettoDettaglio0007.IDOggetto"


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = False Then

    While Not rs.EOF
            
        
    
        sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
        Set rsComm = Cn.OpenResultset(sSQL)
        
        Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        
        While Not rsComm.EOF
            Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rsComm!Percentuale))
        rsComm.MoveNext
        Wend
        
        rsComm.CloseResultset
        Set rsComm = Nothing
    
        rs!RV_POImportoLiq = Prezzo
        rs.Update
        
        If Link_Oggetto <> fnNotNullN(rs!IDOggetto) Then
            Link_Oggetto = fnNotNullN(rs!IDOggetto)
        End If
        

    rs.MoveNext
    
    Wend

End If
rs.Close
Set rs = Nothing
End Sub
Private Function fnControlloEsistenzaDocTMP(IDOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

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

