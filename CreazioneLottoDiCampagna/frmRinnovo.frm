VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRinnovo 
   Caption         =   "RINNOVO LOTTI DI PRODUZIONE"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRinnovo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   16725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA OPERAZIONE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   7560
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Chiudi lotto di produzione da rinnovare"
      Height          =   195
      Left            =   6360
      TabIndex        =   4
      Top             =   7200
      Width           =   4455
   End
   Begin VB.CheckBox chkRiportaDescrizioneLotto 
      Caption         =   "Riporta la descrizione del lotto di produzione"
      Height          =   195
      Left            =   6360
      TabIndex        =   3
      Top             =   6840
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8400
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   11668
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      ColumnsHeaderHeight=   20
   End
   Begin VB.Label lblInfoCodice 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8700
      Width           =   16335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   8160
      Width           =   16335
   End
End
Attribute VB_Name = "frmRinnovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConferma_Click()
Dim UnitaProgresso As Double
Dim NumeroElaborazioni As Long
Dim rsNew As ADODB.Recordset
Dim sSQL As String
Dim I As Integer
Dim NumeroLotto As Long

If NUMERO_RINNOVO = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NUMERO_RINNOVO), 4)

NumeroElaborazioni = 1
NumeroLotto = GET_NUMERO_LOTTO

sSQL = "SELECT * FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rsRinnovo.Filter = "Rinnova=1"

While Not rsRinnovo.EOF
    Me.lblInfo.Caption = "Elaborazione " & NumeroElaborazioni & " di " & NUMERO_RINNOVO
    Me.lblInfoCodice.Caption = fnNotNull(rsRinnovo!CodiceLotto)
    DoEvents
    
    If (rsRinnovo!Rinnova = 1) Then
        rsNew.AddNew
            For I = 0 To rsNew.Fields.Count - 1
                Select Case rsNew.Fields(I).Name
                    Case "IDRV_PO01_LottoCampagna"
                        rsNew.Fields(I).Value = fnGetNewKey("RV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna")
                    Case "CodiceLotto"
                        rsNew.Fields(I).Value = GET_NUMERO_LOTTO_FORMATTATO(NumeroLotto)
                    Case "DescrizioneLotto"
                        If Me.chkRiportaDescrizioneLotto.Value = vbChecked Then
                            rsNew.Fields(I).Value = rsRinnovo.Fields(rsNew.Fields(I).Name).Value
                        Else
                            rsNew.Fields(I).Value = GET_NUMERO_LOTTO_FORMATTATO(NumeroLotto)
                        End If
                    'INSERIMENTO
                    Case "IDUtenteInserimento"
                        rsNew.Fields(I).Value = TheApp.IDUser
                    Case "PCInserimento"
                        rsNew.Fields(I).Value = GET_NOMECOMPUTER
                    Case "UtentePCInserimento"
                        rsNew.Fields(I).Value = GET_NOMEUTENTE
                    Case "DataInserimento"
                        rsNew.Fields(I).Value = Date
                    Case "OraInserimento"
                        rsNew.Fields(I).Value = GET_ORARIO(Now)
                    'MODIFICA
                    Case "IDUtenteUltimaModifica"
                        rsNew.Fields(I).Value = TheApp.IDUser
                    Case "PCUltimaModifica"
                        rsNew.Fields(I).Value = GET_NOMECOMPUTER
                    Case "UtentePCUltimaModifica"
                        rsNew.Fields(I).Value = GET_NOMEUTENTE
                    Case "DataUltimaModifica"
                        rsNew.Fields(I).Value = Date
                    Case "OraUltimaModifica"
                        rsNew.Fields(I).Value = GET_ORARIO(Now)
                    Case Else
                        rsNew.Fields(I).Value = rsRinnovo.Fields(rsNew.Fields(I).Name).Value
                End Select
            Next I
        rsNew.Update
        
        If RIPORTA_PRODOTTI = 1 Then
            RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO fnNotNullN(rsRinnovo!IDRV_PO01_LottoCampagna), fnNotNullN(rsNew!IDRV_PO01_LottoCampagna)
        End If
        
        If RIPORTA_SERRE_APPEZZAMENTI = 1 Then
            RIPORTA_SERRE_DA_SALVA_COME_NUOVO fnNotNullN(rsRinnovo!IDRV_PO01_LottoCampagna), fnNotNullN(rsNew!IDRV_PO01_LottoCampagna)
        End If
        
        If Me.Check1.Value = vbChecked Then
            AGGIORNA_LOTTO fnNotNullN(rsRinnovo!IDRV_PO01_LottoCampagna)
        End If
        
        NumeroLotto = NumeroLotto + 1
        
        
    End If
        
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    DoEvents

NumeroElaborazioni = NumeroElaborazioni + 1
rsRinnovo.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing


SALVA_NUMERO_LOTTO NumeroLotto

Unload Me

End Sub

Private Sub Form_Load()
    GET_GRIGLIA
    
    lblInfo.Caption = "LOTTI DI PRODUZIONE DA RINNOVARE: " & NUMERO_RINNOVO
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long
Dim sSQL As String
Dim cl As DmtGridCtl.dgColumnHeader

OLD_Cursor = Cn.CursorLocation
Cn.CursorLocation = 3
        
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("Rinnova", "Rinnova", dgBoolean, True, 1300, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDRV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDFiliale", "IDFiliale", dgInteger, False, 500, dgAlignleft

            .ColumnsHeader.Add "IDSocio", "IDSocio", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Anagrafica", "Anagrafica", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Nome", "Nome lotto", dgchar, False, 1800, dgAlignleft
            
            .ColumnsHeader.Add "CodiceLotto", "Codice", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DescrizioneLotto", "Descrizione lotto", dgchar, True, 1800, dgAlignleft
            
            .ColumnsHeader.Add "DataSemina", "Data semina", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DataInizioProduzione", "Data inizio raccolta", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DataFineProduzione", "Data fine raccolta", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Chiuso", "Chiuso", dgBoolean, True, 1300, dgAligncenter
            
            .ColumnsHeader.Add "IDRV_PO01_FamigliaProdotti", "IDRV_PO01_FamigliaProdotti", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "FamigliaProdotti", "Famiglia prodotti", dgchar, True, 1800, dgAlignleft
            
            .ColumnsHeader.Add "IDRV_PO01_Varieta", "IDRV_PO01_Varieta", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Varieta", "Varietà", dgchar, True, 1800, dgAlignleft
                   
        Set .Recordset = rsRinnovo
        .Refresh
        .LoadUserSettings
    End With

Cn.CursorLocation = OLD_Cursor


Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName


End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsRinnovo.Fields("Rinnova").Value)), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsRinnovo.EOF And Not rsRinnovo.BOF Then
                
        rsRinnovo.Fields("Rinnova").Value = Abs(CLng(Selected))
        
        rsRinnovo.UpdateBatch
        Me.Griglia.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub

Private Function SALVA_NUMERO_LOTTO(NumeroLotto As Long) As String
Dim rs As ADODB.Recordset
Dim I As Integer
Dim sSQL As String

If LINK_TIPO_CONTATORE = 0 Then

    sSQL = "UPDATE RV_PO01_ParametriFiliale SET "
    sSQL = sSQL & "NumeroLottoDiCampagna=" & fnNormNumber(NumeroLotto)
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Cn.Execute sSQL
    
Else
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_PO01_ContatoreLotto "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_PO01_PeriodoCampagna=" & fnNotNullN(rsRinnovo!IDRV_PO01_PeriodoCampagna)
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!IDRV_PO01_ContatoreLotto = fnGetNewKey("RV_PO01_ContatoreLotto", "IDRV_PO01_ContatoreLotto")
        rs!IDAzienda = TheApp.IDFirm
        rs!IDFiliale = TheApp.Branch
        rs!IDRV_PO01_PeriodoCampagna = fnNotNullN(rsRinnovo!IDRV_PO01_PeriodoCampagna)
        rs!Numerazione = NumeroLotto
    Else
        rs!Numerazione = NumeroLotto
    End If
        
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
End If
End Function

Private Function GET_NUMERO_LOTTO() As Long
Dim Codice As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer

If LINK_TIPO_CONTATORE = 0 Then
    sSQL = "SELECT NumeroLottoDiCampagna "
    sSQL = sSQL & "FROM RV_PO01_ParametriFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_NUMERO_LOTTO = 1
    Else
        GET_NUMERO_LOTTO = fnNotNullN(rs!NumeroLottoDiCampagna)
    End If
'
'    GET_NUMERO_LOTTO = ""
'    For I = 6 To (Len(CStr(Codice)) + 1) Step -1
'        GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & "0"
'    Next I
'
'    GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & CStr(Codice)
    rs.CloseResultset
    Set rs = Nothing
Else
    sSQL = "SELECT Numerazione "
    sSQL = sSQL & "FROM RV_PO01_ContatoreLotto "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_PO01_PeriodoCampagna=" & fnNotNullN(rsRinnovo!IDRV_PO01_PeriodoCampagna)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_NUMERO_LOTTO = 1
    Else
        GET_NUMERO_LOTTO = fnNotNullN(rs!Numerazione)
    End If
    
'    GET_NUMERO_LOTTO = ""
'    For I = 6 To (Len(CStr(Codice)) + 1) Step -1
'        GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & "0"
'    Next I
'
'    GET_NUMERO_LOTTO = GET_NUMERO_LOTTO & CStr(Codice)

    rs.CloseResultset
    Set rs = Nothing
End If

End Function

Private Function FORMATTA_CAMPO(Valore As String, Lunghezza As Long, filler As Boolean, StringaFiller As String) As String
Dim LunghezzaCampo As Long
Dim I As Long

LunghezzaCampo = Len(Valore)
FORMATTA_CAMPO = ""

If filler = True Then
    Select Case (LunghezzaCampo - Lunghezza)
        Case Is = 0
            FORMATTA_CAMPO = Valore
        Case Is > 0
            FORMATTA_CAMPO = Mid(Valore, 1, Lunghezza)
        Case Is < 0
            FORMATTA_CAMPO = Valore
            For I = LunghezzaCampo To Lunghezza - 1
                FORMATTA_CAMPO = StringaFiller & FORMATTA_CAMPO
            Next
    End Select
Else
    If LunghezzaCampo >= Lunghezza Then
        FORMATTA_CAMPO = Mid(Valore, 1, Lunghezza)
    Else
        FORMATTA_CAMPO = Valore
    End If
End If



End Function
Private Sub RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO(IDLottoCampagna As Long, IDLottoCampagnaNew As Long)
On Error GoTo ERR_RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_PO01_DettaglioLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PO01_DettaglioLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagnaNew

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PO01_DettaglioLotto = fnGetNewKey("RV_PO01_DettaglioLotto", "IDRV_PO01_DettaglioLotto")
        rsNew!IDRV_PO01_LottoCampagna = IDLottoCampagnaNew
        rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsNew!Calibro = fnNotNull(rs!Calibro)
        rsNew!QuantitaPresunta = 0
        rsNew("IDUtenteInserimento").Value = TheApp.IDUser
        rsNew("PCInserimento").Value = GET_NOMECOMPUTER
        rsNew("UtentePCInserimento").Value = GET_NOMEUTENTE
        rsNew("DataInserimento").Value = Date
        rsNew("OraInserimento").Value = GET_ORARIO(Now)
        rsNew("IDUtenteUltimaModifica").Value = TheApp.IDUser
        rsNew("PCUltimaModifica").Value = GET_NOMECOMPUTER
        rsNew("UtentePCUltimaModifica").Value = GET_NOMEUTENTE
        rsNew("DataUltimaModifica").Value = Date
        rsNew("OraUltimaModifica").Value = GET_ORARIO(Now)
        
        
    rsNew.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO:
    MsgBox Err.Description, vbCritical, "RIPORTA_PRODOTTI_DA_SALVA_COME_NUOVO"
End Sub

Private Sub RIPORTA_SERRE_DA_SALVA_COME_NUOVO(IDLottoCampagna As Long, IDLottoCampagnaNew As Long)
On Error GoTo ERR_RIPORTA_SERRE_DA_SALVA_COME_NUOVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagnaNew

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PO01_SerraPerLotto = fnGetNewKey("RV_PO01_SerraPerLotto", "IDRV_PO01_SerraPerLotto")
        rsNew!IDRV_PO01_LottoCampagna = IDLottoCampagnaNew
        rsNew!IDRV_PO01_Serra = fnNotNullN(rs!IDRV_PO01_Serra)
        rsNew!DimensioneMq = fnNotNullN(rs!DimensioneMq)
        rsNew!DimensioneHA = fnNotNull(rs!DimensioneHA)
        
    rsNew.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_RIPORTA_SERRE_DA_SALVA_COME_NUOVO:
    MsgBox Err.Description, vbCritical, "RIPORTA_SERRE_DA_SALVA_COME_NUOVO"
End Sub
Private Function GET_NUMERO_LOTTO_FORMATTATO(Codice As Long) As String

    GET_NUMERO_LOTTO_FORMATTATO = ""
    For I = 6 To (Len(CStr(Codice)) + 1) Step -1
        GET_NUMERO_LOTTO_FORMATTATO = GET_NUMERO_LOTTO_FORMATTATO & "0"
    Next I

    GET_NUMERO_LOTTO_FORMATTATO = GET_NUMERO_LOTTO_FORMATTATO & CStr(Codice)

End Function
Private Sub AGGIORNA_LOTTO(IDLottoCampagnaOLD)
Dim sSQL As String

sSQL = "UPDATE RV_PO01_LottoCampagna SET "
sSQL = sSQL & "Chiuso=" & fnNormBoolean(1)
sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagnaOLD

Cn.Execute sSQL

End Sub
Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
    
End Function
Private Function GET_ORARIO(StringaData As String) As String
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(DatePart("h", StringaData)) = 1 Then
    Ora = "0" & DatePart("h", StringaData)
Else
    Ora = DatePart("h", StringaData)
End If
If Len(DatePart("n", StringaData)) = 1 Then
    Minuti = "0" & DatePart("n", StringaData)
Else
    Minuti = DatePart("n", StringaData)
End If
If Len(DatePart("s", StringaData)) = 1 Then
    Secondi = "0" & DatePart("s", StringaData)
Else
    Secondi = DatePart("s", StringaData)
End If

GET_ORARIO = Ora & "." & Minuti & "." & Secondi

End Function
Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaBuoniFiltro_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If Griglia.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If Griglia.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsRinnovo.Fields("Rinnova").Value), 2
            End If
        End If
    End If

Exit Sub
ERR_GrigliaBuoniFiltro_MouseUp:
    MsgBox Err.Description, vbCritical, "Griglia_MouseUp"

End Sub
