VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVisDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONA DOCUMENTI DA LIQUIDARE"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14100
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVisDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeselTutto 
      Caption         =   "DESELEZIONA TUTTO"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   8160
      Width           =   2415
   End
   Begin VB.CommandButton cmdSelTutto 
      Caption         =   "SELEZIONA TUTTO"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   8160
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   12360
      TabIndex        =   2
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Frame FraFiltri 
      Caption         =   "FILTRI DI RICERCA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   13935
      Begin DMTDataCmb.DMTCombo cboSINO 
         Height          =   315
         Left            =   12720
         TabIndex        =   6
         Top             =   465
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DmtSearchAccount2.DmtSearchACS2 ACS 
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1058
         WidthDescription=   4000
         WidthSecondDescription=   1500
         Object.Visible         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideLeaf        =   0   'False
         BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionDescription=   "Cliente"
         CaptionCode     =   "Codice"
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboSezionale 
         Height          =   315
         Left            =   9840
         TabIndex        =   10
         Top             =   465
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboDestinazione 
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Top             =   465
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Destinazione diversa"
         Height          =   255
         Index           =   2
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Vis. Doc. liq."
         Height          =   255
         Index           =   0
         Left            =   12720
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11456
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
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   13935
   End
End
Attribute VB_Name = "frmVisDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub ACS_ChangedElement()
    With Me.cboDestinazione
        Set .Database = CnDMT
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica "
        .Sql = .Sql & "WHERE IDAnagrafica=" & Me.ACS.IDAnagrafica
        .Fill
    End With



    fncGriglia
End Sub

Private Sub cboDestinazione_Click()
fncGriglia
End Sub

Private Sub cboSezionale_Click()
    fncGriglia
End Sub

Private Sub cboSINO_Click()
    fncGriglia
End Sub

Private Sub cmdConferma_Click()
    CONFERMA_SEL_DOCUMENTI = 1
    Unload Me
End Sub

Private Sub cmdDeselTutto_Click()
Dim sSQL As String

sSQL = "UPDATE RV_POTMPLiqFattureSel SET "
sSQL = sSQL & "Registra=0 "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

CnDMT.Execute sSQL

fncGriglia

End Sub

Private Sub cmdSelTutto_Click()
Dim sSQL As String

sSQL = "UPDATE RV_POTMPLiqFattureSel SET "
sSQL = sSQL & "Registra=1 "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

If Me.ACS.IDAnagrafica > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & Me.ACS.IDAnagrafica
End If

If Me.cboSezionale.CurrentID > 0 Then
    sSQL = sSQL & " AND IDSezionale=" & Me.cboSezionale.CurrentID
End If

If Me.cboDestinazione.CurrentID > 0 Then
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & Me.cboDestinazione.CurrentID
End If

Select Case Me.cboSINO.CurrentID
    Case 1
        sSQL = sSQL & " AND Liquidato=1"
    Case 2
        sSQL = sSQL & " AND Liquidato=0"
End Select


CnDMT.Execute sSQL

fncGriglia

End Sub

Private Sub Form_Activate()
Dim sSQL As String

Me.lblInfo.Caption = "START........."
DoEvents
    
sSQL = "DELETE FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
CnDMT.Execute sSQL

fncGriglia

ELABORAZIONE_DOCUMENTI "ValoriOggettoPerTipo0002", "Documento di trasporto"
ELABORAZIONE_DOCUMENTI "ValoriOggettoPerTipo0072", "Fattura accompnatoria"
ELABORAZIONE_DOCUMENTI "ValoriOggettoPerTipo0008", "Corrispettivi"

Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"
Me.ProgressBar1.Value = 0
DoEvents

Me.cboSINO.WriteOn 2

fncGriglia


End Sub

Private Sub ELABORAZIONE_DOCUMENTI(NomeTabella As String, Oggetto As String)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsTmp As ADODB.Recordset
Dim Unita_Progresso As Double
Dim NumeroRecord As Long

'''''''''''''''''''''''''''''''CONTEGGIO RECORD''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(Oggetto.IDOggetto) AS NumeroRecord "
sSQL = sSQL & "FROM " & NomeTabella & " INNER JOIN "
sSQL = sSQL & "Oggetto ON " & NomeTabella & ".IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "AND " & NomeTabella & ".IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND " & NomeTabella & ".Doc_Data>=" & fnNormDate(FrmNuovoPeriodo.txtDataInizio.Text)
sSQL = sSQL & " AND " & NomeTabella & ".Doc_Data<=" & fnNormDate(FrmNuovoPeriodo.txtDataFine.Text)

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumeroRecord = 0 Then Exit Sub
Me.ProgressBar1.Max = 100
Me.ProgressBar1.Value = 0
Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)


Me.lblInfo.Caption = "ELABORAZIONE " & Oggetto
DoEvents

sSQL = "SELECT " & NomeTabella & ".* "
sSQL = sSQL & "FROM " & NomeTabella & " INNER JOIN "
sSQL = sSQL & "Oggetto ON " & NomeTabella & ".IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "AND " & NomeTabella & ".IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND " & NomeTabella & ".Doc_Data>=" & fnNormDate(FrmNuovoPeriodo.txtDataInizio.Text)
sSQL = sSQL & " AND " & NomeTabella & ".Doc_Data<=" & fnNormDate(FrmNuovoPeriodo.txtDataFine.Text)

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection


sSQL = "SELECT * FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

Set rsTmp = New ADODB.Recordset
rsTmp.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsTmp.AddNew
        rsTmp!IDRV_POTMPLiqFattureSel = fnGetNewKey("RV_POTMPLiqFattureSel", "IDRV_POTMPLiqFattureSel")
        rsTmp!IDUtente = TheApp.IDUser
        rsTmp!IDOggetto = fnNotNullN(rs!IDOggetto)
        rsTmp!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
        rsTmp!Oggetto = Oggetto
        rsTmp!Registra = 0
        rsTmp!DataDocumento = fnNotNull(rs!doc_data)
        rsTmp!NumeroDocumento = fnNotNull(rs!doc_numero)
        rsTmp!IDAnagrafica = fnNotNullN(rs!Link_nom_anagrafica)
        rsTmp!Anagrafica = fnNotNull(rs!Nom_ragione_sociale_o_cognome)
        rsTmp!Codice = fnNotNull(rs!Nom_codice)
        rsTmp!Nome = fnNotNull(rs!Nom_nome)
        rsTmp!IDSezionale = fnNotNullN(rs!Link_Doc_sezionale)
        rsTmp!Sezionale = GET_SEZIONALE(rsTmp!IDSezionale)
        rsTmp!IDSitoPerAnagrafica = fnNotNullN(rs!Link_Nom_ult_sito)
        rsTmp!SitoPerAnagrafica = GET_DESTINAZIONE(rsTmp!IDSitoPerAnagrafica)
        
        If fnNotNullN(rs!RV_POIDLiquidazionePeriodo) = 0 Then
            rsTmp!Liquidato = 0
            rsTmp!IDPeriodoLiquidazione = 0
            rsTmp!NumeroLiquidazione = 0
        Else
            rsTmp!Liquidato = 1
            rsTmp!IDPeriodoLiquidazione = fnNotNullN(rs!RV_POIDLiquidazionePeriodo)
            rsTmp!NumeroLiquidazione = GET_NUMERO_LIQUIDAZIONE(fnNotNullN(rs!RV_POIDLiquidazionePeriodo))
        End If
    
    rsTmp.Update
    
    If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
    End If
    DoEvents
    
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
rsTmp.Close
Set rsTmp = Nothing
End Sub
Private Function GET_NUMERO_LIQUIDAZIONE(IDPeriodoLiquidazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NumeroLiquidazione "
sSQL = sSQL & " FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_LIQUIDAZIONE = 0
Else
    GET_NUMERO_LIQUIDAZIONE = fnNotNullN(rs!NumeroLiquidazione)
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Sub fncGriglia()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
sSQL = "SELECT * FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
If Me.ACS.IDAnagrafica > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & Me.ACS.IDAnagrafica
End If
If Me.cboSezionale.CurrentID > 0 Then
    sSQL = sSQL & " AND IDSezionale=" & Me.cboSezionale.CurrentID
End If
If Me.cboDestinazione.CurrentID > 0 Then
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & Me.cboDestinazione.CurrentID
End If
Select Case Me.cboSINO.CurrentID
    Case 1
        sSQL = sSQL & " AND Liquidato=1"
    Case 2
        sSQL = sSQL & " AND Liquidato=0"
End Select

sSQL = sSQL & " ORDER BY DataDocumento DESC, NumeroDocumento DESC "

OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = 3

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient
rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            
With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    
        .ColumnsHeader.Add "IDRV_POTMPLiqFattureSel", "IDRV_POTMPLiqFattureSel", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
        Set cl = .ColumnsHeader.Add("Registra", "Sel.", dgBoolean, True, 1000, dgAligncenter)
            cl.Editable = True
        .ColumnsHeader.Add "Liquidato", "Liquidato", dgBoolean, True, 1000, dgAligncenter
        .ColumnsHeader.Add "Oggetto", "Documento", dgVarChar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "IDSezionale", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Sezionale", "Sezionale", dgVarChar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "DataDocumento", "data documento", dgDate, True, 1500, dgAlignleft
        .ColumnsHeader.Add "NumeroDocumento", "N° documento", dgVarChar, True, 1500, dgAlignRight
        .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Codice", "Codice", dgVarChar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "Anagrafica", "Anagrafica", dgVarChar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "Nome", "Nome", dgVarChar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione diversa", dgVarChar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "IDPeriodoLiquidazione", "IDPeriodoLiquidazione", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "NumeroLiquidazione", "N° Liq.", dgInteger, True, 1500, dgAlignRight
        
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
    
CnDMT.CursorLocation = OLDCursor
End Sub

Private Sub INIT_CONTROLLI()
'Cliente
With Me.ACS
    'Imposta la connessione attiva al controllo
    Set .Connection = CnDMT
    'Imposta il nome dell'applicazione
    'Set .ApplicationName = TheApp
    'Imposta il nome dell'eseguibile dell'applicazione
    .Client = App.EXEName
    'Imposta l'identificativo dell'azienda corrente
    .IDFirm = TheApp.IDFirm
    'Imposta l'identificativo dell'utente corrente
    .IDUser = TheApp.IDUser
    .UserName = TheApp.IDUser
    'Impostare con la proprietà Hwnd del form che contiene
    'il controllo. Serve per l'esegui gestione
    .HwndContainer = Me.hWnd
End With
    
With Me.cboSINO
    Set .Database = CnDMT
    .AddFieldKey "IDRV_POSINO"
    .DisplayField = "SiNo"
    .Sql = "SELECT * FROM RV_POSINO"
    .Fill
End With
    
    
With Me.cboSezionale
    Set .Database = CnDMT
    .AddFieldKey "IDSezionale"
    .DisplayField = "Sezionale"
    .Sql = "SELECT * FROM Sezionale "
    .Sql = .Sql & "WHERE IDFiliale=" & TheApp.Branch
    .Fill
End With
End Sub

Private Sub Form_Load()
    INIT_CONTROLLI
End Sub
Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Registra").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
            rsGriglia.Fields("Registra").Value = Abs(CLng(Selected))
            'sbCheckSelected
            rsGriglia.UpdateBatch
            Me.GrigliaCorpo.Refresh
        End If
End Sub
Private Function GET_SEZIONALE(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sezionale FROM Sezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_SEZIONALE = ""
Else
    GET_SEZIONALE = fnNotNull(rs!Sezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESTINAZIONE(IDDestinazione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SitoPerAnagrafica FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDDestinazione


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESTINAZIONE = ""
Else
    GET_DESTINAZIONE = fnNotNull(rs!SitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
