VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmSelDocStampa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONA DOCUMENTI PER STAMPA"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelDocStampa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   16440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ESPORTA IN PDF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   11
      Top             =   10320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STAMPA ANTEPRIMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   10320
      Width           =   2895
   End
   Begin VB.CommandButton cmdDeSelTutto 
      Caption         =   "Deseleziona tutto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelTutto 
      Caption         =   "Seleziona tutto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opzioni di stampa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   9120
      Width           =   12015
      Begin VB.ComboBox cboStampa 
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   5175
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Non stampare come singolo report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Stampante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   7
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Stampa numero copie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "STAMPA IMMEDIATA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   10320
      Width           =   2895
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   15901
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
End
Attribute VB_Name = "frmSelDocStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oReport As dmtReportLib.dmtReport

Private Sub cmdConferma_Click()
    STAMPA Me.Check1.Value
End Sub

Private Sub cmdDeSelTutto_Click()
On Error GoTo ERR_cmdDeSelTutto_Click
If ((rsGrigliaSelDocTMP.EOF) And (rsGrigliaSelDocTMP.BOF)) Then Exit Sub

rsGrigliaSelDocTMP.MoveFirst

Griglia.UpdatePosition = False

While Not rsGrigliaSelDocTMP.EOF
    rsGrigliaSelDocTMP!Selezionato = False
    'rsGrigliaSelDocTMP.Update
rsGrigliaSelDocTMP.MoveNext
Wend

Griglia.UpdatePosition = True

Me.Griglia.Refresh

Exit Sub
ERR_cmdDeSelTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdDeSelTutto_Click"
End Sub

Private Sub cmdSelTutto_Click()
On Error GoTo ERR_cmdSelTutto_Click
If ((rsGrigliaSelDocTMP.EOF) And (rsGrigliaSelDocTMP.BOF)) Then Exit Sub

rsGrigliaSelDocTMP.MoveFirst

Griglia.UpdatePosition = False

While Not rsGrigliaSelDocTMP.EOF
    rsGrigliaSelDocTMP!Selezionato = True
    'rsGrigliaSelDocTMP.Update
rsGrigliaSelDocTMP.MoveNext
Wend

Griglia.UpdatePosition = True

Me.Griglia.Refresh

Exit Sub
ERR_cmdSelTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdSelTutto_Click"
End Sub

Private Sub Command1_Click()
    STAMPA_ANTEPRIMA
End Sub

Private Sub Command2_Click()
On Error GoTo ERR_STAMPA
Dim NomeCartella As String
Dim f As FileSystemObject
    rsGrigliaSelDocTMP.Filter = "Selezionato=" & fnNormBoolean(1)
    
    If ((rsGrigliaSelDocTMP.EOF) And (rsGrigliaSelDocTMP.BOF)) Then
        MsgBox "Non risultano documenti da stampare", vbCritical, "Validazione dati"
        Exit Sub
    End If
    
    fnDeleteTabellaRicorsione TheApp.IDUser

    Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = Cn
    oReport.Password = TheApp.Password
    oReport.User = TheApp.User
    oReport.Copies = Me.txtNumeroCopie.Value
    oReport.Orientation = ORIENTAMENTO_SEL_DOC
    oReport.BranchID = oDoc.IDFiliale 'IDFiliale
    oReport.DocTypeID = fncIDTipoOggettoPrg(App.EXEName)
    
    
    ''''RECUPERO DATI CARTELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    NomeCartella = TrovaCartella(CSIDL_COMMON_APPDATA)
    
    Set f = New FileSystemObject
    
    If (f.FolderExists(NomeCartella & "GreenTop") = False) Then
        f.CreateFolder NomeCartella & "GreenTop"
    End If
    
    If (f.FolderExists(NomeCartella & "GreenTop\ExpPDF") = True) Then
        f.DeleteFolder NomeCartella & "GreenTop\ExpPDF", True
    End If
    
    f.CreateFolder NomeCartella & "GreenTop\ExpPDF"
    
    Set f = Nothing
    
    NomeCartella = NomeCartella & "GreenTop\ExpPDF\"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    rsGrigliaSelDocTMP.MoveFirst

    While Not rsGrigliaSelDocTMP.EOF
        
        If modalita = 0 Then
            fnDeleteTabellaRicorsione TheApp.IDUser
        End If
        
        oDoc.Prepare2Print rsGrigliaSelDocTMP!IDAzienda, TheApp.IDUser, rsGrigliaSelDocTMP!IDOggetto, rsGrigliaSelDocTMP!IDTipoOggetto
        oReport.Where = "IDOggetto = " & rsGrigliaSelDocTMP!IDOggetto
        oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
        
        If modalita = 0 Then
            SendDocument recPDF, NomeCartella, 0
        End If
              
    rsGrigliaSelDocTMP.MoveNext
    Wend
    
    If modalita = 1 Then
        SendDocument recPDF, NomeCartella, 0
    End If
    
    
    Shell "explorer.exe /e, " & NomeCartella, vbNormalFocus
    
Exit Sub
ERR_STAMPA:
    MsgBox Err.Description, vbCritical, "STAMPA IMMEDIATA"

End Sub

Private Sub Form_Load()

    fncStampanti
    
    Me.txtNumeroCopie.Value = NUMERO_COPIE_SEL_DOC
    If Me.txtNumeroCopie.Value = 0 Then Me.txtNumeroCopie.Value = 1
    
    CREA_RECORDSET
    
End Sub
Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim I As Long


Screen.MousePointer = 11
DoEvents

Set rsGrigliaSelDocTMP = New ADODB.Recordset
rsGrigliaSelDocTMP.CursorLocation = adUseClient

rsGrigliaSelDocTMP.Fields.Append "Selezionato", adBoolean, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "IDTipoOggetto", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "IDAzienda", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "IDAttivitaAzienda", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "IDFiliale", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Link_Doc_sezionale", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Sezionale", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_data", adDBDate, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_numero", adInteger, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "SitoPerAnagrafica", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Nom_ragione_sociale_o_cognome", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Nom_nome", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Nom_codice", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Nom_partita_IVA", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Nom_codice_fiscale", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_data_ns_ordine_di_rifer", adDBDate, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_numero_ns_ordine_di_rifer", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_data_vs_ordine_di_rifer", adDBDate, , adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_numero_vs_ordine_di_rifer", adVarChar, 250, adFldIsNullable
rsGrigliaSelDocTMP.Fields.Append "Doc_prefisso", adVarChar, 250, adFldIsNullable

rsGrigliaSelDocTMP.Open , , adOpenKeyset, adLockBatchOptimistic

If Not ((rsGrigliaSelDoc.EOF) And (rsGrigliaSelDoc.BOF)) Then

    rsGrigliaSelDoc.MoveFirst
    
    While Not rsGrigliaSelDoc.EOF
        rsGrigliaSelDocTMP.AddNew
            For I = 0 To rsGrigliaSelDocTMP.Fields.Count - 1
                If ((rsGrigliaSelDocTMP.Fields(I).Name <> "Selezionato") And (rsGrigliaSelDocTMP.Fields(I).Name <> "Doc_prefisso")) Then
                    rsGrigliaSelDocTMP.Fields(I).Value = rsGrigliaSelDoc.Fields(rsGrigliaSelDocTMP.Fields(I).Name).Value
                End If
            Next
            rsGrigliaSelDocTMP!Selezionato = False
            rsGrigliaSelDocTMP!Doc_prefisso = GET_PREFISSO(fnNotNullN(rsGrigliaSelDocTMP!Link_Doc_sezionale))
        rsGrigliaSelDocTMP.Update
    rsGrigliaSelDoc.MoveNext
    Wend
End If

GET_GRIGLIA

Screen.MousePointer = 0

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
    Screen.MousePointer = 0
End Sub

Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto
    sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
    
    Cn.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function fncIDTipoOggettoPrg(Gestore As String) As Long
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore"
    sSQL = sSQL & " FROM Gestore INNER JOIN TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore"
    sSQL = sSQL & " WHERE (((Gestore.Gestore)=" & fnNormString(Gestore) & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub STAMPA(modalita As Long)
On Error GoTo ERR_STAMPA
    
    rsGrigliaSelDocTMP.Filter = "Selezionato=" & fnNormBoolean(1)
    
    If ((rsGrigliaSelDocTMP.EOF) And (rsGrigliaSelDocTMP.BOF)) Then
        MsgBox "Non risultano documenti da stampare", vbCritical, "Validazione dati"
        Exit Sub
    End If
    
    fnDeleteTabellaRicorsione TheApp.IDUser

    Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = Cn
    oReport.Password = TheApp.Password
    oReport.User = TheApp.User
    oReport.Copies = Me.txtNumeroCopie.Value
    oReport.Orientation = ORIENTAMENTO_SEL_DOC
    oReport.BranchID = oDoc.IDFiliale 'IDFiliale
    oReport.DocTypeID = fncIDTipoOggettoPrg(App.EXEName)
            
    rsGrigliaSelDocTMP.MoveFirst
    
    While Not rsGrigliaSelDocTMP.EOF
        
        DoEvents
        Screen.MousePointer = 11
        If modalita = 0 Then
            fnDeleteTabellaRicorsione TheApp.IDUser
        End If
        
        oDoc.Prepare2Print rsGrigliaSelDocTMP!IDAzienda, TheApp.IDUser, rsGrigliaSelDocTMP!IDOggetto, rsGrigliaSelDocTMP!IDTipoOggetto
        oReport.Where = "IDOggetto = " & rsGrigliaSelDocTMP!IDOggetto
        oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
        
        Screen.MousePointer = 0
        DoEvents
        
        If modalita = 0 Then
            If Len(Me.cboStampa.Text) > 0 Then
                oReport.DoPrint Me.cboStampa.Text
            Else
                oReport.DoPrint oReport.PrinterName
            End If
                
        End If
                
    rsGrigliaSelDocTMP.MoveNext
    Wend
    
    If modalita = 1 Then
        If Len(Me.cboStampa.Text) > 0 Then
            oReport.DoPrint Me.cboStampa.Text
        Else
            oReport.DoPrint oReport.PrinterName
        End If
    End If
    
    Unload Me

Exit Sub
ERR_STAMPA:
    MsgBox Err.Description, vbCritical, "STAMPA IMMEDIATA"
    Screen.MousePointer = 0
End Sub
Private Sub fnDeleteTabellaRicorsione(IDUtente As Long)
On Error GoTo ERR_fnDeleteTabellaRicorsione
Dim sSQL As String
    
    sSQL = "DELETE FROM TabellaRicorsione "
    sSQL = sSQL & "WHERE IDUtente=" & IDUtente
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM TabellaRicorsione2 "
    sSQL = sSQL & "WHERE IDUtente=" & IDUtente
    Cn.Execute sSQL

Exit Sub
ERR_fnDeleteTabellaRicorsione:
    MsgBox Err.Description, vbCritical, "Cancellazione tabella ricorsione"
End Sub

Private Sub fncStampanti()
Dim prn As Printer

Me.cboStampa.Clear


For Each prn In Printers
    Me.cboStampa.AddItem prn.DeviceName
Next

End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        Set cl = .ColumnsHeader.Add("Selezionato", "Sel.", dgBoolean, True, 1500, dgAligncenter)
            cl.Editable = True
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDAttivitaAzienda", "IDAttivitaAzienda", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDFiliale", "IDFiliale", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "Link_Doc_sezionale", "Link_Doc_sezionale", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, True, 2000, dgAlignleft
        .ColumnsHeader.Add "Doc_data", "Data doc.", dgDate, True, 2000, dgAlignleft
        .ColumnsHeader.Add "Doc_numero", "N° doc.", dgNumeric, True, 1500, dgAlignRight
        .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 3000, dgAlignleft
        .ColumnsHeader.Add "Nom_nome", "Nome", dgchar, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Nom_codice", "Codice", dgchar, False, 2000, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagrafica", "Altra destinazione", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "Nom_partita_IVA", "Partita I.V.A.", dgchar, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Nom_codice_fiscale", "Codice fiscale", dgchar, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Doc_data_ns_ordine_di_rifer", "Data ordine interno", dgDate, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Doc_numero_ns_ordine_di_rifer", "N° ordine interno", dgNumeric, False, 1500, dgAlignRight
        .ColumnsHeader.Add "Doc_data_vs_ordine_di_rifer", "Data ordine cliente", dgDate, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Doc_numero_vs_ordine_di_rifer", "N° ordine cliente", dgNumeric, False, 1500, dgAlignRight
        
    Set .Recordset = rsGrigliaSelDocTMP
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                sbSelectSelectedRow Not CBool(rsGrigliaSelDocTMP.Fields("Selezionato").Value), 2
            End If
        End If
    End If
    
    
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGrigliaSelDocTMP.Fields("Selezionato").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsGrigliaSelDocTMP.EOF And Not rsGrigliaSelDocTMP.BOF Then
        rsGrigliaSelDocTMP.Fields("Selezionato").Value = Abs(CLng(Selected))
        'sbCheckSelected
        Me.Griglia.Refresh
    End If
End Sub

Private Sub STAMPA_ANTEPRIMA()
On Error GoTo ERR_STAMPA
    
    rsGrigliaSelDocTMP.Filter = "Selezionato=" & fnNormBoolean(1)
    
    If ((rsGrigliaSelDocTMP.EOF) And (rsGrigliaSelDocTMP.BOF)) Then
        MsgBox "Non risultano documenti da stampare", vbCritical, "Validazione dati"
        Exit Sub
    End If
    
    fnDeleteTabellaRicorsione TheApp.IDUser

    Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = Cn
    oReport.Password = TheApp.Password
    oReport.User = TheApp.User
    oReport.Copies = Me.txtNumeroCopie.Value
    oReport.Orientation = ORIENTAMENTO_SEL_DOC
    oReport.BranchID = oDoc.IDFiliale 'IDFiliale
    oReport.DocTypeID = fncIDTipoOggettoPrg(App.EXEName)
            
    rsGrigliaSelDocTMP.MoveFirst
    
    While Not rsGrigliaSelDocTMP.EOF
        
        oDoc.Prepare2Print rsGrigliaSelDocTMP!IDAzienda, TheApp.IDUser, rsGrigliaSelDocTMP!IDOggetto, rsGrigliaSelDocTMP!IDTipoOggetto

    rsGrigliaSelDocTMP.MoveNext
    Wend
    

    oReport.Where = "IDUtente = " & TheApp.IDUser
    oReport.Preview 0, 0, 0
    
    Unload Me
    
Exit Sub
ERR_STAMPA:
    MsgBox Err.Description, vbCritical, "STAMPA ANTEPRIMA"
End Sub

'**+
'Nome: SendDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub SendDocument(ByVal Appl As Long, percorso As String, Optional InvioEmail As Long = 1)
    On Error GoTo errHandler


    Dim OLDCursor As Integer
    Dim SExt As String
    Dim DataDocumento As String
    Dim NomeCartella As String
    Dim NomeFile As String
    Dim InvioEmailPersonalizzata As Boolean
    
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
   
   
    Select Case Appl
        Case 0
            SExt = ".xls"
        Case 1
            SExt = ".doc"
        Case 2
            SExt = ".html"
        Case 3
            SExt = ".pdf"
    End Select
 
    NomeFile = GET_NOME_FILE_DOCUMENTO '& SExt
    
    oReport.ExportFileName = percorso & NomeFile
    oReport.ShowExportFile = False
    oReport.Export Appl
    
    Screen.MousePointer = OLDCursor
    
    
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    MsgBox Err.Description, vbCritical, "SendDocument"
    

End Sub
Private Function GET_INVIO_EMAIL_PERSONALIZZATA(IDUtente As Long) As Boolean
On Error GoTo ERR_GET_INVIO_EMAIL_PERSONALIZZATA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUtente, eMail_ServerSMTP "
sSQL = sSQL & "FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND eMail_ProtocolloInvio=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_INVIO_EMAIL_PERSONALIZZATA = False
Else
    If (Len(fnNotNull(rs!eMail_ServerSMTP)) > 0) Then
        GET_INVIO_EMAIL_PERSONALIZZATA = True
    Else
        GET_INVIO_EMAIL_PERSONALIZZATA = False
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_INVIO_EMAIL_PERSONALIZZATA:
    MsgBox Err.Description, vbCritical, "GET_INVIO_EMAIL_PERSONALIZZATA"
End Function
Private Function GET_INDIRIZZO_EMAIL_CLIENTE(IDAnagrafica As Long) As String
On Error GoTo ERR_GET_INDIRIZZO_EMAIL_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica, EmailInternet "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_INDIRIZZO_EMAIL_CLIENTE = ""
Else
    GET_INDIRIZZO_EMAIL_CLIENTE = fnNotNull(rs!EmailInternet)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_INDIRIZZO_EMAIL_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_INDIRIZZO_EMAIL_CLIENTE"
End Function

Private Function PREPARAZIONE_EMAIL(NomeFile As String, IDAnagrafica As Long, IDDestinazione As Long, NumeroDocumento As String, DataDocumento As String, DescrizioneOggetto As String) As Boolean
On Error GoTo ERR_PREPARAZIONE_EMAIL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

PREPARAZIONE_EMAIL = False
'On Error Resume Next
'''''''VARIABILI PER OUTLOOK''''''''''''''''''''
Dim Out 'As Outlook.Application
Dim Name 'As Outlook.NameSpace
Dim f 'As Outlook.MAPIFolder 'MAIN MAPIFOLDER
Dim G 'As Outlook.MAPIFolder 'CONTATCT MAPIFOLDER
Dim H 'As Outlook.MAPIFolder 'CARTELLA POINT MAPIFOLDER
Dim Rubrica ' As Outlook.AddressList
Dim Cont 'As Outlook.ContactItem
Dim RCP 'As Outlook.Recipient
Dim Lista 'As Outlook.DistListItem
Dim olMail 'As Outlook.MailItem
Dim oAttach 'As Outlook.Attachment
''''''VARIABILI GENERALI''''''''''''''''''''''''''''
Dim percorso As String
Dim StringaBody As String
Dim ret

    Set Out = CreateObject("Outlook.Application")
    
    Set Name = Out.GetNamespace("MAPI")
    Name.Logon
    
    'Set olMail = Out.CreateItem(olMailItem)
    Set olMail = Out.CreateItem(0)
    olMail.To = GET_EMAIL_ANAGRAFICA(IDAnagrafica)
    If IDDestinazione > 0 Then
        olMail.cc = GET_EMAIL_ANAGRAFICA_DEST(IDDestinazione)
    End If
    olMail.Subject = DescrizioneOggetto & " n° " & NumeroDocumento & " del " & DataDocumento
    
    ''''RECUPERO DAI DEL CORPO DEL MESSAGGIO
    StringaBody = DescrizioneOggetto & " n° " & NumeroDocumento & " del " & DataDocumento
    olMail.Body = StringaBody
    
    olMail.Attachments.Add NomeFile
    
    olMail.Display
    
    Name.Logoff
    Set Name = Nothing
    Set olMail = Nothing
    Set Out = Nothing
    PREPARAZIONE_EMAIL = True
    
    'ret = Shell("explorer """ & NomeFile & """,/select", vbNormalNoFocus)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function
ERR_PREPARAZIONE_EMAIL:
    MsgBox Err.Description, vbCritical, "Preparazione e-mail"
End Function

Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
End Function
Private Function GET_EMAIL_ANAGRAFICA(IDAnagrafica As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT EmailInternet FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_EMAIL_ANAGRAFICA = ""
Else
    GET_EMAIL_ANAGRAFICA = fnNotNull(rs!EmailInternet)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_EMAIL_ANAGRAFICA_DEST(IDSitoPerAnagrafica As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Email FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_EMAIL_ANAGRAFICA_DEST = ""
Else
    GET_EMAIL_ANAGRAFICA_DEST = fnNotNull(rs!Email)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_NOME_FILE_DOCUMENTO() As String

Dim NomeFile As String

GET_NOME_FILE_DOCUMENTO = ""

'GET_NOME_FILE_DOCUMENTO = oDoc.Descrizione & ""
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " " & fnNotNull(oDoc.Field("Nom_ragione_sociale_o_cognome", , sTabellaTestata)) & fnNotNull(oDoc.Field("Nom_nome", , sTabellaTestata))
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " (" & fnNotNull(oDoc.Field("Nom_codice", , sTabellaTestata)) & ")"
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " [" & GET_DATA_FORMATTATA(oDoc.DataEmissione) & "]"
    
GET_NOME_FILE_DOCUMENTO = oDoc.Descrizione & ""
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " " & fnNotNull(rsGrigliaSelDocTMP!Nom_ragione_sociale_o_cognome) & fnNotNull(rsGrigliaSelDocTMP!Nom_nome)
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " (" & fnNotNull(rsGrigliaSelDocTMP!Nom_codice) & ")"
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " [" & GET_DATA_FORMATTATA(rsGrigliaSelDocTMP!Doc_data) & "]"


End Function
Private Function GET_DATA_FORMATTATA(DataF As String) As String
On Error GoTo ERR_GET_DATA_FORMATTATA
Dim Anno As String
Dim mese As String
Dim giorno As String

GET_DATA_FORMATTATA = ""

Anno = Year(DataF)
mese = Month(DataF)
giorno = Day(DataF)

If Len(mese) = 1 Then mese = "0" & mese
If Len(giorno) = 1 Then giorno = "0" & giorno

GET_DATA_FORMATTATA = Anno & "-" & mese & "-" & giorno

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & " n. "

If Len(Trim(fnNotNull(rsGrigliaSelDocTMP!Doc_prefisso))) > 0 Then
    GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & Trim(fnNotNull(rsGrigliaSelDocTMP!Doc_prefisso)) & "-"
End If

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & rsGrigliaSelDocTMP!Doc_Numero

Exit Function
ERR_GET_DATA_FORMATTATA:

End Function
Private Function GET_PREFISSO(IDSezionale As Long) As String
On Error GoTo ERR_GET_PREFISSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_PREFISSO = ""

sSQL = "SELECT IDSezionale, Prefisso "
sSQL = sSQL & "FROM Sezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO = ""
Else
    GET_PREFISSO = Trim(fnNotNull(rs!Prefisso))
End If


rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PREFISSO:
    MsgBox Err.Description, vbCritical, "GET_PREFISSO"
End Function
