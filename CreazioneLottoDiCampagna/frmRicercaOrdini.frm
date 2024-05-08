VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRicercaOrdini 
   Caption         =   "RICERCA ORDINI DA CLIENTE"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin DmtGridCtl.DmtGrid GrigliaCond 
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   13573
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
      GuiMode         =   1
   End
   Begin VB.CommandButton cmdApplicaFiltro 
      Caption         =   "APPLICA IL FILTRO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "frmRicercaOrdini.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Applica filtro"
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton cmdRicerca 
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
      Left            =   120
      Picture         =   "frmRicercaOrdini.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ricerca dettagliata"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA SELEZIONI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17760
      TabIndex        =   2
      Top             =   60
      Width           =   2415
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7695
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   13573
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmRicercaOrdini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub cmdApplicaFiltro_Click()
On Error GoTo ERR_cmdApplicaFiltro_Click
    
    
    APPLICA_FILTRO
    Me.Griglia.ApplyFilter
    
    'Me.Griglia.GuiMode = dgFilterDefinition

Exit Sub
ERR_cmdApplicaFiltro_Click:
    MsgBox Err.Description, vbCritical, "cmdApplicaFiltro_Click"
End Sub

Private Sub cmdConferma_Click()
Dim sSQL As String
Dim SERRE As String

rsGriglia.Filter = "Registra=1"

While Not rsGriglia.EOF

    SERRE = GET_STRINGA_SERRE(Link_LottoCampagna)
    
    sSQL = "UPDATE ValoriOggettoDettaglio0010 SET "
    sSQL = sSQL & "RV_PO01_UbiLottoDiCampagna=" & fnNormString(SERRE) & ", "
    sSQL = sSQL & "RV_PO01_LottoDiCampagna=" & fnNormString(frmMain.txtCodiceLotto.Text) & ", "
    sSQL = sSQL & "RV_PO01_DescrLottoDiCampagna=" & fnNormString(frmMain.txtDescrizioneLotto.Text) & ", "
    sSQL = sSQL & "RV_PO01_IDLottoCampagna=" & Link_LottoCampagna & ", "
    sSQL = sSQL & "RV_PO01_DataSemina=" & fnNormDate(frmMain.txtDataSemina.Text) & " "
    sSQL = sSQL & "WHERE RV_POLinkRiga=" & fnNotNullN(rsGriglia!Link_Riga)
    sSQL = sSQL & " AND IDOggetto=" & fnNotNullN(rsGriglia!IDOggetto)
    
    Cn.Execute sSQL
    
rsGriglia.MoveNext
Wend

Unload Me

End Sub

Private Sub cmdRicerca_Click()
    Me.Griglia.GuiMode = dgFilterDefinition
End Sub

Private Sub Form_Activate()
    INSERISCI_ORDINI_CLIENTE
        
    GET_GRIGLIA
    
    CREA_CONDIZIONI
End Sub

Private Sub Form_Load()
    CREA_RECORDSET_TMP
End Sub
Private Sub CREA_RECORDSET_TMP()
Dim sSQL As String
Dim I As Integer

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    
    Set rsGriglia = Nothing
End If

Set rsGriglia = New ADODB.Recordset

rsGriglia.Fields.Append "IDValoriOggettoDettaglio", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Link_Riga", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "DataOrdine", adDBTimeStamp, , adFldIsNullable
rsGriglia.Fields.Append "NumeroOrdine", adVarChar, 50, adFldIsNullable
rsGriglia.Fields.Append "IDAnagrafica", adInteger, , adFldMayBeNull
rsGriglia.Fields.Append "Anagrafica", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Codice", adVarChar, 50, adFldIsNullable
rsGriglia.Fields.Append "Nome", adVarChar, 50, adFldIsNullable
rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldMayBeNull
rsGriglia.Fields.Append "CodiceArticolo", adVarChar, 50, adFldIsNullable
rsGriglia.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Quantita", adDouble, 250, adFldIsNullable
rsGriglia.Fields.Append "CodiceLotto", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "DescrizioneLotto", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "DataSemina", adDBTimeStamp, , adFldIsNullable
rsGriglia.Fields.Append "DataEvasione", adDBTimeStamp, , adFldIsNullable
rsGriglia.Fields.Append "DataEvasioneDoc", adDBTimeStamp, , adFldIsNullable
rsGriglia.Fields.Append "Registra", adSmallInt, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic



End Sub
Private Sub INSERISCI_ORDINI_CLIENTE()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Const NumeroRecord As Long = 1000
Dim ICont As Long


'sSQL = "SELECT ValoriOggettoDettaglio0010.IDValoriOggettoDettaglio, ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, "
'sSQL = sSQL & "ValoriOggettoDettaglio0010.Link_Art_articolo, ValoriOggettoDettaglio0010.Art_codice, ValoriOggettoDettaglio0010.Art_descrizione, "
'sSQL = sSQL & "ValoriOggettoDettaglio0010.Art_quantita_totale, ValoriOggettoPerTipo000F.Doc_data, ValoriOggettoPerTipo000F.Doc_numero, "
'sSQL = sSQL & "ValoriOggettoPerTipo000F.Link_Nom_anagrafica, ValoriOggettoPerTipo000F.Nom_nome, ValoriOggettoPerTipo000F.Nom_codice, "
'sSQL = sSQL & "ValoriOggettoPerTipo000F.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo000F.Doc_ordine_chiuso, "
'sSQL = sSQL & "ValoriOggettoDettaglio0010.RV_PO01_UbiLottoDiCampagna, ValoriOggettoDettaglio0010.RV_PO01_LottoDiCampagna, "
'sSQL = sSQL & "ValoriOggettoDettaglio0010.Art_data_prevista_evasione, ValoriOggettoDettaglio0010.RV_PO01_DescrLottoDiCampagna, "
'sSQL = sSQL & "ValoriOggettoDettaglio0010.RV_PO01_DataSemina, ValoriOggettoDettaglio0010.RV_POLinkRiga "
'sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
'sSQL = sSQL & "ValoriOggettoPerTipo000F ON ValoriOggettoDettaglio0010.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto INNER JOIN "
'sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto "

sSQL = "SELECT * FROM RV_PO01_IETrovaOrdinePerLotto "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Doc_ordine_chiuso<>1 "
sSQL = sSQL & " AND Art_quantita_totale>0 "
sSQL = sSQL & " AND RV_POTipoRiga=1 "
sSQL = sSQL & " AND IDRV_PO01_LottoCampagna=" & Link_LottoCampagna
sSQL = sSQL & " AND ((RV_PO01_IDLottoCampagna=0)"
sSQL = sSQL & " OR (RV_PO01_IDLottoCampagna IS NULL)) "
sSQL = sSQL & " ORDER BY Doc_data DESC, Doc_numero DESC"

Set rs = Cn.OpenResultset(sSQL)
ICont = 1
Me.ProgressBar1.Value = 0
While Not rs.EOF
    rsGriglia.AddNew
        rsGriglia!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
        rsGriglia!IDOggetto = fnNotNullN(rs!IDOggetto)
        rsGriglia!DataOrdine = rs!Doc_data
        rsGriglia!NumeroOrdine = fnNotNull(rs!Doc_numero)
        rsGriglia!IDAnagrafica = fnNotNullN(rs!Link_nom_anagrafica)
        rsGriglia!Anagrafica = fnNotNull(rs!Nom_ragione_sociale_o_cognome)
        rsGriglia!Codice = fnNotNull(rs!nom_codice)
        rsGriglia!Nome = fnNotNull(rs!Nom_nome)
        rsGriglia!IDArticolo = fnNotNullN(rs!Link_Art_articolo)
        rsGriglia!CodiceArticolo = fnNotNull(rs!Art_codice)
        rsGriglia!Articolo = fnNotNull(rs!Art_descrizione)
        rsGriglia!Quantita = fnNotNullN(rs!Art_quantita_totale)
        rsGriglia!CodiceLotto = fnNotNull(rs!RV_PO01_LottoDiCampagna)
        rsGriglia!DescrizioneLotto = fnNotNull(rs!RV_PO01_DescrLottoDiCampagna)
        rsGriglia!DataEvasione = rs!Art_data_prevista_evasione
        rsGriglia!DataSemina = rs!RV_PO01_DataSemina
        rsGriglia!Registra = 0
        rsGriglia!Link_Riga = fnNotNullN(rs!RV_POLinkRiga)
        rsGriglia!DataEvasioneDoc = rs!Doc_data_prevista_evasione
    rsGriglia.Update
    
    If (Me.ProgressBar1.Value + ICont) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + ICont
    End If
ICont = ICont + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Me.ProgressBar1.Visible = False

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
            .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
            Set cl = .ColumnsHeader.Add("Registra", "Seleziona", dgBoolean, True, 1300, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "DataOrdine", "Data documento", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "NumeroOrdine", "Numero documento", dgchar, True, 1800, dgAlignRight
            .ColumnsHeader.Add "DataEvasione", "Data evasione", dgDate, True, 1800, dgAlignRight
            .ColumnsHeader.Add "DataEvasioneDoc", "Data evasione doc.", dgDate, True, 1800, dgAlignRight
            .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Anagrafica", "Cliente", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Nome", "Nome", dgchar, True, 1800, dgAlignleft
            
            .ColumnsHeader.Add "IDArticolo", "IDRV_PO08_ModelloVendita", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 1800, dgAlignleft
            
            .ColumnsHeader.Add "CodiceLotto", "Codice lotto", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DescrizioneLotto", "Descrizione lotto", dgchar, True, 1800, dgAlignleft
            .ColumnsHeader.Add "DataSemina", "Data semina", dgDate, True, 1800, dgAligncenter
            
            Set cl = .ColumnsHeader.Add("Quantita", "Quantità", dgDouble, True, 1300, dgAlignRight)
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
                    
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With

Cn.CursorLocation = OLD_Cursor

GET_GRIGLIA_COND

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName


End Sub
Private Sub GET_GRIGLIA_COND()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long
Dim sSQL As String
Dim cl As DmtGridCtl.dgColumnHeader

OLD_Cursor = Cn.CursorLocation
Cn.CursorLocation = 3
        
    With Me.GrigliaCond
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        
    End With

Cn.CursorLocation = OLD_Cursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName


End Sub
Private Sub Form_Resize()
    Me.Griglia.Width = Me.Width - 500 - Me.GrigliaCond.Width
    
    Me.cmdConferma.Left = Me.Width - 330 - Me.cmdConferma.Width
    
    Me.Griglia.Height = Me.Height - Me.cmdConferma.Height - 800
    Me.GrigliaCond.Height = Me.Griglia.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsGriglia.Close
Set rsGriglia = Nothing
End Sub
Private Function GET_STRINGA_SERRE(IDLottoCampagna) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_STRINGA_SERRE = ""

sSQL = "SELECT RV_PO01_SerraPerLotto.IDRV_PO01_Serra, RV_PO01_Serra.Codice, RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna "
sSQL = sSQL & "FROM RV_PO01_SerraPerLotto INNER JOIN "
sSQL = sSQL & "RV_PO01_Serra ON RV_PO01_SerraPerLotto.IDRV_PO01_Serra = RV_PO01_Serra.IDRV_PO01_Serra "
sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_STRINGA_SERRE = GET_STRINGA_SERRE & fnNotNull(rs!Codice) & "|"
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing


GET_STRINGA_SERRE = Mid(GET_STRINGA_SERRE, 1, 250)
End Function
Private Sub Griglia_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("Registra").Value)), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
                
        rsGriglia.Fields("Registra").Value = Abs(CLng(Selected))
        'sbCheckSelected
        

        rsGriglia.UpdateBatch
        Me.Griglia.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub

Private Sub CREA_CONDIZIONI()
On Error Resume Next
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    
    
    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        Me.Griglia.ConnectionString = MenuOptions.ConnectionString & "User Id=" & TheApp.User & ";Password=" & TheApp.Password
    Else
        Me.Griglia.ConnectionString = MenuOptions.ConnectionString & ";" & "User Id=" & TheApp.User & ";Password=" & TheApp.Password
    End If

    
    Me.Griglia.Conditions.Clear

    Me.Griglia.Conditions.WidthConditions = 200
    Me.Griglia.Conditions.WidthFields = 200
    Me.Griglia.Conditions.WidthIntervals = 100
    
    Me.Griglia.Title.BackColor = vb3DFace
    Me.Griglia.Title.ForeColor = vbBlack
    Me.Griglia.Title.Font.Bold = True
                
    Set Cond = Me.Griglia.Conditions.Add("CodiceLotto", "Codice lotto", "", False, False, , dgCondTypeText)
    Set Cond = Me.Griglia.Conditions.Add("DescrizioneLotto", "Descrizione lotto", "", False, False, , dgCondTypeText)
    Set Cond = Me.Griglia.Conditions.Add("Anagrafica", "Cliente", "", False, False, , dgCondTypeText)
    Set Cond = Me.Griglia.Conditions.Add("CodiceArticolo", "Codice articolo", "", False, False, , dgCondTypeText)
    Set Cond = Me.Griglia.Conditions.Add("Articolo", "Articolo", "", False, False, , dgCondTypeText)
    Set Cond = Me.Griglia.Conditions.Add("DataOrdine", "Data ordine", "", False, True, , dgCondTypeDate)
    Set Cond = Me.Griglia.Conditions.Add("DataEvasione", "Data evasione", "", False, True, , dgCondTypeDate)
    Set Cond = Me.Griglia.Conditions.Add("DataEvasioneDoc", "Data evasione doc.", "", False, True, , dgCondTypeDate)
    Set Cond = Me.Griglia.Conditions.Add("DataSemina", "Data semina", "", False, True, , dgCondTypeDate)
    
    
    CREA_CONDIZIONI_COND

End Sub
Private Sub CREA_CONDIZIONI_COND()
'On Error Resume Next
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    
    
    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        Me.GrigliaCond.ConnectionString = MenuOptions.ConnectionString & "User Id=" & TheApp.User & ";Password=" & TheApp.Password
    Else
        Me.GrigliaCond.ConnectionString = MenuOptions.ConnectionString & ";" & "User Id=" & TheApp.User & ";Password=" & TheApp.Password
    End If

    
    Me.GrigliaCond.Conditions.Clear

    Me.GrigliaCond.Conditions.WidthConditions = 200
    Me.GrigliaCond.Conditions.WidthFields = 150
    Me.GrigliaCond.Conditions.WidthIntervals = 35
    
    Me.GrigliaCond.Title.BackColor = vb3DFace
    Me.GrigliaCond.Title.ForeColor = vbBlack
    Me.GrigliaCond.Title.Font.Bold = True
                
    Set Cond = Me.GrigliaCond.Conditions.Add("CodiceLotto", "Codice lotto", "", False, False, , dgCondTypeText)
    Set Cond = Me.GrigliaCond.Conditions.Add("DescrizioneLotto", "Descrizione lotto", "", False, False, , dgCondTypeText)
    Set Cond = Me.GrigliaCond.Conditions.Add("Anagrafica", "Cliente", "", False, False, , dgCondTypeText)
    Set Cond = Me.GrigliaCond.Conditions.Add("CodiceArticolo", "Codice articolo", "", False, False, , dgCondTypeText)
    Set Cond = Me.GrigliaCond.Conditions.Add("Articolo", "Articolo", "", False, False, , dgCondTypeText)
    Set Cond = Me.GrigliaCond.Conditions.Add("DataOrdine", "Data ordine", "", False, True, , dgCondTypeDate)
    Set Cond = Me.GrigliaCond.Conditions.Add("DataEvasione", "Data evasione lotto", "", False, True, , dgCondTypeDate)
    Set Cond = Me.GrigliaCond.Conditions.Add("DataEvasioneDoc", "Data evasione doc.", "", False, True, , dgCondTypeDate)
    
    Set Cond = Me.GrigliaCond.Conditions.Add("DataSemina", "Data semina", "", False, True, , dgCondTypeDate)
    
    
    Me.GrigliaCond.Refresh

End Sub
Private Sub APPLICA_FILTRO()
Dim I As Long




For I = 1 To Me.GrigliaCond.Conditions.Count

    Me.Griglia.Conditions(Me.GrigliaCond.Conditions(I).FieldName).FromValue = Me.GrigliaCond.Conditions(I).FromValue
    Me.Griglia.Conditions(Me.GrigliaCond.Conditions(I).FieldName).RangeChecked = Me.GrigliaCond.Conditions(I).RangeChecked
    If Me.GrigliaCond.Conditions(I).RangeChecked = True Then
        If Len(Trim(Me.GrigliaCond.Conditions(I).ToValue)) > 0 Then
             Me.Griglia.Conditions(Me.GrigliaCond.Conditions(I).FieldName).RangeChecked = Me.GrigliaCond.Conditions(I).RangeChecked
             Me.Griglia.Conditions(Me.GrigliaCond.Conditions(I).FieldName).ToValue = Me.GrigliaCond.Conditions(I).ToValue
        End If
    End If
Next

End Sub
