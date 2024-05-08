VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEliminaLav 
   Caption         =   "Elimina lavorazioni"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15690
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEliminaLav.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   15690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "DESELEZIONA TUTTO"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   7240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELEZIONA TUTTO"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7240
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   5160
      TabIndex        =   2
      Top             =   7440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   7240
      Width           =   2055
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   12515
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
      Left            =   5160
      TabIndex        =   3
      Top             =   7200
      Width           =   8295
   End
End
Attribute VB_Name = "frmEliminaLav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
        
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    
        With Me.GrigliaCorpo
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("Elimina", "Elimina", dgBoolean, True, 1000, dgAligncenter, True, True, False)
                cl.Editable = True
            
            .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "ID", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft, True, True, False
            Set cl = .ColumnsHeader.Add("Qta_UM", "Quantità", dgDouble, True, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                'cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 3
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 900, dgAlignRight, True, True, False)
                'cl.Editable = True
                'c'l.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 1
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
            Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
            Set cl = .ColumnsHeader.Add("PesoNetto", "PesoNetto", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                
            .ColumnsHeader.Add "Pezzi", "Pezzi", dgDouble, False, 1100, dgAlignRight, True, True, False
           
            
            
            .ColumnsHeader.Add "CodicePedana", "CodicePedana", dgchar, True, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", dgInteger, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "AnnoProcesso", "Anno processo IV gamma", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NumeroProcesso", "Numero processo IV gamma", dgchar, True, 1700, dgAlignleft, True, True, False
                                                                            
                                         
            Set .Recordset = rsLavorazioni
            .LoadUserSettings
            .Refresh
        End With
        
        CnDMT.CursorLocation = OLDCursor
        
        Me.lblInfo.Caption = "Numero lavorazioni: " & N_EliminaLav
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim IDTipoOggetto As Long
Dim NumeroElaborazioni As Long
Dim UnitaProgresso As Long

IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")

Me.ProgressBar1 = 0
Me.ProgressBar1.Max = N_EliminaLav
UnitaProgresso = 1
NumeroElaborazioni = 1
rsLavorazioni.MoveFirst

While Not rsLavorazioni.EOF

    Me.lblInfo.Caption = "ELABORAZIONE " & NumeroElaborazioni & " di " & N_EliminaLav
    DoEvents
    If (fnNotNullN(rsLavorazioni!Elimina) = 1) Then
        EliminaMovimento fnNotNullN(rsLavorazioni!IDRV_POAssegnazioneMerce), IDTipoOggetto
    End If

    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    DoEvents
    NumeroElaborazioni = NumeroElaborazioni + 1
rsLavorazioni.MoveNext
DoEvents
Wend


MsgBox "Operazione avvenuta con successo", vbInformation, "Conferma"

Unload Me

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub
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
Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
    rsLavorazioni.Filter = "Elimina=0"
    
    While Not rsLavorazioni.EOF
        rsLavorazioni!Elimina = 1
        rsLavorazioni.Update
    rsLavorazioni.MoveNext
    DoEvents
    Wend
       
    rsLavorazioni.Filter = vbNullString
    GET_GRIGLIA
Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Seleziona tutto"
End Sub

Private Sub Command2_Click()
On Error GoTo ERR_Command2_Click
    rsLavorazioni.Filter = "Elimina=1"
    
    While Not rsLavorazioni.EOF
        rsLavorazioni!Elimina = 0
        rsLavorazioni.Update
    rsLavorazioni.MoveNext
    DoEvents
    Wend
    
    
    rsLavorazioni.Filter = vbNullString
    GET_GRIGLIA
Exit Sub
ERR_Command2_Click:
    MsgBox Err.Description, vbCritical, "Deseleziona tutto"
End Sub

Private Sub Form_Load()
    GET_GRIGLIA
    
End Sub
Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaCorpo_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaCorpo.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaCorpo.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaCorpo.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsLavorazioni.Fields("Elimina").Value), 2
            End If
        End If
    End If
Exit Sub
ERR_GrigliaCorpo_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_MouseUp"
End Sub

Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaCorpo_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsLavorazioni.Fields("Elimina").Value), 2
        End If
    End If
Exit Sub
ERR_GrigliaCorpo_KeyPress:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_KeyPress"
End Sub

Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsLavorazioni.EOF And Not rsLavorazioni.BOF Then
        rsLavorazioni.Fields("Elimina").Value = Abs(CLng(Selected))
        'sbCheckSelected
        Me.GrigliaCorpo.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
Private Function EliminaMovimento(IDRiga As Long, IDTipoOggetto As Long) As Boolean
On Error GoTo ERR_EliminaMovimento
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim IDMovimento_Local As Long

EliminaMovimento = True

sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDRiga
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection

While Not rs.EOF
    EliminaMovimento = mov.Delete(fnNotNullN(rs!IDMovimento))
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Set mov = Nothing

sSQL = "DELETE RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDRiga
CnDMT.Execute sSQL

Exit Function
ERR_EliminaMovimento:
    MsgBox Err.Description, vbCritical, "EliminaMovimento"


End Function
