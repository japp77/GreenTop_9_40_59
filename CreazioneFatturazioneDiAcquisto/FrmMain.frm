VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Passaggio in fattura di acquisto della liquidazione (Passo 1 di 2)"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di ricerca"
      ForeColor       =   &H00400000&
      Height          =   2655
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtNomeSocioFatt 
         BackColor       =   &H8000000F&
         Height          =   340
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtNomeSocio 
         BackColor       =   &H8000000F&
         Height          =   340
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkFatturate 
         Caption         =   "Liquidazioni già fatturate"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton CmdRicerca 
         Caption         =   "RICERCA"
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
         Left            =   4680
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin DmtCodDescCtl.DmtCodDesc CDPeriodoLiquidazione 
         Height          =   615
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1085
         PropCodice      =   $"FrmMain.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmMain.frx":47872
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmMain.frx":478D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         TableName       =   "e"
      End
      Begin DmtCodDescCtl.DmtCodDesc CDSocio 
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         PropCodice      =   $"FrmMain.frx":4792F
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmMain.frx":4797D
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmMain.frx":479CD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         PropCodice      =   $"FrmMain.frx":47A27
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmMain.frx":47A75
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmMain.frx":47ADA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvLiquidazioni 
      Height          =   2295
      Left            =   2880
      TabIndex        =   15
      Top             =   5760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdTutto 
      Caption         =   "Sposta tutto"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSu 
      Caption         =   "Sposta su"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGiu 
      Caption         =   "Sposta giù"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvRicerca 
      Height          =   1935
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   8160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   0
      Picture         =   "FrmMain.frx":47B34
      ScaleHeight     =   4785
      ScaleWidth      =   2745
      TabIndex        =   8
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Liquidazioni "
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   16
      Top             =   5520
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Liquidazioni "
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   2760
      Width           =   6375
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1

'Variabili recordset per le visualizzazioni delle griglie

Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    Set CnDMT = m_App.Database.Connection
    
    PrelevaAzienda
    
    VarPassword = m_App.Password
    VarUtente = m_App.User
    ParametroSocio
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property
Private Sub CDPeriodoLiquidazione_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Me.CDPeriodoLiquidazione.KeyFieldID

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_IMPORTO_DOCUMENTO_LIQ = 0
Else
    LINK_TIPO_IMPORTO_DOCUMENTO_LIQ = fnNotNullN(rs!IDTipoImportoDocumento)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della liquidazione?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If
End Sub

Public Sub InitControlli()
    With Me.CDPeriodoLiquidazione
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
        .CodeField = "NumeroLiquidazione"
        .DescriptionField = "Periodo"
        .KeyField = "IDRV_POLiquidazionePeriodo"
        .TableName = "RV_POLiquidazionePeriodo"
        .Filter = "IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Numero liquidazione"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Periodo liquidazione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Numero liquidazione"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Periodo liquidazione"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = True
    End With
    'Anagrafica socio
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    'Inizializza la ListView contenente la ricerca
    With Me.lvRicerca
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "N° Liq.", 800
        .ColumnHeaders.Add , , "Socio", 3000
        .ColumnHeaders.Add , , "Imp. Liq.", 1500, lvwColumnRight
        .ColumnHeaders.Add , , "Ufficiale", 2000
    End With

    'Inizializza la ListView contenente la ricerca
    With Me.lvLiquidazioni
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "N° Liq.", 800
        .ColumnHeaders.Add , , "Socio", 3000
        .ColumnHeaders.Add , , "Imp. Liq.", 1500, lvwColumnRight
    End With
    
    With Me.CDSocioFatt
        Set .Application = m_App
        Set .Database = TheApp.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
End Sub

Private Sub cmdAvanti_Click()
    If Me.lvLiquidazioni.ListItems.Count = 0 Then
        MsgBox "Non è stata inserita nessuna liquidazione", vbCritical, TheApp.FunctionName
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdGiu_Click()
On Error GoTo ERR_cmdGiu_Click
Dim sSQL As String
Dim IDLiquidazione As Long
IDLiquidazione = 0
If Me.lvRicerca.ListItems.Count > 0 Then
    IDLiquidazione = fnNotNullN(Me.lvRicerca.SelectedItem)
    If IDLiquidazione > 0 Then
        If ControlloEsistenzaLiquidazioneInFatturazione(IDLiquidazione, "RV_POTMPLiqDaFatt") = False Then
            sSQL = "DELETE FROM RV_POTMPLiqRic WHERE IDLiquidazione=" & IDLiquidazione
            CnDMT.Execute sSQL
            
            sSQL = "INSERT INTO RV_POTMPLiqDaFatt ("
            sSQL = sSQL & "IDLiquidazione) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & IDLiquidazione & ")"
            
            CnDMT.Execute sSQL
            
        Else
            MsgBox "La liquidazione è presente tra le liquidazioni da fatturare", vbInformation, "Creazione fatturazione"
        End If
    End If
GET_GRIGLIA_LIQUIDAZIONI
GET_GRIGLIA_RICERCA
End If

Exit Sub
ERR_cmdGiu_Click:
    MsgBox Err.Description, vbCritical, "cmdGiu_Click"

End Sub

Private Sub cmdReset_Click()
On Error GoTo ERR_cmdReset_Click
Dim sSQL As String
Dim IDLiquidazione As Long
IDLiquidazione = 0

For I = 1 To Me.lvLiquidazioni.ListItems.Count
IDLiquidazione = fnNotNullN(Me.lvLiquidazioni.ListItems(I))
    If IDLiquidazione > 0 Then
        If ControlloEsistenzaLiquidazioneInFatturazione(IDLiquidazione, "RV_POTMPLiqRic") = False Then
            sSQL = "DELETE FROM RV_POTMPLiqDaFatt WHERE IDLiquidazione=" & IDLiquidazione
            CnDMT.Execute sSQL
            
            sSQL = "INSERT INTO RV_POTMPLiqRic ("
            sSQL = sSQL & "IDLiquidazione) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & IDLiquidazione & ")"
            
            CnDMT.Execute sSQL
            
        End If
    End If
Next
GET_GRIGLIA_LIQUIDAZIONI
GET_GRIGLIA_RICERCA

Exit Sub
ERR_cmdReset_Click:
    MsgBox Err.Description, vbCritical, "cmdReset_Click"
End Sub

Private Sub CmdRicerca_Click()
On Error GoTo ERR_CmdRicerca_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL_WHERE As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem

CnDMT.Execute "DELETE FROM RV_POTMPLiqRic"
CnDMT.Execute "DELETE FROM RV_POTMPLiqDaFatt"

sSQL = "SELECT RV_POLiquidazione.IDRV_POLiquidazione "
sSQL = sSQL & "FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE PassaggioInFatturazione=" & fnNormBoolean(Me.chkFatturate.Value)
sSQL = sSQL & " AND IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND Ufficiale=" & fnNormBoolean(1)

If Me.CDPeriodoLiquidazione.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POLiquidazione.IDRV_POLiquidazionePeriodo=" & Me.CDPeriodoLiquidazione.KeyFieldID
End If
If Me.CDSocio.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POLiquidazione.IDAnagrafica=" & Me.CDSocio.KeyFieldID
End If
If Me.CDSocioFatt.KeyFieldID > 0 Then
    sSQL = sSQL & " AND RV_POLiquidazione.IDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
End If
Set rs = CnDMT.OpenResultset(sSQL)


While Not rs.EOF
    sSQL = "INSERT INTO RV_POTMPLiqRic ("
    sSQL = sSQL & "IDLiquidazione) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnNotNullN(rs!IDRV_POLiquidazione) & ")"

    
CnDMT.Execute sSQL
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

GET_GRIGLIA_RICERCA
Exit Sub
ERR_CmdRicerca_Click:
    MsgBox Err.Description, vbCritical, "CmdRicerca_Click"
End Sub

Private Sub cmdSu_Click()
On Error GoTo ERR_cmdSu_Click
Dim sSQL As String
Dim IDLiquidazione As Long
IDLiquidazione = 0
IDLiquidazione = fnNotNullN(Me.lvLiquidazioni.SelectedItem)
If IDLiquidazione > 0 Then
    If ControlloEsistenzaLiquidazioneInFatturazione(IDLiquidazione, "RV_POTMPLiqRic") = False Then
        sSQL = "DELETE FROM RV_POTMPLiqDaFatt WHERE IDLiquidazione=" & IDLiquidazione
        CnDMT.Execute sSQL
        
        sSQL = "INSERT INTO RV_POTMPLiqRic ("
        sSQL = sSQL & "IDLiquidazione) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDLiquidazione & ")"
        
        CnDMT.Execute sSQL
        
    End If
End If
GET_GRIGLIA_LIQUIDAZIONI
GET_GRIGLIA_RICERCA
Exit Sub
ERR_cmdSu_Click:
    MsgBox Err.Description, vbCritical, "cmdSu_Click"
End Sub

Private Sub cmdTutto_Click()
On Error GoTo ERR_cmdTutto_Click
Dim sSQL As String
Dim IDLiquidazione As Long
IDLiquidazione = 0
For I = 1 To Me.lvRicerca.ListItems.Count
    
    IDLiquidazione = fnNotNullN(Me.lvRicerca.ListItems(I))
    If IDLiquidazione > 0 Then
        If ControlloEsistenzaLiquidazioneInFatturazione(IDLiquidazione, "RV_POTMPLiqDaFatt") = False Then
            sSQL = "DELETE FROM RV_POTMPLiqRic WHERE IDLiquidazione=" & IDLiquidazione
            CnDMT.Execute sSQL
            
            sSQL = "INSERT INTO RV_POTMPLiqDaFatt ("
            sSQL = sSQL & "IDLiquidazione) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & IDLiquidazione & ")"
            
            CnDMT.Execute sSQL
            
        Else
            MsgBox "La liquidazione è presente tra le liquidazioni da fatturare", vbInformation, "Creazione fatturazione"
        End If
    End If


Next
GET_GRIGLIA_LIQUIDAZIONI
GET_GRIGLIA_RICERCA
Exit Sub
ERR_cmdTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdTutto_Click"
End Sub

Private Sub Form_Load()
    If BLoading = False Then
        BLoading = True
    Else
        InitControlli
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        FrmFine.Show
    Exit Sub
    End If
End Sub


Private Sub ParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSocio = rs!IDCategoriaAnagrafica
Else
    Link_TipoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_GRIGLIA_RICERCA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem

Me.lvRicerca.ListItems.Clear

sSQL = "SELECT RV_POTMPLiqRic.IDLiquidazione, RV_POLiquidazione.NumeroLiquidazione, RV_POLiquidazione.IDAnagrafica, Anagrafica.Anagrafica, RV_POLiquidazione.NettoLiquidazione "
sSQL = sSQL & "FROM RV_POTMPLiqRic INNER JOIN "
sSQL = sSQL & "RV_POLiquidazione ON RV_POTMPLiqRic.IDLiquidazione = RV_POLiquidazione.IDRV_POLiquidazione INNER JOIN "
sSQL = sSQL & "Fornitore ON RV_POLiquidazione.IDAnagrafica = Fornitore.IDAnagrafica INNER JOIN "
sSQL = sSQL & "Anagrafica ON Fornitore.IDAnagrafica = Anagrafica.IDAnagrafica "
sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & VarIDAzienda
sSQL = sSQL & "ORDER BY NumeroLiquidazione"


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Set oItem = Me.lvRicerca.ListItems.Add
    
    'Popola l 'item della listview
    oItem.Text = fnNotNullN(rs!IDLiquidazione)
    oItem.SubItems(1) = fnNotNullN(rs!NumeroLiquidazione)
    oItem.SubItems(2) = fnNotNull(rs!Anagrafica)
    oItem.SubItems(3) = FormatNumber((rs!NettoLiquidazione), 2)
    
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub GET_GRIGLIA_LIQUIDAZIONI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem

Me.lvLiquidazioni.ListItems.Clear

sSQL = "SELECT RV_POTMPLiqDaFatt.IDLiquidazione, RV_POLiquidazione.NumeroLiquidazione, RV_POLiquidazione.IDAnagrafica, Anagrafica.Anagrafica, RV_POLiquidazione.NettoLiquidazione "
sSQL = sSQL & "FROM RV_POTMPLiqDaFatt INNER JOIN "
sSQL = sSQL & "RV_POLiquidazione ON RV_POTMPLiqDaFatt.IDLiquidazione = RV_POLiquidazione.IDRV_POLiquidazione INNER JOIN "
sSQL = sSQL & "Fornitore ON RV_POLiquidazione.IDAnagrafica = Fornitore.IDAnagrafica INNER JOIN "
sSQL = sSQL & "Anagrafica ON Fornitore.IDAnagrafica = Anagrafica.IDAnagrafica "
sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & VarIDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Set oItem = Me.lvLiquidazioni.ListItems.Add
    
    'Popola l 'item della listview
    oItem.Text = fnNotNullN(rs!IDLiquidazione)
    oItem.SubItems(1) = fnNotNullN(rs!NumeroLiquidazione)
    oItem.SubItems(2) = fnNotNull(rs!Anagrafica)
    oItem.SubItems(3) = FormatNumber((rs!NettoLiquidazione), 2)
    
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Sub
Private Function ControlloEsistenzaLiquidazioneInFatturazione(IDLiquidazione As Long, Tabelle As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM " & Tabelle & " WHERE IDLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    ControlloEsistenzaLiquidazioneInFatturazione = False
Else
    ControlloEsistenzaLiquidazioneInFatturazione = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
