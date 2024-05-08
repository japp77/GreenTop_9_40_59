VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form frmContratto 
   Caption         =   "SELEZIONA CONTRATTO"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContratto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   21255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiltri 
      Caption         =   "FILTRI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   21015
      Begin VB.TextBox txtCausaleTrasporto 
         Height          =   315
         Left            =   11280
         TabIndex        =   10
         Top             =   1050
         Width           =   2175
      End
      Begin VB.TextBox txtDescrContratto 
         Height          =   315
         Left            =   6480
         TabIndex        =   9
         Top             =   1050
         Width           =   4695
      End
      Begin VB.TextBox txtNumeroPressoCliente 
         Height          =   315
         Left            =   4200
         TabIndex        =   8
         Top             =   1050
         Width           =   2175
      End
      Begin DMTDATETIMELib.dmtDate txtScadenzaContratto 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   1050
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DmtSearchAccount.DmtSearchACS ACSCliente 
         Height          =   585
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   1032
         WidthCode       =   900
         WidthDescription=   3500
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideLeaf        =   0   'False
         BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionDescription=   "Cliente"
         CaptionCode     =   "Codice"
         IDSearchTypeConto=   6
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboAltroSito 
         Height          =   315
         Left            =   5760
         TabIndex        =   1
         Top             =   450
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
      Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
         Height          =   315
         Left            =   8640
         TabIndex        =   2
         Top             =   450
         Width           =   2535
         _ExtentX        =   4471
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
      Begin DMTDataCmb.DMTCombo cboVettore 
         Height          =   315
         Left            =   11280
         TabIndex        =   3
         Top             =   450
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
      Begin DMTDataCmb.DMTCombo cboContrattoChiuso 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
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
      Begin DMTDataCmb.DMTCombo cboVettoreSuccessivo 
         Height          =   315
         Left            =   14160
         TabIndex        =   4
         Top             =   450
         Width           =   2655
         _ExtentX        =   4683
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
      Begin DMTDataCmb.DMTCombo cboRaggrFatturato 
         Height          =   315
         Left            =   16920
         TabIndex        =   5
         Top             =   450
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
         Caption         =   "Causale trasporto"
         Height          =   255
         Index           =   6
         Left            =   11280
         TabIndex        =   22
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Descrizione contratto"
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Numero presso cliente"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Raggruppamento fatturato"
         Height          =   255
         Index           =   0
         Left            =   16920
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Vettore successivo"
         Height          =   255
         Index           =   3
         Left            =   14160
         TabIndex        =   18
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Scadenza contratto"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Contratto chiuso"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Vettore"
         Height          =   255
         Index           =   2
         Left            =   11280
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Luogo di presa merce"
         Height          =   255
         Index           =   10
         Left            =   8760
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Altra destinazione"
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7455
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   21015
      _ExtentX        =   37068
      _ExtentY        =   13150
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
End
Attribute VB_Name = "frmContratto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private bloading As Boolean

Private Sub INIT_CONTROLLI()
    Set Me.ACSCliente.Connection = TheApp.Database.Connection
    ACSCliente.ApplicationName = App.Title
    ACSCliente.Client = App.EXEName
    ACSCliente.IDFirm = TheApp.IDFirm
    ACSCliente.IDUser = TheApp.IDUser
    ACSCliente.UserName = TheApp.User
    ACSCliente.SearchType = DmtSearchCustomers
    ACSCliente.HwndContainer = Me.hwnd
    
    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica  "
        .SQL = .SQL & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
    
    With Me.cboVettore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT IDVettore, Vettore FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
    End With
    
    With Me.cboVettoreSuccessivo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT IDVettore, Vettore FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
    End With
    
    With Me.cboContrattoChiuso
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POSiNO"
        .DisplayField = "SiNO"
        .SQL = "SELECT * FROM RV_POSiNO"
    End With

    With Me.cboRaggrFatturato
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRaggruppamentoFatturato"
        .DisplayField = "RaggruppamentoFatturato"
        .SQL = "SELECT * FROM RaggruppamentoFatturato"
        .SQL = .SQL & " ORDER BY RaggruppamentoFatturato"
    End With
    
End Sub
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub ACSCliente_ChangedElement()
    If bloading = False Then GET_GRIGLIA
End Sub
Private Sub cboAltroSito_Click()
    If bloading = False Then GET_GRIGLIA
End Sub

Private Sub cboContrattoChiuso_Click()
    If bloading = False Then GET_GRIGLIA
End Sub

Private Sub cboLuogoPresaMerce_Click()
    If bloading = False Then GET_GRIGLIA
End Sub

Private Sub cboRaggrFatturato_Click()
If bloading = False Then GET_GRIGLIA
End Sub

Private Sub cboVettore_Click()
    If bloading = False Then GET_GRIGLIA
End Sub

Private Sub cboVettoreSuccessivo_Click()
If bloading = False Then GET_GRIGLIA
End Sub

Private Sub Form_Activate()
    GET_GRIGLIA
End Sub

Private Sub Form_Load()
    bloading = True

    INIT_CONTROLLI
    
    PREPARAZIONE_FILTRI
    
    bloading = False
    
End Sub
Private Sub PREPARAZIONE_FILTRI()
    If LINK_CLIENTE_CONTRATTO > 0 Then
        Me.ACSCliente.sbLoadCFByIDAnagrafica 0, LINK_CLIENTE_CONTRATTO
    End If
    Me.txtScadenzaContratto.Value = Date 'DA PARAMETRIZZARE
    Me.cboContrattoChiuso.WriteOn 2 'DA PARAMETRIZZARE
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * FROM RV_POIEContrattoSel "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
If Me.ACSCliente.IDAnagrafica > 0 Then
    sSQL = sSQL & " AND Link_Nom_anagrafica=" & Me.ACSCliente.IDAnagrafica
End If
If Me.cboAltroSito.CurrentID > 0 Then
    sSQL = sSQL & " AND Link_Nom_ult_sito=" & Me.cboAltroSito.CurrentID
End If
If Me.cboLuogoPresaMerce.CurrentID > 0 Then
    sSQL = sSQL & " AND RV_POIDLuogoPresaMerce=" & Me.cboLuogoPresaMerce.CurrentID
End If
If Me.cboVettore.CurrentID > 0 Then
    sSQL = sSQL & " AND Link_Vet_vettore=" & Me.cboVettore.CurrentID
End If
If Me.cboContrattoChiuso.CurrentID > 0 Then
    If Me.cboContrattoChiuso.CurrentID = 1 Then
        sSQL = sSQL & " AND RV_POContrattoChiuso=1"
    End If
    If Me.cboContrattoChiuso.CurrentID = 2 Then
        sSQL = sSQL & " AND RV_POContrattoChiuso=0"
    End If
End If

If Me.cboVettoreSuccessivo.CurrentID > 0 Then
    sSQL = sSQL & " AND RV_POIDTrasportatoreSuccessivo=" & Me.cboVettoreSuccessivo.CurrentID
End If
If Me.cboRaggrFatturato.CurrentID > 0 Then
    sSQL = sSQL & " AND Link_Nom_raggrup_fatturato=" & Me.cboRaggrFatturato.CurrentID
End If
If Len(Trim(Me.txtNumeroPressoCliente.Text)) > 0 Then
    sSQL = sSQL & " AND Doc_numero_vs_ordine_di_rifer LIKE " & fnNormString("%" & Me.txtNumeroPressoCliente.Text & "%")
End If
If Len(Trim(Me.txtCausaleTrasporto.Text)) > 0 Then
    sSQL = sSQL & " AND Doc_causale_trasporto LIKE " & fnNormString("%" & Me.txtCausaleTrasporto.Text & "%")
End If
If Len(Trim(Me.txtDescrContratto.Text)) > 0 Then
    sSQL = sSQL & " AND Doc_descrizione_offerta LIKE " & fnNormString("%" & Me.txtDescrContratto.Text & "%")
End If
If Me.txtScadenzaContratto.Value > 0 Then
    sSQL = sSQL & " AND ((Doc_data_scadenza>=" & fnNormDate(Me.txtScadenzaContratto.Text) & ") OR (Doc_data_scadenza IS NULL))"
End If


sSQL = sSQL & "  ORDER BY doc_data, doc_numero"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Doc_data_scadenza", "Data scadenza", dgDate, True, 1500, dgAlignleft
    .ColumnsHeader.Add "Doc_data", "Data inizio validità", dgDate, False, 1500, dgAlignleft
    .ColumnsHeader.Add "Doc_numero", "Numero documento", dgNumeric, False, 1500, dgAlignRight
    .ColumnsHeader.Add "RV_POContrattoChiuso", "Chiuso", dgBoolean, True, 1500, dgAligncenter
    .ColumnsHeader.Add "Doc_descrizione_offerta", "Descrizione offerta", dgchar, False, 3500, dgAlignleft
    
    .ColumnsHeader.Add "Link_Nom_anagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 3500, dgAlignleft
    .ColumnsHeader.Add "Nom_nome", "Nome cliente", dgchar, False, 3500, dgAlignleft
    .ColumnsHeader.Add "Doc_causale_trasporto", "Causale trasporto", dgchar, True, 2000, dgAlignleft
    .ColumnsHeader.Add "Link_Nom_ult_sito", "IDDestinazioneDiversa", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "DestinazioneDiversa", "Destinazione merce", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "Link_Vet_vettore", "IDVettore", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Vettore", "Vettore", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POTargaAutomezzo", "Targa", dgchar, True, 2000, dgAlignleft
    .ColumnsHeader.Add "RV_POIstruzioniMittente", "Istruzioni al mittente", dgchar, False, 4000, dgAlignleft
    
    .ColumnsHeader.Add "RV_POIDLuogoPresaMerce", "IDLuogoPresaMerce", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "LuogoPresaMerce", "Luogo presa merce", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "RV_POIDTrasportatoreSuccessivo", "IDVettoreSuccessivo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "VettoreSuccessivo", "Vettore successivo", dgchar, True, 2500, dgAlignleft
    
    .ColumnsHeader.Add "Doc_data_vs_ordine_di_rifer", "Data presso cliente", dgDate, False, 1500, dgAlignleft
    .ColumnsHeader.Add "Doc_numero_vs_ordine_di_rifer", "Numero presso cliente", dgchar, True, 2000, dgAlignleft
    
    .ColumnsHeader.Add "Link_Nom_raggrup_fatturato", "IDRaggruppamento fatturato", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "RaggruppamentoFatturato", "Raggruppamento fatturato", dgchar, True, 2000, dgAlignleft
    
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub Griglia_DblClick()
    LINK_CONTRATTO = fnNotNullN(Me.Griglia.AllColumns("IDOggetto").Value)
    Unload Me
End Sub
Private Sub txtCausaleTrasporto_LostFocus()
    If bloading = False Then GET_GRIGLIA
End Sub
Private Sub txtDescrContratto_LostFocus()
    If bloading = False Then GET_GRIGLIA
End Sub
Private Sub txtNumeroPressoCliente_LostFocus()
    If bloading = False Then GET_GRIGLIA
End Sub
Private Sub txtScadenzaContratto_LostFocus()
    If bloading = False Then GET_GRIGLIA
End Sub
