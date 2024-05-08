VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPesaturaPedana 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ripesatura pedana"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13725
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodicePedana 
      Height          =   285
      Left            =   0
      TabIndex        =   18
      Top             =   40
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5895
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   7800
      TabIndex        =   14
      Top             =   3720
      Width           =   5775
   End
   Begin VB.CommandButton cmdElabora 
      Caption         =   "ELABORA"
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
      Left            =   7800
      TabIndex        =   13
      Top             =   3240
      Width           =   5775
   End
   Begin VB.Frame fraScarti 
      Caption         =   "QUADRATURA"
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
      Height          =   1815
      Left            =   7800
      TabIndex        =   7
      Top             =   1320
      Width           =   5775
      Begin VB.CheckBox chkRiportaOrdine 
         Caption         =   "Riporta le pedane nell'ordine selezionato"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   5535
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticoloScarto 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         PropCodice      =   $"frmPesaturaPedana.frx":0000
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmPesaturaPedana.frx":004F
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmPesaturaPedana.frx":00BD
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
      Begin DMTDataCmb.DMTCombo CboTipoLavScarti 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
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
      Begin DMTDataCmb.DMTCombo cboUMScarti 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin VB.Label Label2 
         Caption         =   "Unità di misura"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   12
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo lavorazione"
         Height          =   255
         Index           =   12
         Left            =   2040
         TabIndex        =   11
         Top             =   765
         Width           =   2655
      End
   End
   Begin VB.Frame fraPesoPedana 
      Caption         =   "RISCONTRO PESO REALE"
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
      Height          =   1335
      Left            =   7800
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      Begin DMTEDITNUMLib.dmtNumber txtTaraTotaleImballi 
         Height          =   255
         Left            =   4560
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoPedana 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoPedanaReale 
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PESO REALE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PESO PEDANE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.ListBox lstPedane 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   5430
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PEDANE DISPONIBILI"
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
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   40
      Width           =   7695
   End
End
Attribute VB_Name = "frmPesaturaPedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Link_TipoImballo As Long
Private Link_TipoScarto As Long
Private Link_TipoAumentoPeso As Long
Private Link_TipoCaloPeso As Long
Private Link_TipoProdotto_Q As Long

Private Link_Tipo_Oggetto_Lav As Long
Private Link_Tipo_Oggetto_Quad As Long
Private Link_Funzione_Carico As Long
Private Link_Funzione_Scarico As Long

Private Link_Articolo_Neg As Long
Private Link_Articolo_Pos As Long

Private mov As DmtMovim.cMovimentazione
Private rsLav As ADODB.Recordset
Private rsQuad As ADODB.Recordset

Private Sub CDArticoloScarto_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraVendita, IDTipoProdotto, RV_POIDTipoLavorazione "
sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & Me.CDArticoloScarto.KeyFieldID

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.cboUMScarti.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
    Link_TipoProdotto_Q = fnNotNullN(rs!IDTipoProdotto)
    Me.CboTipoLavScarti.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub cmdElabora_Click()
'On Error GoTo ERR_cmdElabora_Click
Dim Testo As String
Const Titolomsg As String = "Elaborazione ripesatura pedana"

If Me.CDArticoloScarto.KeyFieldID = 0 Then
    Testo = "Inserire l'articolo di quadratura"
    MsgBox Testo, vbCritical, Titolomsg
    Exit Sub
End If

If Me.txtPesoPedana.Value = 0 Then
    Testo = "Selezionare le pedane di ripesare"
    MsgBox Testo, vbCritical, Titolomsg
    Exit Sub
End If

If Me.txtPesoPedanaReale.Value = 0 Then
    Testo = "Inserire il peso reale delle pedane"
    MsgBox Testo, vbCritical, Titolomsg
    Exit Sub
End If

If Me.txtPesoPedana.Value = Me.txtPesoPedanaReale.Value Then
    Testo = "Il peso delle pedane con il peso reale coincidono"
    MsgBox Testo, vbCritical, Titolomsg
    Exit Sub
End If
If Me.txtPesoPedanaReale.Value <= Me.txtTaraTotaleImballi.Value Then
    Testo = "Il peso delle pedane reale è uguale o minore al totale della tara degli imballi della merce" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, Titolomsg) = vbNo Then Exit Sub
    
End If
If (Me.txtPesoPedana.Value - Me.txtPesoPedanaReale) > 0 Then
    If Link_TipoProdotto_Q = Link_TipoAumentoPeso Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Il peso reale è minore del peso precedente ma si è scelto un articolo di quadratura non congruente" & vbCrLf
        Testo = Testo & "Continuando con la procedura si potrebbero avere delle incongruenze nella quadratura del conferimento"
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion, Titolomsg) = vbNo Then Exit Sub
    End If
End If

If (Me.txtPesoPedana.Value - Me.txtPesoPedanaReale) < 0 Then
    If ((Link_TipoProdotto_Q = Link_TipoScarto) Or ((Link_TipoProdotto_Q = Link_TipoCaloPeso))) Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Il peso reale è maggiore del peso precedente ma si è scelto un articolo di quadratura non congruente" & vbCrLf
        Testo = Testo & "Continuando con la procedura si potrebbero avere delle incongruenze nella quadratura del conferimento"
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion, Titolomsg) = vbNo Then Exit Sub
    End If
End If

If Me.cboUMScarti.CurrentID = 0 Then
    Testo = "Inserire l'unità di misura del'articolo di quadratura"
    MsgBox Testo, vbCritical, testomsg
    Exit Sub
End If

If GET_CONTROLLO_UM_ARTICOLO_QUAD(Me.cboUMScarti.CurrentID) = False Then
    Testo = "L'unità di misura dell'articolo di quadratura non è congruente con l'operazione"
    MsgBox Testo, vbCritical, testomsg
    Exit Sub
End If

If Me.chkRiportaOrdine.Value = vbChecked Then
    If FrmMain.txtIDOrdine.Value = 0 Then
        Testo = "ATTENZIONE!!!"
        Testo = Testo & "Non è stato selezionato l'ordine da preparare, pertanto è impossibile continuare"
        MsgBox Testo, vbCritical, Titolomsg
        Exit Sub
    End If
End If

ELABORAZIONE_QUADRATURA

If Me.chkRiportaOrdine.Value = vbChecked Then
    If PREZZI_ARTICOLI_DA_ORDINE = 1 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "È stato impostato di prelevare gli importi dall'ordine, pertanto prima di procedere alla conferma dell'ordine eseguire la prezzatura veloce dal comando 'PREZZATURA DA ORDINE'"
        
        MsgBox Testo, vbInformation, "Prezzatura da ordine"
        
    End If
End If

Unload Me

Exit Sub
ERR_cmdElabora_Click:
    MsgBox Err.Description, vbCritical, "cmdElabora_Click"
End Sub

Private Sub Form_Activate()
    GET_PEDANE
    
    If NUMERO_PEDANA_PESATURA = 1 Then
        Me.lstPedane.Selected(0) = True
    End If
End Sub

Private Sub Form_Load()
    INIT_CONTROLLI
End Sub
Private Sub INIT_CONTROLLI()
    ParametroImballo
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroTipoScarto
    ParametroAggiornaTipoLavorazioneDaConf
    
    Link_Tipo_Oggetto_Lav = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    Link_Tipo_Oggetto_Quad = fnGetTipoOggetto("RV_POLavorazioneL")
    Link_Funzione_Carico = GET_FUNZIONE_MAGAZZINO(10, 1)
    Link_Funzione_Scarico = GET_FUNZIONE_MAGAZZINO(10, 2)
    
    With Me.CDArticoloScarto
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto = " & Link_TipoScarto & ") OR (IDTipoProdotto = " & Link_TipoAumentoPeso & ") OR (IDTipoProdotto = " & Link_TipoCaloPeso & "))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
   
    With Me.CboTipoLavScarti
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .Sql = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With
    
    'Unita di misura scarto
    With Me.cboUMScarti
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .Sql = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With

    Me.chkRiportaOrdine.Value = RIPORTA_IN_ORDINE_DA_PESATURA
    
    Link_Articolo_Neg = fnGetParametriMagazzino("IDArtQuadNegDaPesPed")
    Link_Articolo_Pos = fnGetParametriMagazzino("IDArtQuadPosDaPesPed")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsPedana.Close
    Set rsPedana = Nothing
End Sub
Private Sub GET_PEDANE()
On Error GoTo ERR_GET_PEDANE
Dim Unita_Progresso As Double
Me.lstPedane.Clear

If ((rsPedana.EOF) And (rsPedana.BOF)) Then Exit Sub

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NUMERO_PEDANA_PESATURA), 4)

rsPedana.MoveFirst

While Not rsPedana.EOF
    
    Me.lstPedane.AddItem fnNotNull(rsPedana!CodicePedana) & " (" & fnNotNull(rsPedana!CodiceArticolo) & ")"
    Me.lstPedane.ItemData(Me.lstPedane.NewIndex) = fnNotNullN(rsPedana!IDPedana)
    
rsPedana.MoveNext
Wend
Exit Sub
ERR_GET_PEDANE:
    MsgBox Err.Description, vbCritical, "GET_PEDANE"
End Sub

Private Sub lstPedane_Click()
On Error GoTo ERR_lstPedane_Click
    rsPedana.Filter = "IDPedana=" & Me.lstPedane.ItemData(Me.lstPedane.ListIndex)
    
    If Not rsPedana.EOF Then
        rsPedana!Registra = Abs(Me.lstPedane.Selected(Me.lstPedane.ListIndex))
        rsPedana.Update
    End If
    
    rsPedana.Filter = vbNullString
    
    GET_TOTALE_PEDANA
    
Exit Sub
ERR_lstPedane_Click:
    MsgBox Err.Description, vbCritical, "lstPedane_Click"
End Sub
Private Sub GET_TOTALE_PEDANA()
On Error GoTo ERR_GET_TOTALE_PEDANA
rsPedana.Filter = "Registra=1"

Me.txtPesoPedana.Value = 0
Me.txtTaraTotaleImballi.Value = 0

While Not rsPedana.EOF
    Me.txtPesoPedana.Value = Me.txtPesoPedana.Value + fnNotNullN(rsPedana!PesoPedana) + fnNotNullN(rsPedana!PesoMerceLorda)
    'Me.txtPesoPedana.Value = Me.txtPesoPedana.Value + fnNotNullN(rsPedana!PesoPedana) + fnNotNullN(rsPedana!PesoMerceLorda)
    Me.txtTaraTotaleImballi.Value = Me.txtTaraTotaleImballi.Value + rsPedana!TaraTotaleImballi
rsPedana.MoveNext
Wend
rsPedana.Filter = vbNullString
Exit Sub
ERR_GET_TOTALE_PEDANA:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_PEDANA"
End Sub
Private Sub ParametroImballo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoImballo = rs!IDTipoImballo
Else
    Link_TipoImballo = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroTipoCaloPeso()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCaloPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
Else
    Link_TipoCaloPeso = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub ParametroTipoAumentoPeso()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoAumentoPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
Else
    Link_TipoAumentoPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroTipoScarto()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoScarto FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoScarto = fnNotNullN(rs!IDTipoScarto)
Else
    Link_TipoScarto = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function fncTrovaIDFunzione(Gestore As String, Optional Funzione As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE Gestore.Gestore = " & fnNormString(Gestore)
sSQL = sSQL & " AND Funzione = " & fnNormString(Funzione)
sSQL = sSQL & " AND Funzione.IDFunzione >= 10000"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELABORAZIONE_QUADRATURA()
'On Error GoTo ERR_ELABORAZIONE_QUADRATURA
Dim sSQL As String
Dim PesoPedanaNetto As Double
Dim PesoPedanaRealeNetto As Double

Dim rsConf As DmtOleDbLib.adoResultset

PesoPedanaNetto = Me.txtPesoPedana.Value
PesoPedanaRealeNetto = Me.txtPesoPedanaReale.Value

rsPedana.Filter = "Registra=1"

If rsPedana.EOF Then Exit Sub

While Not rsPedana.EOF
    PesoPedanaNetto = PesoPedanaNetto - fnNotNullN(rsPedana!PesoPedana)
    PesoPedanaRealeNetto = PesoPedanaRealeNetto - fnNotNullN(rsPedana!PesoPedana)
rsPedana.MoveNext
Wend

rsPedana.MoveFirst

CREA_RECORDSET_TMP_LAVORAZIONE

While Not rsPedana.EOF

    Me.List1.AddItem "Elaborazione della pedana " & rsPedana!CodicePedana
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents

    sSQL = "SELECT IDRV_POCaricoMerceRighe "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce"
    sSQL = sSQL & " WHERE IDRV_POPedana=" & fnNotNullN(rsPedana!IDPedana)
    sSQL = sSQL & " GROUP BY IDRV_POCaricoMerceRighe"
    
    Set rsConf = CnDMT.OpenResultset(sSQL)
    
    While Not rsConf.EOF
        
        ELABORAZIONE_LAVORAZIONI fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), fnNotNullN(rsPedana!IDPedana), PesoPedanaNetto, PesoPedanaRealeNetto
    
    rsConf.MoveNext
    Wend
    
    rsConf.CloseResultset
    Set rsConf = Nothing


    Me.List1.AddItem "Elaborazione avvenuta con successo"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    DoEvents

rsPedana.MoveNext
Wend



AVVIA_MOVIMENTAZIONE_LAVORAZIONE
AVVIA_MOVIMENTAZIONE_QUADRATURA

Exit Sub
ERR_ELABORAZIONE_QUADRATURA:
    MsgBox Err.Description, vbCritical, "ELABORAZIONE_QUADRATURA"
End Sub
Private Sub ELABORAZIONE_LAVORAZIONI(IDRigaConferimento As Long, IDPedana As Long, PesoPrec As Double, PesoReale As Double)
'On Error GoTo ERR_ELABORAZIONE_LAVORAZIONI
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim QuantitaQuadratura As Double

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDRV_POPedana=" & IDPedana

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

QuantitaQuadratura = 0

While Not rs.EOF
    
    QuantitaQuadratura = QuantitaQuadratura + (rs!PesoLordo - ((PesoReale / PesoPrec) * fnNotNullN(rs!PesoLordo)))
    
    rs!PesoLordo = (PesoReale / PesoPrec) * fnNotNullN(rs!PesoLordo)
    rs!PesoNetto = rs!PesoLordo - rs!Tara
    
    Select Case fnNotNullN(rs!IDUnitaDiMisuraCoop)
        Case 2 'Peso lordo
            rs!Qta_UM = rs!PesoLordo
        Case 3 'Peso netto
            rs!Qta_UM = rs!PesoNetto
        Case 4 'Tara
            rs!Qta_UM = rs!Tara
    End Select
    
    If Me.chkRiportaOrdine.Value = vbChecked Then
        rs!IDOggettoOrdine = FrmMain.txtIDOrdine.Value
        rs!IDCliente = FrmMain.cdCliente.KeyFieldID
        rs!NumeroOrdine = FrmMain.txtNumeroOrdine.Value
        rs!DataOrdine = FrmMain.txtDataOrdine.Text
        rs!NumeroListaPrelievo = FrmMain.txtNListaPrelievo.Value
        rs!IDOggettoOrdinePadre = FrmMain.txtIDOrdinePadre.Value
        
        GET_CONFIGURAZIONE_IMPORTI_ARTICOLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_ARTICOLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
        
        GET_CONFIGURAZIONE_IMPORTI_IMBALLO FrmMain.cdCliente.KeyFieldID, fnNotNullN(rs!IDArticolo), LINK_LISTINO_CLIENTE, LINK_LISTINO_AZIENDA, fnNotNullN(rs!Qta_UM), PREZZI_IMBALLI_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.txtIDOrdinePadre.Value, rs, FrmMain.txtDataOrdine.Text, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria)
        
        rs!MerceInclusoImballo = GET_PREZZO_IMBALLO_INCLUSO_2(fnNotNullN(rs!IDArticolo), FrmMain.txtIDOrdinePadre.Value, PREZZO_INCLUSO_IMBALLO_DA_ORDINE, fnNotNullN(rs!IDRV_POPedana), fnNotNullN(rs!IDImballoVendita), FrmMain.cdCliente.KeyFieldID, fnNotNull(rs!NotaRigaOrdRaggr), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoCategoria))
    
    End If
    
    
    rs.Update

        
    SALVA_LAVORAZIONE_DA_MOVIMENTARE fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    
    
rs.MoveNext
Wend
rs.Close
Set rs = Nothing

INSERIMENTO_QUADRATURA_CONFERIMENTO IDRigaConferimento, QuantitaQuadratura, 0


Exit Sub
ERR_ELABORAZIONE_LAVORAZIONI:
    MsgBox Err.Description, vbCritical, "ELABORAZIONE_LAVORAZIONI"
End Sub
Private Sub GET_CONFIGURAZIONE_IMPORTI_ARTICOLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset, DataOrdine As String, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long

ImportoUnitario = 0

If PrezziDaOrdine = 1 Then
    IDArticoloPadre = IDArticolo
    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
        If GESTIONE_ORDINE_VIVAIO = 0 Then
            If TROVA_PREZZI_NO_IMB = 0 Then
                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            End If
        End If
        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
        If (TROVA_PREZZI_ORD_CAT = 1) Then
            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
        End If
        If (TROVA_PREZZI_ORD_CAL = 1) Then
            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
        End If
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            If GESTIONE_ORDINE_VIVAIO = 0 Then
                If TROVA_PREZZI_NO_IMB = 0 Then
                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
                End If
            End If
            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
            If (TROVA_PREZZI_ORD_CAT = 1) Then
                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
            End If
            If (TROVA_PREZZI_ORD_CAL = 1) Then
                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
            End If
            
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                rstmp!ImportoUnitarioArticolo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                
                rstmp!Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
                rstmp!Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
                ImportoUnitario = rstmp!ImportoUnitarioArticolo
            End If
            
            rs.CloseResultset
            Set rs = Nothing
        End If
    End If
End If

If ImportoUnitario > 0 Then Exit Sub


Set ObjDoc = New DmtDocs.cDocument
Set ObjDoc.Connection = TheApp.Database.Connection
ObjDoc.SetTipoOggetto 2
ObjDoc.IDFunzione = 105
ObjDoc.TablesNames ObjDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
ObjDoc.IDAzienda = TheApp.IDFirm
ObjDoc.IDFiliale = TheApp.Branch
ObjDoc.IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
ObjDoc.IDTipoAnagrafica = 2
ObjDoc.IDUtente = TheApp.IDUser
ObjDoc.DataEmissione = DataOrdine

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal


ObjDoc.ReadDataFromPriceList IDListino
ObjDoc.ReadDataFromDiscountsList

If Quantita = 0 Then
    ObjDoc.Field "Art_quantita_totale", "0,01", sTabellaDettaglioLocal
Else
    ObjDoc.Field "Art_quantita_totale", Quantita, sTabellaDettaglioLocal
End If

rstmp!ImportoUnitarioArticolo = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))

rstmp!Sconto1 = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglioLocal))
rstmp!Sconto2 = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglioLocal))


Set ObjDoc = Nothing
End Sub

Private Sub GET_CONFIGURAZIONE_IMPORTI_IMBALLO(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDOggettoOrdine As Long, rstmp As ADODB.Recordset, DataOrdine As String, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long)
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Riga_Ordine As Long

ImportoUnitario = 0

If PrezziDaOrdine = 1 Then
    Link_Riga_Ordine = 0
    IDArticoloPadre = IDArticolo
    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
        If GESTIONE_ORDINE_VIVAIO = 0 Then
            If TROVA_PREZZI_NO_IMB = 0 Then
                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            End If
        End If
        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
        If (TROVA_PREZZI_ORD_CAT = 1) Then
            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
        End If
        If (TROVA_PREZZI_ORD_CAL = 1) Then
            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
        End If
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            '''''''''''''''''''TROVO IL LINK_RIGA DELL'ORDINE'''''''''''''''''''''''''''
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            If GESTIONE_ORDINE_VIVAIO = 0 Then
                If TROVA_PREZZI_NO_IMB = 0 Then
                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
                End If
            End If
            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
            If (TROVA_PREZZI_ORD_CAT = 1) Then
                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
            End If
            If (TROVA_PREZZI_ORD_CAL = 1) Then
                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
            End If
            
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                Link_Riga_Ordine = fnNotNullN(rs!RV_POLinkRiga)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Link_Riga_Ordine > 0 Then
                sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
                sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
                sSQL = sSQL & " AND RV_POTipoRiga=2 "
                sSQL = sSQL & " AND RV_POLinkRiga=" & Link_Riga_Ordine
                sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo
                
                Set rs = CnDMT.OpenResultset(sSQL)
                
                If Not rs.EOF Then
                    rstmp!ImportoUnitarioImballo = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
                    ImportoUnitario = fnNotNullN(rstmp!ImportoUnitarioImballo)
                End If
                
                rs.CloseResultset
                Set rs = Nothing
            End If
            
        End If
    End If
End If

If ImportoUnitario > 0 Then Exit Sub


Set ObjDoc = New DmtDocs.cDocument
Set ObjDoc.Connection = TheApp.Database.Connection
ObjDoc.SetTipoOggetto 2
ObjDoc.IDFunzione = 105
ObjDoc.TablesNames ObjDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
ObjDoc.IDAzienda = TheApp.IDFirm
ObjDoc.IDFiliale = TheApp.Branch
ObjDoc.IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
ObjDoc.IDTipoAnagrafica = 2
ObjDoc.IDUtente = TheApp.IDUser
ObjDoc.DataEmissione = DataOrdine

ObjDoc.ClearValues

ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglioLocal).NumRetails
ObjDoc.ReadDataFromCliFo IDAnagrafica
ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
ObjDoc.ReadDataFromArticle IDImballo, sTabellaDettaglioLocal
ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
ObjDoc.Field "Doc_data", ObjDoc.DataEmissione, sTabellaTestataLocal


ObjDoc.ReadDataFromPriceList IDListino
ObjDoc.ReadDataFromDiscountsList

If Quantita = 0 Then
    ObjDoc.Field "Art_quantita_totale", "0,01", sTabellaDettaglioLocal
Else
    ObjDoc.Field "Art_quantita_totale", Quantita, sTabellaDettaglioLocal
End If

rstmp!ImportoUnitarioImballo = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))



Set ObjDoc = Nothing
End Sub


Private Function GET_LINK_ATTIVITA_AZIENDA(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda "
sSQL = sSQL & "FROM Filiale "
sSQL = sSQL & "WHERE Filiale.IDFiliale = " & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PREZZO_IMBALLO_INCLUSO_2(IDArticolo As Long, IDOggettoOrdine As Long, PrezziDaOrdine As Long, IDPedana As Long, IDImballo As Long, IDCliente As Long, RaggrOrdine As String, IDCalibro As Long, IDCategoria As Long) As Long
On Error GoTo ERR_GET_PREZZO_IMBALLO_INCLUSO_2
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCli As DmtOleDbLib.adoResultset
Dim IDArticoloPadre As Long
Dim IDTipoPedana As Long
Dim NumeroCombinazioni As Long
Dim Link_Listino_Dest As Long

GET_PREZZO_IMBALLO_INCLUSO_2 = 0

If PrezziDaOrdine = 1 Then
    IDArticoloPadre = IDArticolo 'GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticolo)
    'IDTipoPedana = GET_TIPO_PEDANA(IDPedana)
    
    If IDArticoloPadre > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "SELECT Count(IDValoriOggettoDettaglio) AS NumeroRecord "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
        sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
        sSQL = sSQL & " AND RV_POTipoRiga=1 "
        sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
        'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
        If GESTIONE_ORDINE_VIVAIO = 0 Then
            If TROVA_PREZZI_NO_IMB = 0 Then
                sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
            End If
        End If
        sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
        If (TROVA_PREZZI_ORD_CAT = 1) Then
            sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
        End If
        If (TROVA_PREZZI_ORD_CAL = 1) Then
            sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
        End If
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            NumeroCombinazioni = 0
        Else
            NumeroCombinazioni = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If NumeroCombinazioni = 1 Then
            sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
            sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
            sSQL = sSQL & " AND RV_POTipoRiga=1 "
            sSQL = sSQL & " AND Link_Art_articolo=" & IDArticoloPadre
            'sSQL = sSQL & " AND RV_POIDTipoPedana=" & IDTipoPedana
            If GESTIONE_ORDINE_VIVAIO = 0 Then
                If TROVA_PREZZI_NO_IMB = 0 Then
                    sSQL = sSQL & " AND RV_POIDImballo=" & IDImballo
                End If
            End If
            sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(RaggrOrdine)
            If (TROVA_PREZZI_ORD_CAT = 1) Then
                sSQL = sSQL & " AND RV_POIDTipoCategoria=" & IDCategoria
            End If
            If (TROVA_PREZZI_ORD_CAL = 1) Then
                sSQL = sSQL & " AND RV_POIDCalibro=" & IDCalibro
            End If
            
            Set rs = CnDMT.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rs!RV_POImportoImballoInArticolo)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            Exit Function
        End If
    End If
End If


If GET_PREZZO_IMBALLO_INCLUSO_2 = 0 Then
    sSQL = "SELECT PrezzoInclusoImballo "
    sSQL = sSQL & "FROM RV_POConfigurazioneClienteImb "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDArticoloImballo=" & IDImballo
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        sSQL = "SELECT PrezzoInclusoImballo "
        sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
        sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
        sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
        'sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
        
        Set rsCli = CnDMT.OpenResultset(sSQL)
        
        If rsCli.EOF Then
            GET_PREZZO_IMBALLO_INCLUSO_2 = 0
        Else
            GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rsCli!PrezzoInclusoImballo)
        End If
        
        rsCli.CloseResultset
        Set rsCli = Nothing
        
    Else
        GET_PREZZO_IMBALLO_INCLUSO_2 = fnNotNullN(rs!PrezzoInclusoImballo)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

Exit Function
ERR_GET_PREZZO_IMBALLO_INCLUSO_2:
    GET_PREZZO_IMBALLO_INCLUSO_2 = 0
End Function

Private Sub MOVIMENTAZIONE_RIGA_LAVORAZIONE(IDRigaConferimento As Long, IDAssegnazioneMerce As Long)
On Error GoTo ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE
Dim OLD_Cursor As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsMov As DmtOleDbLib.adoResultset
Dim IDEsercizio As Long
Dim Movimentato As Long

OLD_Cursor = CnDMT.CursorLocation
CnDMT.CursorLocation = adUseClient
Movimentato = 1
Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection

'''''''''''''''''''''''ELIMINAZIONE MOVIMENTI DELLA RIGA DI LAVORAZIONE'''''''''''''''''''''
sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
sSQL = sSQL & " AND IDOggetto=" & IDRigaConferimento
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    CnDMT.BeginTrans
        mov.Delete fnNotNullN(rs!IDMovimento)
    CnDMT.CommitTrans
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Movimentato = 1

sSQL = "SELECT * FROM RV_POIEMovimentazioneLavorazioni "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    
        
    IDEsercizio = fncEsercizio(fnNotNull(rs!DataDocumento))
        
    If GeneraMovimentoDiCarico(fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLottoVendita), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDUnitaDiMisura), _
    fnNotNull(rs!Articolo), fnNotNullN(rs!Qta_UM), fnNotNull(rs!DataDocumento), _
    fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
    fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), fnNotNullN(rs!IDRV_POTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), fnNotNullN(rs!IDRV_POPedana), _
    fnNotNullN(rs!IDRV_POTipoPedana), fnNotNull(rs!CodicePedana), fnNotNullN(rs!PesoPedana), fnNotNullN(rs!IDUnitaDiMisuraConf), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), _
    fnNotNullN(rs!NumeroConferimento), fnNotNull(rs!CodiceLottoConf), fnNotNullN(rs!IDMagazzinoVendita), IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Carico) = False Then Movimentato = 0

    If GeneraMovimentoDiScarico(fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNull(rs!DataDocumento), fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(IDAnagraficaSocio), _
    fnNotNullN(rs!IDArticoloConf), fnNotNull(rs!ArticoloConf), fnNotNullN(rs!IDUnitaDiMisuraDiamanteConf), fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), _
    fnNotNullN(rs!Pezzi), fnNotNullN(rs!IDUnitaDiMisuraConf), Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), IDEsercizio) = False Then Movimentato = 0
    
    If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballoVendita), GET_LINK_UM_ARTICOLO(fnNotNullN(rs!IDImballoVendita)), fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli), _
    IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Carico, fnNotNullN(rs!IDMagazzinoVendita), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento)) = False Then Movimentato = 0
    
    If GeneraMovimentoScaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballoVendita), GET_LINK_UM_ARTICOLO(fnNotNullN(rs!IDImballoVendita)), fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli), _
    IDEsercizio, Link_Tipo_Oggetto_Lav, Link_Funzione_Scarico, fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDAnagraficaSocio), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), fnNotNull(CodiceLottoVendita), fnNotNull(rs!DataDocumento)) = False Then Movimentato = 0
            
End If

rs.CloseResultset
Set rs = Nothing

CnDMT.CursorLocation = OLD_Cursor
Set mov = Nothing

sSQL = "UPDATE RV_POAssegnazioneMerce SET "
sSQL = sSQL & " Movimentato=" & Movimentato
sSQL = sSQL & " WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce
CnDMT.Execute sSQL

Exit Sub
ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
End Sub

Private Function GeneraMovimentoDiCarico(IDAssegnazione As Long, IDRigaConferimento As Long, CodiceLottoVendita As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, DataLavorazione As String, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double, IDUMConferimentoCoop As Long, _
IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoEntrata As String, _
IDMagazzino As Long, IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long) As Boolean

On Error GoTo ERR_GeneraMovimentoDiCarico

mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 1
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_PONumeroColli", Colli
mov.Field "RV_POPesoLordo", PesoLordo
mov.Field "RV_POPesoNetto", PesoNetto
mov.Field "RV_POTara", Tara
mov.Field "RV_POQuantitaPezzi", Pezzi

Select Case IDUMConferimentoCoop
    Case 1
        mov.Field "RV_POQuantitaMovimentata", Colli
    Case 2
        mov.Field "RV_POQuantitaMovimentata", PesoLordo
    Case 3
        mov.Field "RV_POQuantitaMovimentata", PesoNetto
    Case 4
        mov.Field "RV_POQuantitaMovimentata", Tara
    Case 5
        mov.Field "RV_POQuantitaMovimentata", Pezzi
End Select

mov.Field "RV_PODataLavorazione", DataLavorazione
mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
mov.Field "RV_POIDCalibro", IDCalibro
mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf

mov.Field "RV_POIDPedana", IDPedana
mov.Field "RV_POIDTipoPedana", IDTipoPedana
mov.Field "RV_POCodicePedana", CodicePedana
mov.Field "RV_POPesoPedana", PesoPedana

mov.Field "TipoRiga", trcNessuno

'CnDMT.BeginTrans
GeneraMovimentoDiCarico = mov.Insert
'CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoDiCarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function



Private Function GeneraMovimentoDiScarico(IDAssegnazione As Long, DataLavorazione As String, IDRigaConferimento As Long, _
IDAnagraficaSocio As Long, IDArticoloConferito As Long, ArticoloConferito As String, IDUnitaDiMisura As Long, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, IDUnitaDiMosuraConfCoop As Long, _
IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDEsercizio As Long) As Boolean

On Error GoTo ERR_GeneraMovimentoDiScarico


mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.IDMagazzinoUscita = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticoloConferito
mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", ArticoloConferito

Select Case IDUnitaDiMosuraConfCoop
    Case 1
        mov.Field "QuantitaTotale", Colli
    Case 2
        mov.Field "QuantitaTotale", PesoLordo
    Case 3
        mov.Field "QuantitaTotale", PesoNetto
    Case 4
        mov.Field "QuantitaTotale", Tara
    Case 5
        mov.Field "QuantitaTotale", Pezzi
End Select

mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 0
mov.Field "RV_POIDCaricoMerceRighe", 0
mov.Field "RV_POIDAssegnazioneMerce", 0
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", 0
mov.Field "RV_PODataConferimento", ""
mov.Field "RV_PONumeroConferimento", ""
mov.Field "RV_POCodiceLotto", ""
mov.Field "RV_POCodiceLottoCampagna", ""
mov.Field "RV_POCodiceLottoVendita", ""
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0


mov.Field "RV_PODataLavorazione", ""
mov.Field "RV_POIDTipoLavorazione", 0
mov.Field "RV_POIDCalibro", 0
mov.Field "RV_POIDTipoCategoria", 0
mov.Field "RV_POIDTipoLavorazioneConf", 0
mov.Field "RV_POPrezzoMedioConf", 0

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0


mov.Field "TipoRiga", trcNessuno
'CnDMT.BeginTrans
GeneraMovimentoDiScarico = mov.Insert
'CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoDiScarico:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans

End Function


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
Private Function GET_FUNZIONE_MAGAZZINO(IDTipoDocumentoCoop As Long, IDTipoProcesso As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POProcessiDocumentoCoop.IDFunzione "
sSQL = sSQL & "FROM RV_POProcessiDocumentoCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE RV_POSchemaCoop.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDDocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDTipoProcessoCoop=" & IDTipoProcesso

Set rs = CnDMT.OpenResultset(sSQL)


If rs.EOF Then
    Select Case IDTipoProcesso
        Case 1 'Carico
            GET_FUNZIONE_MAGAZZINO = fnGetParametriMagazzino("IDCausale_Carico_Mag_Vendita")
        Case 2 'Scarico
            GET_FUNZIONE_MAGAZZINO = fnGetParametriMagazzino("IDCausale_Scarico_Mag_vendita")
    End Select
Else
    If fnNotNullN(rs!IDFunzione) = 0 Then
        Select Case IDTipoProcesso
            Case 1 'Carico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleCarico
            Case 2 'Scarico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleScarico
        End Select
        
    Else
        GET_FUNZIONE_MAGAZZINO = fnNotNullN(rs!IDFunzione)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Public Function GeneraMovimentoCaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoVendita, DataLavorazione As String) As Boolean

On Error GoTo ERR_GeneraMovimentoCaricoImballo
mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 2
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0

mov.Field "RV_PODataLavorazione", ""
mov.Field "RV_POIDTipoLavorazione", 0
mov.Field "RV_POIDCalibro", 0
mov.Field "RV_POIDTipoCategoria", 0
mov.Field "RV_POIDTipoLavorazioneConf", 0
mov.Field "RV_POPrezzoMedioConf", 0

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

mov.Field "TipoRiga", trcNessuno

'CnDMT.BeginTrans
    GeneraMovimentoCaricoImballo = mov.Insert
'CnDMT.CommitTrans

Exit Function
ERR_GeneraMovimentoCaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
    
End Function

Public Function GeneraMovimentoScaricoImballo(IDRigaConferimento As Long, IDAssegnazione As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
IDEsercizio As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, CodiceLottoVendita, DataLavorazione As String) As Boolean

On Error GoTo ERR_GeneraMovimentoScaricoImballo

mov.DataMovimento = DataLavorazione
mov.FattoreDiConversione = Null

mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
mov.IDFunzione = IDFunzione
mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoUscita = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUMDiamante
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", Articolo
mov.Field "QuantitaTotale", Qta_UM
mov.Field "Importo", 0
mov.Field "DataDocumento", DataLavorazione
mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
mov.Field "RV_POTipoRiga", 2
mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0

mov.Field "RV_PODataLavorazione", ""
mov.Field "RV_POIDTipoLavorazione", 0
mov.Field "RV_POIDCalibro", 0
mov.Field "RV_POIDTipoCategoria", 0
mov.Field "RV_POIDTipoLavorazioneConf", 0
mov.Field "RV_POPrezzoMedioConf", 0

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

mov.Field "TipoRiga", trcNessuno

'CnDMT.BeginTrans
    GeneraMovimentoScaricoImballo = mov.Insert
'CnDMT.CommitTrans
Exit Function
ERR_GeneraMovimentoScaricoImballo:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    CnDMT.RollbackTrans
End Function
Private Function GET_LINK_UM_ARTICOLO(IDArticolo) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraVendita FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_ARTICOLO = 0
Else
    GET_LINK_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisuraVendita)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function fnGetParametriMagazzino(NomeCampo As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE ((IDUtente=" & TheApp.IDUser & ") "
    sSQL = sSQL & "AND (IDFiliale=" & TheApp.Branch & "))"
    
    Set rsEse = CnDMT.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
            fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
        Else
            sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
            sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
            sSQL = sSQL & "AND (IDUtente=0))"
        
            Set rsEse = CnDMT.OpenResultset(sSQL)
        
            If rsEse.EOF = False Then
                If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                    fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
                Else
                    fnGetParametriMagazzino = 0
                End If
            Else
                fnGetParametriMagazzino = 0
            End If
            
        End If
    Else
        sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
        sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
        sSQL = sSQL & "AND (IDUtente=0))"
        
        Set rsEse = CnDMT.OpenResultset(sSQL)
        
        If rsEse.EOF = False Then
            If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                fnGetParametriMagazzino = fnNotNullN(rsEse.adoColumns(NomeCampo).Value)
            Else
                fnGetParametriMagazzino = 0
            End If
        Else
            fnGetParametriMagazzino = 0
        End If
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Sub CREA_RECORDSET_TMP_LAVORAZIONE()
'''''''''''LAVORAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not (rsLav Is Nothing) Then
    If rsLav.State > 0 Then
        rsLav.Close
    End If
    Set rsLav = Nothing
End If
Set rsLav = New ADODB.Recordset
rsLav.CursorLocation = adUseClient

rsLav.Fields.Append "IDRV_POAssegnazioneMerce", adInteger, , adFldIsNullable
rsLav.Fields.Append "IDRV_POCaricoMerceRighe", adInteger, , adFldIsNullable

rsLav.Open , , adOpenKeyset, adLockBatchOptimistic


'''''''''QUADRATURA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not (rsQuad Is Nothing) Then
    If rsQuad.State > 0 Then
        rsQuad.Close
    End If
    Set rsQuad = Nothing
End If
Set rsQuad = New ADODB.Recordset
rsQuad.CursorLocation = adUseClient

rsQuad.Fields.Append "IDRV_POLavorazione", adInteger, , adFldIsNullable
rsQuad.Fields.Append "IDRV_POCaricoMerceRighe", adInteger, , adFldIsNullable

rsQuad.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Private Sub SALVA_LAVORAZIONE_DA_MOVIMENTARE(IDAssegnazioneMerce As Long, IDCaricoMerceRighe As Long)
On Error GoTo ERR_SALVA_LAVORAZIONE_DA_MOVIMENTARE
    rsLav.Filter = "IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce
    
    If rsLav.EOF Then
        rsLav.AddNew
            rsLav!IDRV_POAssegnazioneMerce = IDAssegnazioneMerce
            rsLav!IDRV_POCaricoMerceRighe = IDCaricoMerceRighe
        rsLav.Update
    End If
    
    rsLav.Filter = vbNullString
Exit Sub
ERR_SALVA_LAVORAZIONE_DA_MOVIMENTARE:
    MsgBox Err.Description, vbCritical, "SALVA_LAVORAZIONE_DA_MOVIMENTARE"
End Sub

Private Sub SALVA_QUADRATURA_DA_MOVIMENTARE(IDLavorazioneMerce As Long, IDCaricoMerceRighe As Long)
On Error GoTo ERR_SALVA_QUADRATURA_DA_MOVIMENTARE
    rsQuad.Filter = "IDRV_POLavorazione=" & IDLavorazioneMerce
    
    If rsQuad.EOF Then
        rsQuad.AddNew
            rsQuad!IDRV_POLavorazione = IDLavorazioneMerce
            rsQuad!IDRV_POCaricoMerceRighe = IDCaricoMerceRighe
        rsQuad.Update
    End If
    
    rsQuad.Filter = vbNullString
Exit Sub
ERR_SALVA_QUADRATURA_DA_MOVIMENTARE:
    MsgBox Err.Description, vbCritical, "SALVA_QUADRATURA_DA_MOVIMENTARE"
End Sub
Private Sub INSERIMENTO_QUADRATURA_CONFERIMENTO(IDCaricoMerceRighe As Long, Quantita As Double, IDUMCoop As Long)
On Error GoTo ERR_INSERIMENTO_QUADRATURA_CONFERIMENTO
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_POLavorazione "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDCaricoMerceRighe

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rsNew.AddNew
    rsNew!IDRV_POLavorazione = fnGetNewKey("RV_POLavorazione", "IDRV_POLavorazione")
    rsNew!IDRV_POCaricoMerceRighe = IDCaricoMerceRighe
    rsNew("IDArticolo").Value = Me.CDArticoloScarto.KeyFieldID
    rsNew("CodiceArticolo").Value = Me.CDArticoloScarto.Code
    rsNew("Articolo").Value = Me.CDArticoloScarto.Description
    rsNew("IDUnitaDiMisuraCoop").Value = GET_LINK_UM_COOP(Me.cboUMScarti.CurrentID)
    rsNew("IDUnitaDiMisura").Value = Me.cboUMScarti.CurrentID
    rsNew("Colli").Value = 0
    rsNew("PesoLordo").Value = Quantita
    rsNew("PesoNetto").Value = Quantita
    rsNew("TaraUnitaria").Value = 0
    rsNew("Tara").Value = 0
    rsNew("Pezzi").Value = 0
    rsNew("Qta_UM").Value = Quantita
    rsNew("IDImballoVendita").Value = 0
    rsNew("CodiceImballoVendita").Value = ""
    rsNew("ImballoVendita").Value = ""
    rsNew("DataDocumento").Value = Date
    rsNew("OraLavorazione").Value = GET_ORARIO(Now)
    rsNew("IDTipoLavorazione").Value = GET_LINK_TIPO_LAVORAZIONE(IDCaricoMerceRighe, Me.CboTipoLavScarti.CurrentID)
rsNew.Update

SALVA_QUADRATURA_DA_MOVIMENTARE fnNotNullN(rsNew!IDRV_POLavorazione), IDCaricoMerceRighe

rsNew.Close
Set rs = Nothing

Exit Sub
ERR_INSERIMENTO_QUADRATURA_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "INSERIMENTO_QUADRATURA_CONFERIMENTO"
End Sub

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

Private Sub ParametroAggiornaTipoLavorazioneDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaTipoLavDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_TIPO_LAVORAZIONE = fnNotNullN(rs!AggiornaTipoLavDaConf)
Else
    AGGIORNA_TIPO_LAVORAZIONE = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_LINK_TIPO_LAVORAZIONE(IDCaricoMerceRighe As Long, IDTipoLavorazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoLavorazioneConf As Long

sSQL = "SELECT IDRV_POTipoLavorazione FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDCaricoMerceRighe

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDTipoLavorazioneConf = 0
Else
    IDTipoLavorazioneConf = fnNotNullN(rs!IDRV_POTipoLavorazione)
End If

rs.CloseResultset
Set rs = Nothing

If AGGIORNA_TIPO_LAVORAZIONE = 1 Then
    If IDTipoLavorazioneConf > 0 Then
        GET_LINK_TIPO_LAVORAZIONE = IDTipoLavorazioneConf
    Else
        GET_LINK_TIPO_LAVORAZIONE = IDTipoLavorazione
    End If
Else
    GET_LINK_TIPO_LAVORAZIONE = IDTipoLavorazione
End If


End Function
Private Function GET_LINK_UM_COOP(Link_UMAcq As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF = False Then
    GET_LINK_UM_COOP = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
Else
    GET_LINK_UM_COOP = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function AVVIA_MOVIMENTAZIONE_LAVORAZIONE()
'On Error GoTo ERR_AVVIA_MOVIMENTAZIONE_LAVORAZIONE
If ((rsLav.EOF) And (rsLav.BOF)) Then Exit Function

rsLav.MoveFirst


While Not rsLav.EOF
    MOVIMENTAZIONE_RIGA_LAVORAZIONE fnNotNullN(rsLav!IDRV_POCaricoMerceRighe), fnNotNullN(rsLav!IDRV_POAssegnazioneMerce)
    
rsLav.MoveNext
Wend

rsLav.Close
Set rsLav = Nothing
Exit Function
ERR_AVVIA_MOVIMENTAZIONE_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "AVVIA_MOVIMENTAZIONE_LAVORAZIONE"
End Function
Private Function AVVIA_MOVIMENTAZIONE_QUADRATURA()
'On Error GoTo ERR_AVVIA_MOVIMENTAZIONE_QUADRATURA
If ((rsQuad.EOF) And (rsQuad.BOF)) Then Exit Function

rsQuad.MoveFirst

While Not rsQuad.EOF
    MOVIMENTAZIONE_RIGA_QUADRATURA fnNotNullN(rsQuad!IDRV_POLavorazione), fnNotNullN(rsQuad!IDRV_POCaricoMerceRighe)
rsQuad.MoveNext
Wend

rsQuad.Close
Set rsQuad = Nothing
Exit Function
ERR_AVVIA_MOVIMENTAZIONE_QUADRATURA:
    MsgBox Err.Description, vbCritical, "AVVIA_MOVIMENTAZIONE_QUADRATURA"
End Function

Private Function MOVIMENTAZIONE_RIGA_QUADRATURA(IDLavorazione As Long, IDRigaConferimento As Long) As String
Dim OLD_Cursor As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim IDTipoProdotto As Long
Dim IDEsercizio As Long


Set mov = New DmtMovim.cMovimentazione

Set mov.Connection = TheApp.Database.Connection


sSQL = "SELECT * FROM RV_POIEMovimentazioneQuadratura "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & IDLavorazione
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    IDTipoProdotto = GET_TIPO_PRODOTTO(fnNotNullN(rs!IDArticolo))
    IDEsercizio = fncEsercizio(fnNotNull(rs!DataDocumento))
    
    If GeneraMovimentoDiCarico_Q(fnNotNullN(rs!IDRV_POLavorazione), IDTipoProdotto, fnNotNullN(rs!Qta_UM), fnNotNull(rs!DataDocumento), fnNotNull(rs!CodiceArticolo), fnNotNull(rs!Articolo), fnNotNullN(rs!IDArticolo), 0, fnNotNullN(rs!IDUnitaDiMisura), _
    fnNotNull(rs!DataDocumento), fnNotNullN(rs!IDTipoLavorazione), 0, 0, fnNotNullN(rs!IDRV_POTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
    IDEsercizio, Link_Tipo_Oggetto_Quad, IDRigaConferimento, fnNotNullN(rs!IDMagazzinoVendita), fnNotNullN(rs!IDAnagrafica), fnNotNull(rs!DataConferimento), fnNotNullN(rs!NumeroConferimento), _
    fnNotNull(rs!CodiceLottoConf), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDUnitaDiMisuraConf), fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!Tara), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Pezzi)) = False Then
        MOVIMENTAZIONE_RIGA_QUADRATURA = MOVIMENTAZIONE_RIGA_QUADRATURA & "Problema riscontrato con la movimentazione della riga di quadratura" & vbCrLf
    End If
    
    If GeneraMovimentoDiScarico_Q(fnNotNullN(rs!IDRV_POLavorazione), IDTipoProdotto, 0, fnNotNull(rs!DataDocumento), _
    IDEsercizio, Link_Tipo_Oggetto_Quad, fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDMagazzinoConferimento), _
    fnNotNullN(rs!IDAnagrafica), fnNotNullN(rs!IDArticoloConf), fnNotNull(rs!ArticoloConf), fnNotNullN(rs!IDUnitaDiMisuraConf), _
    fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!Tara), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Pezzi)) = False Then
        MOVIMENTAZIONE_RIGA_QUADRATURA = MOVIMENTAZIONE_RIGA_QUADRATURA & "Problema riscontrato con la movimentazione della riga di conferimento" & vbCrLf
    End If

rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
Set mov = Nothing

End Function

Private Function GeneraMovimentoDiCarico_Q(IDRigaQuadratura As Long, IDTipoProdotto As Long, Quantita As Double, DataQuadratura As String, CodiceArticolo As String, DescrizioneArticolo As String, IDArticolo As Long, QuantitaMovimentata As Double, IDUnitaDiMisura As Long, _
DataLavorazione As String, IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDEsercizio As Long, IDTipoOggetto As Long, IDOggetto As Long, IDMagazzino As Long, IDAnagraficaSocio As Long, DataConferimento As String, NumeroConferimento As Long, _
CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDUnitaDiMisuraConfCoop As Long, Colli As Double, PesoLordo As Double, Tara As Double, PesoNetto As Double, Pezzi As Double) As Boolean

On Error GoTo ERR_GeneraMovimentoDiCarico_Q

mov.DataMovimento = DataQuadratura
mov.FattoreDiConversione = Null
mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDOggetto

Select Case IDTipoProdotto
    Case Link_TipoCaloPeso
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleCaloPesoCarico")
        mov.Field "Oggetto", "Calo peso"
    Case Link_TipoAumentoPeso
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleAumentoPesoCarico")
        mov.Field "Oggetto", "Aumento di peso"
    Case Link_TipoScarto
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleScartoCarico")
        mov.Field "Oggetto", "Scarto"
End Select

mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoUscita = IDMagazzino
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagraficaSocio
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticolo
mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", DescrizioneArticolo
mov.Field "QuantitaTotale", Quantita
mov.Field "Importo", 0
mov.Field "DataDocumento", DataQuadratura

mov.Field "IDTipoMovimento", 1
mov.Field "TipoRiga", trcNessuno

'DATI DI CONFERIMENTO''''''''''''''''''''''''''''''''''''''''''''''''''
mov.Field "IDValoriOggettoDettaglio", IDRigaQuadratura
mov.Field "RV_POTipoRiga", 1
mov.Field "RV_POIDCaricoMerceRighe", IDOggetto
mov.Field "RV_POIDAssegnazioneMerce", 0
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
mov.Field "RV_PODataConferimento", DataConferimento
mov.Field "RV_PONumeroConferimento", NumeroConferimento
mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
mov.Field "RV_POCodiceLottoVendita", ""
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0

Select Case IDUnitaDiMisuraConfCoop
    Case 1
        mov.Field "RV_POQuantitaMovimentata", Colli
    Case 2
        mov.Field "RV_POQuantitaMovimentata", PesoLordo
    Case 3
        mov.Field "RV_POQuantitaMovimentata", PesoNetto
    Case 4
        mov.Field "RV_POQuantitaMovimentata", Tara
    Case 5
        mov.Field "RV_POQuantitaMovimentata", Pezzi
End Select

mov.Field "RV_PODataLavorazione", DataLavorazione
mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
mov.Field "RV_POIDCalibro", IDCalibro
mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

GeneraMovimentoDiCarico_Q = mov.Insert

Exit Function
ERR_GeneraMovimentoDiCarico_Q:
    MsgBox Err.Description, vbCritical, "GeneraMovimentoDiCarico_Q"
    GeneraMovimentoDiCarico_Q = False

End Function
Private Function GeneraMovimentoDiScarico_Q(IDRigaQuandratura As Long, IDTipoProdotto As Long, Quantita As Double, DataQuadratura As String, _
IDEsercizio As Long, IDTipoOggetto As Long, IDRigaConferimento As Long, IDMagazzino As Long, IDAnagrafica As Long, IDArticoloConferito As Long, _
ArticoloConferito As String, IDUnitaDiMisuraConfCoop As Long, Colli As Double, PesoLordo As Double, Tara As Double, PesoNetto As Double, Pezzi As Double) As Boolean

On errore GoTo ERR_GeneraMovimentoDiScarico_Q

mov.DataMovimento = DataQuadratura
mov.FattoreDiConversione = Null
mov.GestioneMatricole = False
mov.IDEsercizio = IDEsercizio
mov.IDTipoOggetto = IDTipoOggetto
mov.IDOggetto = IDRigaConferimento
Select Case IDTipoProdotto
    Case Link_TipoCaloPeso
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleCaloPeso")
        mov.Field "Oggetto", "Calo peso"
    Case Link_TipoAumentoPeso
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleAumentoPeso")
        mov.Field "Oggetto", "Aumento di peso"
    Case Link_TipoScarto
        mov.IDFunzione = GET_CAUSALE_QUADRATURA("IDCausaleScarto")
        mov.Field "Oggetto", "Scarto"
End Select

mov.IDUtente = TheApp.IDUser
mov.IDMagazzinoUscita = IDMagazzino
mov.IDMagazzinoEntrata = IDMagazzino
mov.Cessione = 0
mov.Field "IDAzienda", TheApp.IDFirm
mov.Field "IDAnagrafica", IDAnagrafica
mov.Field "IDTipoAnagrafica", 3
mov.Field "IDArticolo", IDArticoloConferito
mov.Field "IDcambio", Null
mov.Field "DescrizioneArticolo", ArticoloConferito

mov.Field "Importo", 0
mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
mov.Field "DataDocumento", DataQuadratura

Select Case IDUnitaDiMisuraConfCoop
    Case 1
        mov.Field "QuantitaTotale", Colli
    Case 2
        mov.Field "QuantitaTotale", PesoLordo
    Case 3
        mov.Field "QuantitaTotale", PesoNetto
    Case 4
        mov.Field "QuantitaTotale", Tara
    Case 5
        mov.Field "QuantitaTotale", Pezzi
End Select


mov.Field "IDTipoMovimento", 1
mov.Field "TipoRiga", trcNessuno


'DATI DI CONFERIMENTO
mov.Field "IDValoriOggettoDettaglio", IDRigaQuandratura
mov.Field "RV_POTipoRiga", 1
mov.Field "RV_POIDCaricoMerceRighe", 0
mov.Field "RV_POIDAssegnazioneMerce", 0
mov.Field "RV_POIDProcessoIVGamma", 0
mov.Field "RV_POIDAnagraficaSocio", ""
mov.Field "RV_PODataConferimento", ""
mov.Field "RV_PONumeroConferimento", 0
mov.Field "RV_POCodiceLotto", ""
mov.Field "RV_POCodiceLottoCampagna", ""
mov.Field "RV_POCodiceLottoVendita", ""
mov.Field "RV_POQuantitaLiquidazione", 0
mov.Field "RV_POImportoInclusoImballo", 0
mov.Field "RV_POImportoLiquidazione", 0
mov.Field "RV_POQuantitaMovimentata", 0
mov.Field "RV_PONumeroColli", 0
mov.Field "RV_POPesoLordo", 0
mov.Field "RV_POPesoNetto", 0
mov.Field "RV_POTara", 0
mov.Field "RV_POQuantitaPezzi", 0

mov.Field "RV_PODataLavorazione", DataLavorazione
mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
mov.Field "RV_POIDCalibro", IDCalibro
mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf

mov.Field "RV_POIDPedana", 0
mov.Field "RV_POIDTipoPedana", 0
mov.Field "RV_POCodicePedana", ""
mov.Field "RV_POPesoPedana", 0

GeneraMovimentoDiScarico_Q = mov.Insert


Exit Function
ERR_GeneraMovimentoDiScarico_Q:
    MsgBox Err.Description, vbCritical, "GeneraMovimentoDiScarico_Q"
    GeneraMovimentoDiScarico_Q = False
End Function


Private Function GET_CAUSALE_QUADRATURA(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_CAUSALE_QUADRATURA = fnNotNullN(rs.adoColumns(NomeCampo).Value)
Else
    GET_CAUSALE_QUADRATURA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_PRODOTTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PRODOTTO = 0
Else
    GET_TIPO_PRODOTTO = fnNotNullN(rs!IDTipoProdotto)
End If



rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CONTROLLO_UM_ARTICOLO_QUAD(IDUM As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_UM_ARTICOLO_QUAD = False

sSQL = "SELECT * FROM UnitaDiMisura "
sSQL = sSQL & "WHERE RV_POIDUnitaDiMisuraCoop=3"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If GET_CONTROLLO_UM_ARTICOLO_QUAD = False Then
        If fnNotNullN(rs!IDUnitaDiMisura) = IDUM Then
            GET_CONTROLLO_UM_ARTICOLO_QUAD = True
        End If
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtCodicePedana_Change()
    If Len(Me.txtCodicePedana.Text) = 0 Then
        rsPedana.Filter = vbNullString
        Exit Sub
    End If
    
    rsPedana.Filter = "CodicePedana LIKE " & fnNormString("*" & Me.txtCodicePedana.Text & "*")
    
    GET_PEDANE
    
End Sub

Private Sub txtPesoPedanaReale_Change()
    If Me.txtPesoPedana.Value = 0 Then Exit Sub
    
    If Me.txtPesoPedanaReale.Value = 0 Then
        Me.CDArticoloScarto.Load 0
        Exit Sub
    End If
    
    
    If (Me.txtPesoPedanaReale.Value - Me.txtPesoPedana.Value) < 0 Then
        If Link_Articolo_Neg > 0 Then
            Me.CDArticoloScarto.Load Link_Articolo_Neg
        End If
    End If
    
    If (Me.txtPesoPedanaReale.Value - Me.txtPesoPedana.Value) > 0 Then
        If Link_Articolo_Pos > 0 Then
            Me.CDArticoloScarto.Load Link_Articolo_Pos
        End If
    End If
    
End Sub

