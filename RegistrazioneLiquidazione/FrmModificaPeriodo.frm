VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.0#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmModificaPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creazione liquidazione (2 di 4)"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminaLiquidazione 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   6000
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modifica periodo di liquidazione"
      Height          =   4815
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin DMTDataCmb.DMTCombo cboTipoLiquidazione 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Non calcolare le trattenute su articoli di quadratura"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   6255
      End
      Begin DMTDataCmb.DMTCombo cboPeriodo 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroLiquidazione 
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   1800
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataInizio 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataFine 
         Height          =   315
         Left            =   4800
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DmtSearchAccount2.DmtSearchACS2 ACS 
         Height          =   600
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   1058
         WidthDescription=   4000
         WidthSecondDescription=   2000
         Object.Visible         =   0   'False
         VisibleCode     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SearchType      =   2
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
         CaptionDescription=   "Socio/Fornitore"
         CaptionCode     =   "Codice:"
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboTipoImportoArticolo 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboTipoQuantita 
         Height          =   315
         Left            =   3360
         TabIndex        =   17
         Top             =   3600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboTipoImportoDocumento 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroProtocolloInterno 
         Height          =   315
         Left            =   5160
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Numero di prot. interno"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a valore su:"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   21
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a % su:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data fine"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data inizio"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero di liquidazione"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   0
      Picture         =   "FrmModificaPeriodo.frx":0000
      ScaleHeight     =   4755
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "FrmModificaPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboPeriodo_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT RV_POLiquidazionePeriodo.IDRV_POLiquidazionePeriodo, RV_POLiquidazionePeriodo.Periodo, RV_POLiquidazionePeriodo.NumeroLiquidazione, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.DataInizio, RV_POLiquidazionePeriodo.DataFine, RV_POLiquidazionePeriodo.IDTipoImportoArticolo,"
sSQL = sSQL & "RV_POLiquidazionePeriodo.IDSocio , Anagrafica.Anagrafica, Anagrafica.Nome, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.IDTipoImportoDocumento, RV_POLiquidazionePeriodo.IDTipoQuantita, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.ArticoliDiQuadratura, RV_POLiquidazionePeriodo.IDTipoLiquidazione, RV_POLiquidazionePeriodo.NumeroProtInt "
sSQL = sSQL & "FROM RV_POLiquidazionePeriodo LEFT OUTER JOIN "
sSQL = sSQL & "Anagrafica ON RV_POLiquidazionePeriodo.IDSocio = Anagrafica.IDAnagrafica "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection


If rs.EOF = False Then
    Me.txtDataInizio.Text = fnNotNull(rs!DataInizio)
    Me.txtDataFine.Text = fnNotNull(rs!DataFine)
    Me.txtNumeroLiquidazione.Value = fnNotNullN(rs!NumeroLiquidazione)
    Me.cboTipoImportoArticolo.WriteOn fnNotNullN(rs!IDTipoImportoArticolo)
    Me.cboTipoImportoDocumento.WriteOn fnNotNullN(rs!IDTipoImportoDocumento)
    Me.cboTipoQuantita.WriteOn fnNotNullN(rs!IDTipoQuantita)
    Me.Check1.Value = fnNormBoolean(fnNotNullN(rs!ArticoliDiQuadratura))
    Me.cboTipoLiquidazione.WriteOn fnNotNullN(rs!IDTipoLiquidazione)
    Me.txtNumeroProtocolloInterno.Value = fnNotNullN(rs!NumeroProtInt)
    Me.ACS.Description = fnNotNull(rs!Anagrafica)
    Me.ACS.SecondDescription = fnNotNull(rs!Nome)
    Me.ACS.IDAnagrafica = fnNotNullN(rs!IDSocio)
Else
    Me.txtDataInizio.Text = Date
    Me.txtDataFine.Text = Date
    Me.ACS.IDAnagrafica = 0
    Me.ACS.Description = ""
    Me.ACS.SecondDescription = ""
    Me.txtNumeroLiquidazione.Value = 0
    Me.cboTipoImportoArticolo.WriteOn 0
    Me.cboTipoImportoDocumento.WriteOn 0
    Me.cboTipoQuantita.WriteOn 0
    Me.Check1.Value = Unchecked
    Me.cboTipoLiquidazione.WriteOn 0
    Me.txtNumeroProtocolloInterno.Value = 0
End If

rs.Close
Set rs = Nothing

InitVariabili
End Sub



Private Sub cboTipoQuantita_Click()
If Me.cboTipoQuantita.CurrentID = 3 Then
    Me.Check1.Value = Checked
    'Me.Check1.Enabled = False
Else
    Me.Check1.Value = Unchecked
    'Me.Check1.Enabled = True
End If
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della liquidazione?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdAvanti_Click()
Dim Stringa As String
Dim sSQL As String

Stringa = Permesso

If Stringa = "" Then
    InitVariabili
    
        Nuova_Liquidazione = 0
        CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneTestaConf WHERE IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
        CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneRigheConf WHERE IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
        CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneLavorazione WHERE IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
        CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneVendita WHERE IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
        
        If Me.cboTipoLiquidazione.CurrentID = 1 Then
            EsecuzioneElaborazione Me.ProgressBar1
            Unload Me
        Else
            ElaborazioneLiquidazioneSulVenduto Me.ProgressBar1
            Unload Me
        End If
    
Else
    MsgBox Stringa, vbInformation, "Creazione documenti"
End If
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub



Private Sub cmdEliminaLiquidazione_Click()
Dim TESTO As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


TESTO = "ATTENZIONE!!!" & vbCrLf
TESTO = TESTO & "Con questo comando si eliminano tutte le liquidazioni riferite al periodo in questione" & vbCrLf
TESTO = TESTO & "Continuare?"

If MsgBox(TESTO, vbQuestion + vbYesNo, "Eliminazione periodo di liquidazione") = vbYes Then
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = 1000

'''''''''''''''''''''''ELIMINAZIONE RIGHE AGGIUNTIVE'''''''''''''''''''''''
    sSQL = "SELECT IDRV_POLiquidazione FROM RV_POLiquidazione "
    sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        CnDMT.Execute "DELETE FROM RV_POLiquidazioneRighe WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''ELIMINAZIONE RIGHE ELABORATE''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazioneRigheEla WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''ELIMINAZIONE TESTA''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazione WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''ELIMINAZIONE PERIODO''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazionePeriodo WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Unload Me
End If

End Sub
Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fncTipoImportoArticolo
    fncSocio
    fncPeriodo
    fncTipoLiquidazione
    fncTipoImportoDocumento
    fncTipoQuantita
    'fncTipoImportoDaLiquidare
End Sub
Private Sub fncTipoImportoArticolo()
    With Me.cboTipoImportoArticolo
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoArticolo"
        .DisplayField = "TipoImportoArticolo"
        .Sql = "SELECT * FROM RV_POTipoImportoArticolo WHERE IDRV_POTipoImportoArticolo<>2"
        .Fill
    End With
  
End Sub

Private Sub fncTipoImportoDocumento()
    With Me.cboTipoImportoDocumento
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoDocumento"
        .DisplayField = "TipoImportoDocumento"
        .Sql = "SELECT * FROM RV_POTipoImportoDocumento"
        .Fill
    End With
  
End Sub
Private Sub fncTipoQuantita()
    With Me.cboTipoQuantita
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoQuantita"
        .DisplayField = "TipoQuantita"
        .Sql = "SELECT * FROM RV_POTipoQuantita"
        .Fill
    End With
  
End Sub



Private Sub fncSocio()
    'Socio
    With Me.ACS
        'Imposta la connessione attiva al controllo
        Set .Connection = CnDMT
        'Imposta il nome dell'applicazione
        '.ApplicationName = m_App.
        'Imposta il nome dell'eseguibile dell'applicazione
        .Client = App.EXEName
        'Imposta l'identificativo dell'azienda corrente
        .IDFirm = TheApp.IDFirm
        'Imposta l'identificativo dell'utente corrente
        .IDUser = TheApp.IDUser
        .UserName = "Amministratore"
        'Impostare con la proprietà Hwnd del form che contiene
        'il controllo. Serve per l'esegui gestione
        .HwndContainer = Me.hWnd
    End With


End Sub
Private Sub fncPeriodo()
    With Me.cboPeriodo
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POLiquidazionePeriodo"
        .DisplayField = "Periodo"
        .Sql = "SELECT * FROM RV_POLiquidazionePeriodo ORDER BY NumeroLiquidazione DESC"
        .Fill
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        FrmVisualizzaLiquidazione.Show
        Exit Sub
    End If
    
    If Me.cmdIndietro.Value = True Then
        FrmMain.Show
        Exit Sub
    End If
    If Me.cmdEliminaLiquidazione.Value = True Then
        FrmMain.Show
    End If
End Sub
Private Function Permesso() As String
Permesso = ""

If Me.cboPeriodo.CurrentID = 0 Then
    Permesso = "Inserire il periodo di elaborazione del documento di liquidazione" & vbCrLf
    Exit Function
End If

If Me.txtDataInizio.Text = "" Then
    Permesso = "Inserire la data di inizio elaborazione del documento di liquidazione" & vbCrLf
    Exit Function
End If

If Me.txtDataFine.Text = "" Then
    Permesso = "Inserire la data di fine elaborazione del documento di liquidazione" & vbCrLf
    Exit Function
End If

If DateDiff("d", Me.txtDataInizio.Text, Me.txtDataFine.Text) < 0 Then
    Permesso = "Risulta una incongruenza tra le date di eleborazione" & vbCrLf
    Exit Function
End If

If Me.cboTipoImportoArticolo.CurrentID = 0 Then
    Permesso = "Manca il tipo di importo dell'articolo da utilizzare per il calcolo delle trattenute"
    Exit Function
End If

If Me.cboTipoImportoDocumento.CurrentID = 0 Then
    Permesso = "Manca il tipo di importo totale del documento da utilizzare per il calcolo delle trattenute"
    Exit Function
End If

If Me.cboTipoQuantita.CurrentID = 0 Then
    Permesso = "Manca il tipo di quantita da utilizzare per il calcolo delle trattenute"
    Exit Function
End If

If Me.cboTipoLiquidazione.CurrentID = 0 Then
    Permesso = "Manca il tipo di liquidazione da utilizzare per il calcolo delle trattenute"
    Exit Function
End If


End Function
Private Sub InitVariabili()
    DATA_INIZIO = Me.txtDataInizio.Text
    DATA_FINE = Me.txtDataFine.Text
    NUMERO_LIQUIDAZIONE = Me.txtNumeroLiquidazione.Value
    LINK_SOCIO = Me.ACS.IDAnagrafica
    TIPO_IMPORTO_ARTICOLO = Me.cboTipoImportoArticolo.CurrentID
    TIPO_IMPORTO_DOCUMENTO = Me.cboTipoImportoDocumento.CurrentID
    TIPO_QUANTITA = Me.cboTipoQuantita.CurrentID
    ARTICOLI_DI_QUAD = Me.Check1.Value
    TIPO_LIQUIDAZIONE = Me.cboTipoLiquidazione.CurrentID
    LINK_PERIODO = Me.cboPeriodo.CurrentID
    NUMERO_PROTOCOLLO = Me.txtNumeroProtocolloInterno.Value
End Sub
Private Sub fncTipoLiquidazione()
    With Me.cboTipoLiquidazione
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoLiquidazione"
        .DisplayField = "TipoLiquidazione"
        .Sql = "SELECT * FROM RV_POTipoLiquidazione"
        .Fill
    End With
  
End Sub

