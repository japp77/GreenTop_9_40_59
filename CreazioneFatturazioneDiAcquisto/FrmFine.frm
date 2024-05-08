VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmFine 
   Caption         =   "Passaggio in fattura di acquisto della liquidazione (Passo 2 di 2)"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "FrmFine.frx":4781A
      ScaleHeight     =   4785
      ScaleWidth      =   2745
      TabIndex        =   18
      Top             =   0
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo di documento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Width           =   3735
      Begin VB.CheckBox chkRaggrSocio 
         Caption         =   "Raggruppa per socio/fornitore"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   3495
      End
      Begin VB.CheckBox chkRiportaRiferimentoLiq 
         Caption         =   "Riporta riferimento liquidazione"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         Width           =   3495
      End
      Begin VB.CheckBox chkLordoLiquidazione 
         Caption         =   "Liquidazione inclusa I.V.A."
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   3375
      End
      Begin DMTDataCmb.DMTCombo cboCausaleContabile 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   3495
         _ExtentX        =   6165
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
      Begin DMTDataCmb.DMTCombo cboEsercizio 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   3495
         _ExtentX        =   6165
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
      Begin DMTDataCmb.DMTCombo cboAliquotaIVA 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   3495
         _ExtentX        =   6165
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
      Begin DMTDataCmb.DMTCombo CboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
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
      Begin DMTDATETIMELib.dmtDate txtDataDoc 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo CboValuta 
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
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
      Begin DMTDataCmb.DMTCombo CboSezionale 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label Label1 
         Caption         =   "Causale contabile"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Esercizio"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Aliquota I.V.A. "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Valuta"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Data documento"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Pagamento di Default"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3495
      End
   End
   Begin VB.CommandButton CmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton CmdAvvia 
      Caption         =   "Fine"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   5040
      TabIndex        =   0
      Top             =   5520
      Width           =   3735
      Begin DMTDataCmb.DMTCombo cboReport 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.CheckBox chkStampaFattura 
         Alignment       =   1  'Right Justify
         Caption         =   "Stampa fattura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Report"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Numero di copie"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
   End
   Begin DMTPrinterDialog.DMTDialog DMTDialog 
      Left            =   120
      Top             =   4920
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   7320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label LblFine 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7800
      Width           =   8535
   End
End
Attribute VB_Name = "FrmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RigaFatturazioneIniziale As String
Private RigaFatturazioneFinale As String
Private RigaLiquidazione As String
Private Link_PDC As Long
Private Link_PDC_Righe_Positive As Long
Private Link_PDC_Righe_Negative As Long
Private Link_PianoDeiConti As Long
'Numero documento della fattura per conto del socio
Private NumeroDocumento As Long

Private Const Link_TipoOggettoPerDoc = 301
Private Link_TipoOggetto As Long

Public Cls_Nom As Collection
Private ObjDoc As DmtDocs.cDocument
Private cDefault As Collection
'Private cPerfoming As New CPerforming
Private Const NOMETABELLAPIANA = "ValoriOggettoPerTipo"
Private Const NOMETABELLADETTAGLIO = "ValoriOggettoDettaglio"

Private ArrayCli(0, 12) As String

Private NumeroRecord As Long

Private oReport As dmtReportLib.dmtReport
Private rsLiquidazioni As ADODB.Recordset
Private rsLiquidazioniInFattura As ADODB.Recordset

Private IRigaDoc As Integer


Private Sub DataDoc_Change()
    fncEsercizio Me.txtDataDoc.Text
End Sub

Private Sub CmdAvvia_Click()
On Error GoTo ERR_CmdAvvia_Click
Dim rsLiq As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim I As Integer
Dim Unita_Progresso As Double
Dim NumeroElaborazioni As Long

If Me.CboSezionale.CurrentID = 0 Then
    MsgBox "Inserire il sezionale del documento", vbCritical, "Inserimento dati"
    Exit Sub
End If

Me.LblFine.Caption = "RECUPERO DATI IN CORSO......"
DoEvents

GET_CICLO_LIQUIDAZIONI

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

If NumeroRecord = 0 Then
    Me.LblFine.Caption = ""
    Exit Sub
End If
Unita_Progresso = Me.ProgressBar1.Max / NumeroRecord

Set ObjDoc = New DmtDocs.cDocument

Link_PianoDeiConti = GetPianoDeiConti

Dim idSocio As Long
Dim IDAnagrafica As Long
Dim IDAnagraficaSocio As Long

If ((rsLiquidazioni.EOF) And (rsLiquidazioni.BOF)) Then
    Me.LblFine.Caption = "OPERAZIONE COMPLETATA"
    Exit Sub
End If
NumeroElaborazioni = 1
If Me.chkRaggrSocio.Value = vbUnchecked Then
    rsLiquidazioni.MoveFirst
    
    'For I = 1 To Cls_Nom.Count
    While Not rsLiquidazioni.EOF
        sSQL = "SELECT * FROM RV_POLiquidazione WHERE IDRV_POLiquidazione=" & fnNotNullN(rsLiquidazioni!IDRiferimento)
        
        Me.LblFine.Caption = "ELABORAZIONE LIQUIDAZIONE " & NumeroElaborazioni & " di " & NumeroRecord
        DoEvents
        
        Set rsLiq = CnDMT.OpenResultset(sSQL)
        
        If rsLiq.EOF = False Then
            Settaggio
            
            If fnNotNullN(rsLiq!idanagraficaFatturazione) = 0 Then
                IDAnagrafica = fnNotNullN(rsLiq!IDAnagrafica)
                IDAnagraficaSocio = fnNotNullN(rsLiq!IDAnagrafica)
            Else
                IDAnagrafica = fnNotNullN(rsLiq!idanagraficaFatturazione)
                IDAnagraficaSocio = fnNotNullN(rsLiq!IDAnagrafica)
            End If
            
            
            fncTestata fnNotNullN(rsLiq!IDAnagrafica), fnNotNullN(rsLiq!idanagraficaFatturazione)
            'fncRighe fnNotNullN(rsLiq!NettoLiquidazione), fnNotNullN(rsLiq!IDAnagrafica), fnNotNullN(rsLiq!IDRV_POLiquidazione), fnNotNullN(rsLiq!IDRV_POLiquidazionePeriodo)
            
            IRigaDoc = 1
            fncRighe fnNotNullN(rsLiq!IDAnagrafica), fnNotNullN(rsLiq!IDRV_POLiquidazione), fnNotNullN(rsLiq!IDRV_POLiquidazionePeriodo), IDAnagraficaSocio
            
            idSocio = fnNotNullN(rsLiq!IDAnagrafica)
            
            rsLiq.CloseResultset
            Set rsLiq = Nothing
            
            InserimentoDMT fnNotNullN(rsLiquidazioni!IDRiferimento), IDAnagrafica
            
            fncAggiornaNumeroFattura IDAnagrafica, ObjDoc.IDEsercizio
            
            SCRIVI_RIFERIMENTI_CONFERIMENTI IDAnagrafica, IDAnagraficaSocio
            
            If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
                Me.ProgressBar1.Value = Me.ProgressBar1.Max
            Else
                Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
            End If
            
        End If
        
        NumeroElaborazioni = NumeroElaborazioni + 1
        
        rsLiquidazioni.MoveNext
    Wend

Else
    While Not rsLiquidazioni.EOF 'Ci sono le anagrafiche dei soci
        
        Me.LblFine.Caption = "ELABORAZIONE FATTURA " & NumeroElaborazioni & " di " & NumeroRecord
        DoEvents

        Settaggio
        
        'RAGGRUPPARE LE LIQUIDAZIONI PER SOCIO E SOCIO DI FATTURAZIONE
        If fnNotNullN(rsLiquidazioni!IDRiferimentoAnaFatt) = 0 Then
            IDAnagrafica = fnNotNullN(rsLiquidazioni!IDRiferimento)
            IDAnagraficaSocio = fnNotNullN(rsLiquidazioni!IDRiferimento)
        Else
            IDAnagrafica = fnNotNullN(rsLiquidazioni!IDRiferimentoAnaFatt)
            IDAnagraficaSocio = fnNotNullN(rsLiquidazioni!IDRiferimento)
        End If
        IDAnagrafica = fnNotNullN(rsLiquidazioni!IDRiferimentoAnaFatt)
        IDAnagraficaSocio = fnNotNullN(rsLiquidazioni!IDRiferimento)
        
        fncTestata IDAnagraficaSocio, IDAnagrafica
        
        sSQL = "SELECT RV_POTMPLiqDaFatt.IDLiquidazione, RV_POLiquidazione.NumeroLiquidazione, RV_POLiquidazione.IDAnagrafica, Anagrafica.Anagrafica, "
        sSQL = sSQL & "RV_POLiquidazione.NettoLiquidazione, RV_POLiquidazione.IDRV_POLiquidazionePeriodo "
        sSQL = sSQL & "FROM RV_POTMPLiqDaFatt INNER JOIN "
        sSQL = sSQL & "RV_POLiquidazione ON RV_POTMPLiqDaFatt.IDLiquidazione = RV_POLiquidazione.IDRV_POLiquidazione INNER JOIN "
        sSQL = sSQL & "Fornitore ON RV_POLiquidazione.IDAnagrafica = Fornitore.IDAnagrafica INNER JOIN "
        sSQL = sSQL & "Anagrafica ON Fornitore.IDAnagrafica = Anagrafica.IDAnagrafica "
        sSQL = sSQL & " WHERE Fornitore.IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND RV_POLiquidazione.IDAnagrafica=" & IDAnagraficaSocio
        sSQL = sSQL & " AND RV_POLiquidazione.IDAnagraficaFatturazione=" & IDAnagrafica
        
        Set rsLiq = CnDMT.OpenResultset(sSQL)
        IRigaDoc = 1
        Set rsLiquidazioniInFattura = New ADODB.Recordset
        rsLiquidazioniInFattura.CursorLocation = adUseClient
        
        rsLiquidazioniInFattura.Fields.Append "IDLiquidazione", adInteger, , adFldIsNullable
        rsLiquidazioniInFattura.Open , , adOpenKeyset, adLockBatchOptimistic
        
        While Not rsLiq.EOF
            
            rsLiquidazioniInFattura.AddNew
            rsLiquidazioniInFattura!IDLiquidazione = fnNotNullN(rsLiq!IDLiquidazione)
            rsLiquidazioniInFattura.Update
            
            fncRighe fnNotNullN(rsLiq!IDAnagrafica), fnNotNullN(rsLiq!IDLiquidazione), fnNotNullN(rsLiq!IDRV_POLiquidazionePeriodo), fnNotNullN(rsLiq!IDAnagrafica)
                
        rsLiq.MoveNext
        Wend

        rsLiq.CloseResultset
        Set rsLiq = Nothing
        
        InserimentoDMT fnNotNullN(rsLiquidazioni!IDRiferimento), IDAnagrafica
        
        fncAggiornaNumeroFattura IDAnagrafica, ObjDoc.IDEsercizio
        
        rsLiquidazioniInFattura.Close
        Set rsLiquidazioniInFattura = Nothing
        
        SCRIVI_RIFERIMENTI_CONFERIMENTI IDAnagrafica, IDAnagraficaSocio
        
        If Me.chkStampaFattura.Value = vbChecked Then
            If Me.cboReport.CurrentID > 0 Then
                ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                StampaDocumento fnNotNullN(ObjDoc.IDOggetto)
            End If
        End If
        
        If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
        End If
    
    NumeroElaborazioni = NumeroElaborazioni + 1
    rsLiquidazioni.MoveNext
    Wend

End If
fncSalvataggioImpostazioniDefaultFiliale
Me.LblFine.Caption = "OPERAZIONE COMPLETATA"
Me.CmdAvvia.Enabled = False

rsLiquidazioni.Close
Set rsLiquidazioni = Nothing

Exit Sub
ERR_CmdAvvia_Click:
    MsgBox Err.Description, vbCritical, "ERR_CmdAvvia_Click"
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    fncIva
    fncSezionale
    fncPagamento
    fncValuta
    fncEsercizioCombo
    fncCausaliContabili
    GET_NUMERAZIONE_COOPERATIVA
    Me.txtDataDoc.Text = Date
    DEFAULT_AZIENDA
    Link_TipoOggetto = fncIDTipoOggettoPrg("RV_POFACS")
    
    fncReport
    
    Me.cboReport.WriteOn fnDefaultReport(Link_TipoOggetto, TheApp.Branch)
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
End Sub
Private Sub fncCausaliContabili()
    Dim sSQL As String

    sSQL = "SELECT IDCausaleContabile, CausaleContabile"
    sSQL = sSQL & " FROM CausaleContabile "
    sSQL = sSQL & "WHERE IDRegistroIva=2 AND IDTipoOggetto=122 "
    sSQL = sSQL & "ORDER BY CausaleContabile"

    With Me.cboCausaleContabile
        Set .Database = CnDMT
        .DisplayField = "CausaleContabile"
        .AddFieldKey "IDCausaleContabile"
        .Sql = sSQL
        .Refresh
    End With

End Sub
Private Sub fncPagamento()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDPagamento, Pagamento"
    sSQL = sSQL & " FROM Pagamento"
    sSQL = sSQL & " ORDER BY Pagamento"

    With Me.CboPagamento
        Set .Database = CnDMT
        .DisplayField = "Pagamento"
        .AddFieldKey "IDPagamento"
        .Sql = sSQL
        .Refresh
    End With

End Sub
Private Sub fncIva()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDIva, Iva"
    sSQL = sSQL & " FROM Iva"
    sSQL = sSQL & " ORDER BY Iva"

    With Me.cboAliquotaIVA
        Set .Database = CnDMT
        .DisplayField = "Iva"
        .AddFieldKey "IDIva"
        .Sql = sSQL
        .Refresh
    End With

End Sub
Private Sub fncEsercizioCombo()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDEsercizio, Esercizio"
    sSQL = sSQL & " FROM Esercizio "
    sSQL = sSQL & "WHERE IDAzienda=" & VarIDAzienda
    sSQL = sSQL & " ORDER BY IDEsercizio"

    With Me.cboEsercizio
        Set .Database = CnDMT
        .DisplayField = "Esercizio"
        .AddFieldKey "IDEsercizio"
        .Sql = sSQL
        .Refresh
    End With

End Sub
Private Sub fncValuta()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As dmtoledblib.adoResultset
    
    
    sSQL = "SELECT IDValuta, Valuta"
    sSQL = sSQL & " FROM Valuta"
    'sSQL = sSQL & " WHERE ((IDFiliale=" & VarIDFiliale & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Valuta"

    With Me.CboValuta
        Set .Database = CnDMT
        .DisplayField = "Valuta"
        .AddFieldKey "IDValuta"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub
Private Sub fncSezionale()
    Dim sSQL As String
    
    sSQL = "SELECT IDSezionale, Sezionale"
    sSQL = sSQL & " FROM Sezionale"
    sSQL = sSQL & " WHERE ((IDFiliale=" & VarIDFiliale & ") AND (IDRegistroIva = 2))"
    sSQL = sSQL & " ORDER BY Sezionale"

    With Me.CboSezionale
        Set .Database = CnDMT
        .DisplayField = "Sezionale"
        .AddFieldKey "IDSezionale"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Me.CmdIndietro.Value = True Then
        FrmMain.Show
    End If
End Sub

Private Sub txtDataDoc_Change()
    fncEsercizio Me.txtDataDoc.Text
End Sub
Private Sub DEFAULT_AZIENDA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & VarIDFiliale
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    RigaFatturazioneIniziale = fnNotNull(rs!RigaFatturazioneIniziale)
    RigaFatturazioneFinale = fnNotNull(rs!RigaFatturazioneFinale)
    Link_PDC = 0
    Me.cboAliquotaIVA.WriteOn 0
    Me.CboPagamento.WriteOn 0
    Me.CboValuta.WriteOn 0
    Me.cboCausaleContabile.WriteOn 0
    Link_PDC_Righe_Positive = 0
    Link_PDC_Righe_Negative = 0
    Me.chkLordoLiquidazione.Value = 0
    Me.chkRiportaRiferimentoLiq.Value = 0
    Me.chkRaggrSocio.Value = 0
Else
    RigaFatturazioneIniziale = fnNotNull(rs!RigaFatturazioneIniziale)
    RigaFatturazioneFinale = fnNotNull(rs!RigaFatturazioneFinale)
    Link_PDC = fnNotNullN(rs!IDPDCFatturazione)
    Me.cboAliquotaIVA.WriteOn fnNotNullN(rs!IDIvaFatturazione)
    Me.CboPagamento.WriteOn fnNotNullN(rs!IDPagamentoFatturazione)
    Me.CboValuta.WriteOn fnNotNullN(rs!IDValutaFatturazione)
    Me.cboCausaleContabile.WriteOn fnNotNullN(rs!IDCausaleContabileFatturazione)
    Link_PDC_Righe_Positive = fnNotNullN(rs!IDPDCRigheLiqPositive)
    Link_PDC_Righe_Negative = fnNotNullN(rs!IDPDCRigheLiqNegative)
    Me.chkLordoLiquidazione.Value = Abs(fnNotNullN(rs!LiquidazioneInclusoIVA))
    Me.chkRiportaRiferimentoLiq.Value = Abs(fnNotNullN(rs!RiportaDescrizioneRifLiq))
    Me.chkRaggrSocio = Abs(fnNotNullN(rs!RaggruppaLiqPerSocio))
End If

rs.CloseResultset
Set rs = Nothing

Me.cboEsercizio.WriteOn VarIDEsercizio
Me.CboSezionale.WriteOn GET_SEZIONALE_PER_DEFAULT
End Sub
Private Function GET_SEZIONALE_PER_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale "
sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE (IDTipoOggetto =" & Link_TipoOggettoPerDoc & ") And (IDFiliale = " & VarIDFiliale & ") And (IDReportTipoOggetto Is Null)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_SEZIONALE_PER_DEFAULT = 0
Else
    GET_SEZIONALE_PER_DEFAULT = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub Settaggio()
Set ObjDoc = New cDocument
    With ObjDoc
        Set .Connection = CnDMT
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = VarIDAttivitaAzienda
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Link_TipoOggettoPerDoc
        .IDFunzione = 269
        .UseAutomation = True
        .IDEsercizio = Me.cboEsercizio.CurrentID
        .IDSezionale = Me.CboSezionale.CurrentID
        .IDTipoAnagrafica = 3
        .IDUtente = VarIDUtente
        .Descrizione = "Fattura di acquisto"
        .DataEmissione = Me.txtDataDoc.Text
        
        .Numero = 0
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto Link_TipoOggettoPerDoc
        Else
            .ClearValues
        End If
        
    End With
End Sub
Private Function fncTestata(idSocio As Long, idanagraficaFatturazione As Long) As Boolean
Dim IDAnagrafica As Long
Dim IDAnagraficaSocio As Long

VARErroreFunzione = "fncTestata"

 With ObjDoc.Tables

'Imposta la riga attiva per la tabella di testata
    
    ObjDoc.Tables(NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
    
    If idanagraficaFatturazione = 0 Then
        IDAnagrafica = idSocio
        IDAnagraficaSocio = idSocio
    Else
        IDAnagrafica = idanagraficaFatturazione
        IDAnagraficaSocio = idSocio
    End If
    
    TrovaAnagrafica IDAnagrafica
    
    ObjDoc.ReadDataFromCliFo IDAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    If fnNotNullN(.Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
        ObjDoc.ReadDataFromPayment fnNotNullN(.Field("Link_Doc_pagamento", FrmFine.CboPagamento.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
    End If
            
    'ObjDoc.ReadDataFromPayment CLng(ArrayCli(0, 11)), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    .Field "Link_Doc_sezionale", FrmFine.CboSezionale.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Val_valuta", FrmFine.CboValuta.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Val_cambio", Null, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_data", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_numero_presso_nom", GET_NUMERO_DOC_FATTURA_ACQ(IDAnagrafica, ObjDoc.IDEsercizio), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_data_presso_nom", Me.txtDataDoc.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_crea_scadenze", ObjDoc.DBDefaults.CreaScadenzeDaDocAcquisto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
End With

fncTestata = True
     
Exit Function
ERR_fncTestata:
    fncTestata = False
    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento
    
End Function
Private Function fncRighe(idSocio As Long, IDLiquidazione As Long, IDPeriodoLiquidazione As Long, idsocioliquidazione As Long) As Boolean
Dim LINK_TIPO_CORPO_FATTURA_SOCIO As Long
'Dim IRigaDoc As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim RsCorpo As ADODB.Recordset
Dim IMPORTO_TRATTENUTE_AGGIUNTIVA As Double
Dim TOTALE_TRATTENUTE As Double
Dim IMPONIBILE_LIQUIDAZIONE As Double
Dim IMPORTO_UNITARIO As Double
Dim TRATTENUTA_AGG_PER_RIGA As Double
Dim ALIQUOTA_IVA As Double
Dim LINK_IVA_ARTICOLO As Long
Dim LINK_IVA_CLIENTE As Long
Dim LINK_CATEGORIA_MERCEOLOGICA As Long
Dim LINK_CONTO_PDC As Long
Dim LINK_CATEGORIA_LIQUIDAZIONE As Long

    
VARErroreFunzione = "fncRighe"
    
    GET_TIPO_CORPO_DOCUMENTO IDPeriodoLiquidazione
    
    LINK_IVA_CLIENTE = 0
    
    LINK_IVA_CLIENTE = fnNotNullN(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))
    
    If fnNotNullN(ObjDoc.Field("Link_Nom_lettera_intento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        LINK_IVA_CLIENTE = GET_LINK_IVA_LETTERA_INTENTO(fnNotNullN(ObjDoc.Field("Link_Nom_lettera_intento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), LINK_IVA_CLIENTE)
    End If
    
    
    'IRigaDoc = 1
'*******************INSERIMENTO RIGA INIZIALE***********************************************************************
    If Len(RigaFatturazioneIniziale) > 0 Then
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaFatturazioneIniziale, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End With
        IRigaDoc = IRigaDoc + 1
    End If
'********************************************************************************************************************


'*******************INSERIMENTO RIGHE DI LIQUIDAZIONE********************************
LINK_TIPO_CORPO_FATTURA_SOCIO = GET_TIPO_CORPO(idsocioliquidazione)

If LINK_TIPO_CORPO_FATTURA_SOCIO = 1 Then 'Fattura riepilogata in un unico importo
    sSQL = "SELECT RV_POLiquidazioneRigheEla.IDRV_POLiquidazione, RV_POLiquidazioneRigheEla.IDRV_POLiquidazionePeriodo, Articolo.IDIvaAcquisto, Iva.Codice, "
    sSQL = sSQL & "Iva.AliquotaIva, Iva.Iva, SUM(RV_POLiquidazioneRigheEla.ImponibileDaReg) AS SommaImponibileDaReg, "
    sSQL = sSQL & "SUM(RV_POLiquidazioneRigheEla.TrattenuteTotali) As SommaTrattenuteTotali, "
    sSQL = sSQL & "Sum(RV_POLiquidazioneRigheEla.ImportoLordoDaReg) AS SommaImportoLordoDaReg "
    sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
    sSQL = sSQL & "Articolo ON Iva.IDIva = Articolo.IDIvaAcquisto RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_POLiquidazioneRigheEla ON Articolo.IDArticolo = RV_POLiquidazioneRigheEla.IDArticolo "
    sSQL = sSQL & "GROUP BY RV_POLiquidazioneRigheEla.IDRV_POLiquidazione, RV_POLiquidazioneRigheEla.IDRV_POLiquidazionePeriodo, Articolo.IDIvaAcquisto, Iva.Codice,"
    sSQL = sSQL & "Iva.AliquotaIva , Iva.Iva "
    sSQL = sSQL & "HAVING (RV_POLiquidazioneRigheEla.IDRV_POLiquidazione = " & IDLiquidazione & ")"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    
    
    While Not rs.EOF
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
            RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, fnNotNullN(rs!AliquotaIva), fnNotNull(rs!Iva), fnNotNullN(rs!Codice))
            
            Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                Case 1
                    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione, (fnNotNullN(rs!SommaImponibileDaReg) - fnNotNullN(rs!SommaTrattenuteTotali)))
                    IMPORTO_UNITARIO = fnNotNullN(rs!SommaImponibileDaReg) - (IMPORTO_TRATTENUTE_AGGIUNTIVA + fnNotNullN(rs!SommaTrattenuteTotali))
                Case 2
                    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione, (fnNotNullN(rs!SommaImportoLordoDaReg) - fnNotNullN(rs!SommaTrattenuteTotali)))
                    IMPORTO_UNITARIO = fnNotNullN(rs!SommaImportoLordoDaReg) - (IMPORTO_TRATTENUTE_AGGIUNTIVA + fnNotNullN(rs!SommaTrattenuteTotali))
                Case Else
                    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione, (fnNotNullN(rs!SommaImponibileDaReg) - fnNotNullN(rs!SommaTrattenuteTotali)))
                    IMPORTO_UNITARIO = fnNotNullN(rs!SommaImponibileDaReg) - (IMPORTO_TRATTENUTE_AGGIUNTIVA + fnNotNullN(rs!SommaTrattenuteTotali))
            End Select
    
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (fnNotNullN(rs!AliquotaIva) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE)
                    ALIQUOTA_IVA = (ALIQUOTA_IVA / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                End If
            End If
            
            
            With ObjDoc.Tables
                .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", fnNotNullN(rs!IDIvaAcquisto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                If Link_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
            IRigaDoc = IRigaDoc + 1
            End With
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
End If

If LINK_TIPO_CORPO_FATTURA_SOCIO = 2 Then 'Fattura dettagliata per articolo e per IRigaDoc.V.A.
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        If (Len(RigaLiquidazione) > 0) Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            With ObjDoc.Tables
                .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            IRigaDoc = IRigaDoc + 1
            End With
        End If
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoArticolo "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    'RsCorpo.Filter = "IDRV_POLiquidazione=" & IDLiquidazione
    
    While Not RsCorpo.EOF
        If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))


            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CONTO_PDC = GET_LINK_CONTO_PDC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_LIQUIDAZIONE = GET_LINK_CATEGORIA_LIQ(RsCorpo!IDArticolo)
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
                
            End If
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                End If
            End If
            
'            If ((IMPORTO_UNITARIO < 0) And (fnNotNullN(RsCorpo!Quantita_Per_Totali) < 0)) Then
'                IMPORTO_UNITARIO = Abs(fnNotNullN(IMPORTO_UNITARIO))
'            End If
            
            With ObjDoc.Tables
                .Field "Art_descrizione", fnNotNull(RsCorpo!Articolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                Else
                    If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                
                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                If LINK_CONTO_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    SetPDCArticolo LINK_CONTO_PDC
                Else
                    If Link_PDC > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        SetPDCArticolo Link_PDC
                    End If
                End If
                
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                
            IRigaDoc = IRigaDoc + 1
            
            
            End With
        End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If
'*************************************************************************************
If LINK_TIPO_CORPO_FATTURA_SOCIO = 3 Then 'Fattura dettagliata per riga di liquidazione con importo al lordo delle trattenute
    
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

        IRigaDoc = IRigaDoc + 1
        End With
    End If
    
    'IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    'TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    
    
    sSQL = "SELECT * FROM RV_POLiquidazioneRigheEla "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    'RsCorpo.Filter = "IDRV_POLiquidazione=" & IDLiquidazione
    
    While Not RsCorpo.EOF
        If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))
            
            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CONTO_PDC = GET_LINK_CONTO_PDC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_LIQUIDAZIONE = GET_LINK_CATEGORIA_LIQ(fnNotNullN(RsCorpo!IDArticolo))
            
            'If TOTALE_TRATTENUTE > 0 Then
            '    TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            'Else
            '    TRATTENUTA_AGG_PER_RIGA = IMPORTO_TRATTENUTE_AGGIUNTIVA
            'End If
            
            'TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            'TRATTENUTA_AGG_PER_RIGA = FormatNumber(TRATTENUTA_AGG_PER_RIGA, 2)
            IMPORTO_UNITARIO = fnNotNullN(RsCorpo!ImportoUnitarioDaReg)
            
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                    
                End If
            End If
            
            With ObjDoc.Tables
                
                .Field "Art_descrizione", fnNotNull(RsCorpo!Articolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                If LINK_CONTO_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    SetPDCArticolo LINK_CONTO_PDC
                Else
                    If Link_PDC > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        SetPDCArticolo Link_PDC
                    End If
                End If
                

                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)


                .Field "RV_POOggettoVendita", fnNotNull(RsCorpo!Oggetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POOggettoConferimento", "Conferimento N° " & fnNotNullN(RsCorpo!NumeroDocumento) & " del " & fnNotNull(RsCorpo!DataConferimento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCaricoMerceRighe", fnNotNullN(RsCorpo!IDRV_POCaricoMerceRighe), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
            IRigaDoc = IRigaDoc + 1
            End With
        End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If

If LINK_TIPO_CORPO_FATTURA_SOCIO = 4 Then 'Fattura dettagliata per riga di liquidazione con importo al netto delle trattenute
    
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        IRigaDoc = IRigaDoc + 1
        End With
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_POLiquidazioneRigheEla "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    
    While Not RsCorpo.EOF
        If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))
            
            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CONTO_PDC = GET_LINK_CONTO_PDC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_LIQUIDAZIONE = GET_LINK_CATEGORIA_LIQ(fnNotNullN(RsCorpo!IDArticolo))
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
            End If
            
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                    
                End If
            End If
            
            With ObjDoc.Tables
                
                .Field "Art_descrizione", fnNotNull(RsCorpo!Articolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
                
                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                If LINK_CONTO_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    SetPDCArticolo LINK_CONTO_PDC
                Else
                    If Link_PDC > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        SetPDCArticolo Link_PDC
                    End If
                End If
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POIDCaricoMerceRighe", fnNotNullN(RsCorpo!IDRV_POCaricoMerceRighe), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                .Field "RV_POOggettoVendita", fnNotNull(RsCorpo!Oggetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POOggettoConferimento", "Conferimento N° " & fnNotNullN(RsCorpo!NumeroDocumento) & " del " & fnNotNull(RsCorpo!DataConferimento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

            IRigaDoc = IRigaDoc + 1
            End With
        End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If

If LINK_TIPO_CORPO_FATTURA_SOCIO = 5 Then 'PER CATEGORIA MERCEOLOGICA
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        If (Len(RigaLiquidazione) > 0) Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            With ObjDoc.Tables
                .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            IRigaDoc = IRigaDoc + 1
            End With
        End If
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoCatMerc "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    'RsCorpo.Filter = "IDRV_POLiquidazione=" & IDLiquidazione
    
    While Not RsCorpo.EOF
        'If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))

            'LINK_IVA_ARTICOLO = GET_LINK_IVA_L(Me.cboAliquotaIVA.CurrentID)
            LINK_IVA_ARTICOLO = Me.cboAliquotaIVA.CurrentID
            LINK_CATEGORIA_MERCEOLOGICA = fnNotNullN(RsCorpo!IDCategoriaMerceologica)
            LINK_CONTO_PDC = 0
            LINK_CATEGORIA_LIQUIDAZIONE = 0
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
                
            End If
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                End If
            End If
            
'            If ((IMPORTO_UNITARIO < 0) And (fnNotNullN(RsCorpo!Quantita_Per_Totali) < 0)) Then
'                IMPORTO_UNITARIO = Abs(fnNotNullN(IMPORTO_UNITARIO))
'            End If
            
            With ObjDoc.Tables
                
                .Field "Art_descrizione", fnNotNull(RsCorpo!CategoriaMerceologica), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                Else
                    If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                '.Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali) + fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                
'                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
'                If LINK_CONTO_PDC > 0 Then
'                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                    SetPDCArticolo LINK_CONTO_PDC
'                Else
'                    If Link_PDC > 0 Then
'                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                        SetPDCArticolo Link_PDC
'                    End If
'                End If
                
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                
            IRigaDoc = IRigaDoc + 1
            
            
            End With
        'End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If
If LINK_TIPO_CORPO_FATTURA_SOCIO = 6 Then 'PER LOTTO DI PRODUZIONE
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        If (Len(RigaLiquidazione) > 0) Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            With ObjDoc.Tables
                .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            IRigaDoc = IRigaDoc + 1
            End With
        End If
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoLottoProd "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    'RsCorpo.Filter = "IDRV_POLiquidazione=" & IDLiquidazione
    
    While Not RsCorpo.EOF
        'If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))

            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(Me.cboAliquotaIVA.CurrentID)
            LINK_CATEGORIA_MERCEOLOGICA = 0 'fnNotNullN(RsCorpo!IDCategoriaMerceologica)
            LINK_CONTO_PDC = 0
            LINK_CATEGORIA_LIQUIDAZIONE = 0
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
                
            End If
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                End If
            End If
            
'            If ((IMPORTO_UNITARIO < 0) And (fnNotNullN(RsCorpo!Quantita_Per_Totali) < 0)) Then
'                IMPORTO_UNITARIO = Abs(fnNotNullN(IMPORTO_UNITARIO))
'            End If
            
            With ObjDoc.Tables
                
                .Field "Art_descrizione", fnNotNull(RsCorpo!LottoDiConferimento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                Else
                    If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                '.Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali) + fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                
'                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
'                If LINK_CONTO_PDC > 0 Then
'                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                    SetPDCArticolo LINK_CONTO_PDC
'                Else
'                    If Link_PDC > 0 Then
'                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                        SetPDCArticolo Link_PDC
'                    End If
'                End If
                
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                
            IRigaDoc = IRigaDoc + 1
            
            
            End With
        'End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If


If LINK_TIPO_CORPO_FATTURA_SOCIO = 7 Then 'Fattura dettagliata per lotto di produzione e articolo conferito
    
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        IRigaDoc = IRigaDoc + 1
        End With
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoLottoProdArtConf "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    
    While Not RsCorpo.EOF
        'If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))
            
            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CONTO_PDC = GET_LINK_CONTO_PDC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_LIQUIDAZIONE = GET_LINK_CATEGORIA_LIQ(fnNotNullN(RsCorpo!IDArticolo))
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
            End If
            
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                    
                End If
            End If
            
            With ObjDoc.Tables
                .Field "Art_descrizione", Mid(fnNotNull(RsCorpo!CodiceArticolo) & " - " & fnNotNull(RsCorpo!Articolo) & " (" & fnNotNull(RsCorpo!LottoDiConferimento) & ")", 1, 250), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    If (fnNotNullN(RsCorpo!Quantita_Per_Totali) > 0) Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
                
                .Field "Link_Art_unita_di_misura", GET_LINK_UM(fnNotNullN(RsCorpo!IDArticolo)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                If LINK_CONTO_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    SetPDCArticolo LINK_CONTO_PDC
                Else
                    If Link_PDC > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        SetPDCArticolo Link_PDC
                    End If
                End If
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                '.Field "RV_POIDCaricoMerceRighe", fnNotNullN(RsCorpo!IDRV_POCaricoMerceRighe), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                '.Field "RV_POOggettoVendita", fnNotNull(RsCorpo!Oggetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "RV_POOggettoConferimento", "Conferimento N° " & fnNotNullN(RsCorpo!NumeroDocumento) & " del " & fnNotNull(RsCorpo!DataConferimento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

            IRigaDoc = IRigaDoc + 1
            End With
        'End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If




If LINK_TIPO_CORPO_FATTURA_SOCIO = 8 Then 'Fattura dettagliata per articolo conferito
    
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        IRigaDoc = IRigaDoc + 1
        End With
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoArtConf "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    
    While Not RsCorpo.EOF
        'If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))
            
            LINK_IVA_ARTICOLO = GET_LINK_IVA_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CONTO_PDC = GET_LINK_CONTO_PDC_L(fnNotNullN(RsCorpo!IDArticolo))
            LINK_CATEGORIA_LIQUIDAZIONE = GET_LINK_CATEGORIA_LIQ(fnNotNullN(RsCorpo!IDArticolo))
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
            End If
            
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                    
                End If
            End If
            
            With ObjDoc.Tables
                .Field "Art_descrizione", Mid(fnNotNull(RsCorpo!CodiceArticolo) & " - " & fnNotNull(RsCorpo!Articolo), 1, 250), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    If fnNotNullN(RsCorpo!Quantita_Per_Totali) > 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
                
                .Field "Link_Art_unita_di_misura", GET_LINK_UM(fnNotNullN(RsCorpo!IDArticolo)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                If LINK_CONTO_PDC > 0 Then
                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    SetPDCArticolo LINK_CONTO_PDC
                Else
                    If Link_PDC > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        SetPDCArticolo Link_PDC
                    End If
                End If
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                '.Field "RV_POIDCaricoMerceRighe", fnNotNullN(RsCorpo!IDRV_POCaricoMerceRighe), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

                '.Field "RV_POOggettoVendita", fnNotNull(RsCorpo!Oggetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "RV_POOggettoConferimento", "Conferimento N° " & fnNotNullN(RsCorpo!NumeroDocumento) & " del " & fnNotNull(RsCorpo!DataConferimento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

            IRigaDoc = IRigaDoc + 1
            End With
        'End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If
If LINK_TIPO_CORPO_FATTURA_SOCIO = 9 Then 'PER CATEGORIA LIQUIDAZIONE
    If Me.chkRiportaRiferimentoLiq.Value = vbChecked Then
        RigaLiquidazione = GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione, idSocio, 0, 0, "")
        If (Len(RigaLiquidazione) > 0) Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            With ObjDoc.Tables
                .Field "Art_descrizione", RigaLiquidazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            IRigaDoc = IRigaDoc + 1
            End With
        End If
    End If
    
    IMPORTO_TRATTENUTE_AGGIUNTIVA = GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione)
    TOTALE_TRATTENUTE = GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione)
    IMPONIBILE_LIQUIDAZIONE = GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione)
    
    sSQL = "SELECT * FROM RV_PORepSubRiepilogoCategoriaLiquidazione "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    
    Set RsCorpo = New ADODB.Recordset
    
    RsCorpo.Open sSQL, CnDMT.InternalConnection
    'RsCorpo.Filter = "IDRV_POLiquidazione=" & IDLiquidazione
    
    While Not RsCorpo.EOF
        'If Len(fnNotNull(RsCorpo!CodiceArticolo)) > 0 Then
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
            
'            LINK_IVA_ARTICOLO = GET_LINK_IVA(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CATEGORIA_MERCEOLOGICA = GET_LINK_CATEGORIA_MERC(fnNotNull(RsCorpo!CodiceArticolo))
'            LINK_CONTO_PDC = GET_LINK_CONTO_PDC(fnNotNull(RsCorpo!CodiceArticolo))

            'LINK_IVA_ARTICOLO = GET_LINK_IVA_L(Me.cboAliquotaIVA.CurrentID)
            LINK_IVA_ARTICOLO = Me.cboAliquotaIVA.CurrentID
            LINK_CATEGORIA_MERCEOLOGICA = 0
            LINK_CONTO_PDC = 0
            LINK_CATEGORIA_LIQUIDAZIONE = fnNotNullN(RsCorpo!RV_POIDCategoriaLiquidazione)
            
            If TOTALE_TRATTENUTE > 0 Then
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / TOTALE_TRATTENUTE) * FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            Else
                TRATTENUTA_AGG_PER_RIGA = (IMPORTO_TRATTENUTE_AGGIUNTIVA / IMPONIBILE_LIQUIDAZIONE) * fnNotNullN(RsCorpo!ImponibileDaReg)
            End If
            
            'TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + FormatNumber(fnNotNullN(RsCorpo!TrattenuteTotali), 2)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA + fnNotNullN(RsCorpo!TrattenuteTotali)
            TRATTENUTA_AGG_PER_RIGA = TRATTENUTA_AGG_PER_RIGA
            
            If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                         If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!QuantitaVenduta)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select

            Else
                Select Case LINK_TIPO_IMPORTO_DOCUMENTO_LIQ
                        
                    Case 1
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                    Case 2
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImportoLordoDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                        
                    Case Else
                        If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / fnNotNullN(RsCorpo!Quantita_Per_Totali)
                        Else
                            IMPORTO_UNITARIO = (fnNotNullN(RsCorpo!ImponibileDaReg) - TRATTENUTA_AGG_PER_RIGA) / 1
                        End If
                End Select
                
            End If
            
            If Me.chkLordoLiquidazione.Value = vbChecked Then
                If LINK_IVA_CLIENTE = 0 Then
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_ARTICOLO) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                Else
                    ALIQUOTA_IVA = (GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE) / 100) + 1 '1.04 ' (fnNotNullN(RsCorpo!AliquotaIva_per_Imp_Vend) / 100) + 1
                    IMPORTO_UNITARIO = IMPORTO_UNITARIO / ALIQUOTA_IVA
                End If
            End If
            
'            If ((IMPORTO_UNITARIO < 0) And (fnNotNullN(RsCorpo!Quantita_Per_Totali) < 0)) Then
'                IMPORTO_UNITARIO = Abs(fnNotNullN(IMPORTO_UNITARIO))
'            End If
            
            With ObjDoc.Tables
                
                .Field "Art_descrizione", fnNotNull(RsCorpo!CategoriaLiquidazione), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If IsNull(RsCorpo!Quantita_Per_Totali) = True Then
                    If fnNotNullN(RsCorpo!QuantitaVenduta) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                Else
                    If fnNotNullN(RsCorpo!Quantita_Per_Totali) <> 0 Then
                        .Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                '.Field "Art_quantita_totale", fnNotNullN(RsCorpo!Quantita_Per_Totali) + fnNotNullN(RsCorpo!QuantitaVenduta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                '.Field "Art_prezzo_unitario_lordo_IVA", IMPORTO_UNITARIO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                If LINK_IVA_CLIENTE = 0 Then
                    .Field "Link_Art_IVA", LINK_IVA_ARTICOLO, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_ARTICOLO), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                Else
                    .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", GET_ALIQUOTAIVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                End If
                
'                .Field "Link_Art_unita_di_misura", GET_LINK_UM(RsCorpo!IDArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
'                If LINK_CONTO_PDC > 0 Then
'                    .Field "Link_Art_IDCContropartita", LINK_CONTO_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                    SetPDCArticolo LINK_CONTO_PDC
'                Else
'                    If Link_PDC > 0 Then
'                        .Field "Link_Art_IDCContropartita", Link_PDC, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                        SetPDCArticolo Link_PDC
'                    End If
'                End If
                
                .Field "RV_POIDCategoriaMerceologica", LINK_CATEGORIA_MERCEOLOGICA, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POIDCategoriaLiquidazione", LINK_CATEGORIA_LIQUIDAZIONE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "RV_POColli", fnNotNullN(RsCorpo!Colli), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoLordo", fnNotNullN(RsCorpo!PesoLordo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POTara", fnNotNullN(RsCorpo!Tara), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPesoNetto", fnNotNullN(RsCorpo!PesoNetto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "RV_POPezzi", fnNotNullN(RsCorpo!Pezzi), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                
            IRigaDoc = IRigaDoc + 1
            
            
            End With
        'End If
    RsCorpo.MoveNext
    Wend
    RsCorpo.Close
    Set RsCorpo = Nothing
End If


'*************************************************************************************

'INSERIMENTO RIGHE ACCONTO CON IVA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT RV_POAnticipazioniSocioRighe.IDRV_POAnticipazioniSocioRighe, "
sSQL = sSQL & "RV_POAnticipazioniSocioRighe.ImportoRestituito, RV_POAnticipazioniSocioRighe.IDIva, Iva.AliquotaIva, "
sSQL = sSQL & "DescrizioneTrattenutaAggiuntivaCapitale"
sSQL = sSQL & " FROM RV_POAnticipazioniSocioRighe INNER JOIN "
sSQL = sSQL & " Iva ON RV_POAnticipazioniSocioRighe.IDIva = Iva.IDIva "
sSQL = sSQL & " WHERE IDAnagrafica=" & idSocio
sSQL = sSQL & " AND IDRV_POTipoStatoAnticipazioneCapitale=1"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
    With ObjDoc.Tables
        .Field "Art_descrizione", fnNotNull(rs!DescrizioneTrattenutaAggiuntivaCapitale), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Link_Art_IVA", fnNotNullN(rs!IDIVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        If LINK_IVA_CLIENTE = 0 Then
            .Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", fnNotNullN(rs!IDIVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        Else
            .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End If
        .Field "Art_prezzo_unitario_neutro", -fnNotNullN(rs!ImportoRestituito) / ((fnNotNullN(rs!AliquotaIva) / 100) + 1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End With
    IRigaDoc = IRigaDoc + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing





'*******************INSERIMENTO ALTRE RIGHE DI LIQUIDAZIONE********************************
sSQL = "SELECT * FROM RV_POLiquidazioneRigheFatt "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
                
        
        With ObjDoc.Tables
            .Field "Art_descrizione", fnNotNull(rs!DescrizioneRiga), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", fnNotNullN(rs!IDIVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            If LINK_IVA_CLIENTE = 0 Then
                .Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", fnNotNullN(rs!IDIVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
            End If
            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            If fnNotNullN(rs!IDContoPDC) > 0 Then
                .Field fnNotNullN(rs!IDContoPDC), Link_PDC_Righe_Positive, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                If fnNotNullN(rs!ImportoUnitario) >= 0 Then
                    If Link_PDC_Righe_Positive > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC_Righe_Positive, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                Else
                    If Link_PDC_Righe_Negative > 0 Then
                        .Field "Link_Art_IDCContropartita", Link_PDC_Righe_Negative, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
            End If
        End With
IRigaDoc = IRigaDoc + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

'*******************************************************************************************

'*******************INSERIMENTO RIGA FINALE********************************
    If Len(RigaFatturazioneFinale) > 0 Then
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail IRigaDoc
                
        
        With ObjDoc.Tables
            .Field "Art_descrizione", RigaFatturazioneFinale, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End With
        
    End If
    
'*************************************************************************************

fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False
    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
    'VARErroreIDArticolo = vbCrLf & "Articolo: " & rsArt!Articolo
    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento ' & VARErroreIDArticolo


End Function

Private Function InserimentoDMT(IDLiquidazione As Long, IDAnagraficaSocio As Long) As Boolean
On Error GoTo ERR_InserimentoDMT
Dim VarNumeroDoc As String
Dim TotaleDocumento As Double
Dim TotaleAcconto As Double
Screen.MousePointer = vbHourglass
        
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    TotaleDocumento = ObjDoc.Field("Tot_documento_corr", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))
    If TotaleDocumento > 0 Then
        If NON_RIPORTA_ACCONTI_IN_FATTURA = 0 Then
            If Me.chkRaggrSocio.Value = vbUnchecked Then
                ObjDoc.Field "Doc_acconto_contante", GET_TOTALE_ACCONTO_IN_FATTURA(TotaleDocumento, IDAnagraficaSocio, IDLiquidazione), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            Else
                If Not ((rsLiquidazioniInFattura.EOF) And (rsLiquidazioniInFattura.BOF)) Then
                    TotaleAcconto = 0
                    rsLiquidazioniInFattura.MoveFirst
                    While Not rsLiquidazioniInFattura.EOF
                        TotaleAcconto = TotaleAcconto + GET_TOTALE_ACCONTO_IN_FATTURA(TotaleDocumento, IDAnagraficaSocio, fnNotNullN(rsLiquidazioniInFattura!IDLiquidazione))
                    rsLiquidazioniInFattura.MoveNext
                    Wend
                    
                    ObjDoc.Field "Doc_acconto_contante", TotaleAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                End If
            End If
        End If
    End If
    
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    VarNumeroDoc = ObjDoc.Insert
    If VarNumeroDoc > 0 Then
        If FLAG_NUMERAZIONE_COOPERATIVA = 1 Then
            ObjDoc.Field "Doc_numero_presso_nom", ObjDoc.Field("Doc_numero", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        ObjDoc.GeneratePreviewPNC
        ObjDoc.Update
        SCRIVI_ACCONTI_DA_FILIALE VarNumeroDoc, TotaleDocumento
        SCRIVI_ACCONTI_PER_ANTICIPAZIONE VarNumeroDoc, TotaleDocumento, IDAnagraficaSocio
        If (Me.chkRaggrSocio.Value = vbUnchecked) Then
            SCRIVI_ACCONTI_DA_CONFERIMENTO VarNumeroDoc, TotaleDocumento, IDAnagraficaSocio, IDLiquidazione
            fncAggiornaLiquidazione IDLiquidazione, VarNumeroDoc
        Else
            If Not ((rsLiquidazioniInFattura.EOF) And (rsLiquidazioniInFattura.BOF)) Then
                TotaleAcconto = 0
                rsLiquidazioniInFattura.MoveFirst
                While Not rsLiquidazioniInFattura.EOF
                    SCRIVI_ACCONTI_DA_CONFERIMENTO VarNumeroDoc, TotaleDocumento, IDAnagraficaSocio, fnNotNullN(rsLiquidazioniInFattura!IDLiquidazione)
                    fncAggiornaLiquidazione fnNotNullN(rsLiquidazioniInFattura!IDLiquidazione), VarNumeroDoc
                rsLiquidazioniInFattura.MoveNext
                Wend
            End If
        End If
    End If
    
Screen.MousePointer = vbDefault
    
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    
End Function
Private Sub GET_CICLO_LIQUIDAZIONI()
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer
Dim sSQL As String

Set rsLiquidazioni = New ADODB.Recordset
rsLiquidazioni.CursorLocation = adUseClient

rsLiquidazioni.Fields.Append "IDRiferimento", adInteger, , adFldIsNullable
rsLiquidazioni.Fields.Append "IDRiferimentoAnaFatt", adInteger, , adFldIsNullable

rsLiquidazioni.Open , , adOpenKeyset, adLockBatchOptimistic

If Me.chkRaggrSocio.Value = vbUnchecked Then
    sSQL = "SELECT * FROM RV_POTMPLiqDaFatt"
    Set rs = CnDMT.OpenResultset(sSQL)
    I = 1
    NumeroRecord = 0
    While Not rs.EOF
        'Cls_Nom.Add fnNotNullN(rs!IDLiquidazione), CStr("A" & I)
        rsLiquidazioni.AddNew
        rsLiquidazioni!IDRiferimento = fnNotNullN(rs!IDLiquidazione)
        rsLiquidazioni!IDRiferimentoAnaFatt = 0
        rsLiquidazioni.Update
        
        NumeroRecord = NumeroRecord + 1
        I = I + 1
    rs.MoveNext
    Wend

    rs.CloseResultset
    Set rs = Nothing
Else
    sSQL = "SELECT RV_POTMPLiqDaFatt.IDLiquidazione, RV_POLiquidazione.NumeroLiquidazione, RV_POLiquidazione.IDAnagrafica, Anagrafica.Anagrafica, "
    sSQL = sSQL & "RV_POLiquidazione.NettoLiquidazione, RV_POLiquidazione.IDAnagraficaFatturazione "
    sSQL = sSQL & "FROM RV_POTMPLiqDaFatt INNER JOIN "
    sSQL = sSQL & "RV_POLiquidazione ON RV_POTMPLiqDaFatt.IDLiquidazione = RV_POLiquidazione.IDRV_POLiquidazione INNER JOIN "
    sSQL = sSQL & "Fornitore ON RV_POLiquidazione.IDAnagrafica = Fornitore.IDAnagrafica INNER JOIN "
    sSQL = sSQL & "Anagrafica ON Fornitore.IDAnagrafica = Anagrafica.IDAnagrafica "
    sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    I = 1
    NumeroRecord = 0
    While Not rs.EOF
        rsLiquidazioni.Filter = "IDRiferimento=" & fnNotNullN(rs!IDAnagrafica) & " AND IDRiferimentoAnaFatt=" & fnNotNullN(rs!idanagraficaFatturazione)
        If rsLiquidazioni.EOF Then
            rsLiquidazioni.AddNew
            
            rsLiquidazioni!IDRiferimento = fnNotNullN(rs!IDAnagrafica)
            rsLiquidazioni!IDRiferimentoAnaFatt = fnNotNullN(rs!idanagraficaFatturazione)
            
            rsLiquidazioni.Update
            
            NumeroRecord = NumeroRecord + 1
            I = I + 1
            rsLiquidazioni.Filter = vbNullString
        End If
    rs.MoveNext
    Wend

    rs.CloseResultset
    Set rs = Nothing
End If
End Sub
Private Function GET_RIGA_DI_LIQUIDAZIONE(IDPeriodoLiquidazione As Long, idSocio As Long, AliquotaIva As Double, DescrizioneIVA As String, CodiceIVA As String) As String
Dim rs As DmtOleDbLib.adoResultset
Dim RsStringa As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim CodiceSocio As String
Dim Local_NumeroLiquidazioneAut As Long
Dim Local_NumeroLiquidazioneInt As Long
Dim Local_data_Inizio As String
Dim Local_data_Fine  As String

sSQL = "SELECT * FROM RV_POLiquidazionePeriodo WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Local_NumeroLiquidazioneAut = 0
    Local_NumeroLiquidazioneInt = 0
    Local_data_Inizio = ""
    Local_data_Fine = ""
Else
    Local_NumeroLiquidazioneAut = fnNotNullN(rs!NumeroLiquidazione)
    Local_NumeroLiquidazioneInt = fnNotNullN(rs!NumeroProtInt)
    Local_data_Inizio = fnNotNull(rs!DataInizio)
    Local_data_Fine = fnNotNull(rs!DataFine)
End If

rs.CloseResultset
Set rs = Nothing

''''''''''''COSTRUZIONE STRINGA''''''''''''''''
sSQL = "SELECT RV_POCostruzioneFatturaSocioRighe.Posizione, RV_POCostruzioneFatturaSocioRighe.Testo, "
sSQL = sSQL & "RV_POCostruzioneFatturaSocioRighe.IDRV_POCampiFattura, RV_POCostruzioneFatturaSocioRighe.IDRV_POCostruzioneFatturaSocio, "
sSQL = sSQL & "RV_POCostruzioneFatturaSocioRighe.IDRV_POCostruzioneFatturaSocioRighe , RV_POCostruzioneFatturaSocio.IDFiliale "
sSQL = sSQL & "FROM RV_POCostruzioneFatturaSocioRighe INNER JOIN "
sSQL = sSQL & "RV_POCostruzioneFatturaSocio ON "
sSQL = sSQL & "RV_POCostruzioneFatturaSocioRighe.IDRV_POCostruzioneFatturaSocio = RV_POCostruzioneFatturaSocio.IDRV_POCostruzioneFatturaSocio "
sSQL = sSQL & "WHERE RV_POCostruzioneFatturaSocio.IDFiliale=" & VarIDFiliale
sSQL = sSQL & " ORDER BY Posizione"


Set rs = CnDMT.OpenResultset(sSQL)

GET_RIGA_DI_LIQUIDAZIONE = ""

While Not rs.EOF
    Select Case fnNotNullN(rs!IDRV_POCampiFattura)
    
        Case 1 'TESTO
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & fnNotNull(rs!Testo)
        Case 2 'Valore dell'aliquota IVA
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & AliquotaIva
        Case 3 'Valore del codice IVA
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & CodiceIVA
        Case 4 'Descrizione dell'IVA
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & DescrizioneIVA
        Case 5 'Valore DAL del periodo di liquidazione
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & Local_data_Inizio
        Case 6 'Valore AL del periodo di liquidazione
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & Local_data_Fine
        Case 7 'Codice socio (Numerico)
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & GET_CODICESOCIO(VarIDAzienda, idSocio)
        Case 8 'Numero liquidazione automatico
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & Local_NumeroLiquidazioneAut
        Case 9 'Numero liquidazione interno
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & Local_NumeroLiquidazioneInt
        Case 10 ' Anagrafica del socio
            GET_RIGA_DI_LIQUIDAZIONE = GET_RIGA_DI_LIQUIDAZIONE & GET_ANAGRAFICASOCIO(VarIDAzienda, idSocio)
        
    End Select
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CODICESOCIO(IDAzienda As Long, idSocio As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT codice FROM Fornitore WHERE "
sSQL = sSQL & "IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica=" & idSocio

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_CODICESOCIO = "0"
Else
    GET_CODICESOCIO = fnNotNull(rs!Codice)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_ANAGRAFICASOCIO(IDAzienda As Long, idSocio As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Anagrafica, Nome FROM Anagrafica "
sSQL = sSQL & " WHERE IDAnagrafica=" & idSocio

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_ANAGRAFICASOCIO = ""
Else
    GET_ANAGRAFICASOCIO = fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ALIQUOTAIVA(IDIVA As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva WHERE "
sSQL = sSQL & "IDIva=" & IDIVA


Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_ALIQUOTAIVA = 0
Else
    GET_ALIQUOTAIVA = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub SetDefault()
    'Imposta i default per i campi relativi alla valuta corrente
    'al tipo di arrotondamento, alle spese e ai bolli
    'Questi default sono obbligatori per il calcolo del documento
    Set cDefault = New Collection
    'Valore di arrotondamento per la valuta corrente
    cDefault.Add 1, "Val_arrotondamento"
    'Tipo di arrotondomento per la valuta corrente
    cDefault.Add 1, "Link_Val_tipo_arrotondamento"
    'ID della valuta corrente
    cDefault.Add 0, "Link_Val_valuta"
    'Spesse incassi in percentuale
    cDefault.Add 0, "Nom_spese_incassi_perc"
    'Importo del bollo
    cDefault.Add 0, "Nom_bollo_esente"
    'Importo limite per il pagamento del bollo
    cDefault.Add 0, "Nom_bollo_esente_limite"
    'ID del contratto bancario azienda
    cDefault.Add 0, "Link_Doc_contratto_bancario_az"
    'ID della natura delle scadenze
    cDefault.Add 0, "IDNaturaScadenza"
End Sub
Private Sub TrovaAnagrafica(IDAna As Long)
    Dim sSQL As String
    Dim RSAna As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.CodiceFiscale, Anagrafica.PartitaIva, Anagrafica.Indirizzo, Anagrafica.Cap, Comune.Comune, Provincia.Provincia, Anagrafica.Telefono, Anagrafica.Fax, Fornitore.IDPagamentoDefault, Fornitore.Codice"
    sSQL = sSQL & " FROM ((Anagrafica LEFT JOIN Comune ON Anagrafica.IDComune = Comune.IDComune) LEFT JOIN Provincia ON Comune.IDProvincia = Provincia.IDProvincia) LEFT JOIN Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica"
    sSQL = sSQL & " WHERE (((Anagrafica.IDAnagrafica)=" & IDAna & "))"
    
    
    Set RSAna = CnDMT.OpenResultset(sSQL)
    
        If RSAna.EOF = False Then
            ArrayCli(0, 0) = RSAna!IDAnagrafica
            ArrayCli(0, 1) = RSAna!Anagrafica
            ArrayCli(0, 2) = fnNotNull(RSAna!Nome)
            ArrayCli(0, 3) = fnNotNull(RSAna!CodiceFiscale)
            ArrayCli(0, 4) = fnNotNull(RSAna!PartitaIVA)
            ArrayCli(0, 5) = fnNotNull(RSAna!Indirizzo)
            ArrayCli(0, 6) = fnNotNull(RSAna!CAP)
            ArrayCli(0, 7) = fnNotNull(RSAna!Comune)
            ArrayCli(0, 8) = fnNotNull(RSAna!Provincia)
            ArrayCli(0, 9) = fnNotNull(RSAna!Fax)
            ArrayCli(0, 10) = fnNotNull(RSAna!Telefono)
            ArrayCli(0, 11) = IIf((IsNull(RSAna!IDPagamentoDefault)), FrmFine.CboPagamento.CurrentID, fnNotNull(RSAna!IDPagamentoDefault))
            ArrayCli(0, 12) = fnNotNull(RSAna!Codice)
        End If
        
    RSAna.CloseResultset
    Set RSAna = Nothing
End Sub
Private Function GET_NUMERO_DOC_FATTURA_ACQ(idSocio As Long, IDEsercizio As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Prefisso As String

sSQL = "SELECT Numero, Prefisso FROM RV_PONumerazionePerSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & idSocio
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAzienda=" & VarIDAzienda
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_DOC_FATTURA_ACQ = 1
    NumeroDocumento = 1
Else
    Prefisso = ""
    If Len(Trim(fnNotNull(rs!Prefisso))) > 0 Then
        Prefisso = Replace(Trim(fnNotNull(rs!Prefisso)), "/", "")
        Prefisso = Trim(fnNotNull(rs!Prefisso)) & "/"
    End If
    
    'GET_NUMERO_DOC_FATTURA_ACQ = CStr(Trim(fnNotNull(rs!Prefisso)) & Trim(fnNotNullN(rs!Numero)))
    GET_NUMERO_DOC_FATTURA_ACQ = CStr(Prefisso & Trim(fnNotNullN(rs!Numero)))
    NumeroDocumento = fnNotNullN(rs!Numero)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fncAggiornaLiquidazione(IDLiquidazione As Long, NumeroDocumento As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(Numero=" & fnNormString(NumeroDocumento) & ") AND "
sSQL = sSQL & "(IDTipoOggetto=" & Link_TipoOggettoPerDoc & ") AND "
sSQL = sSQL & "(IDAzienda=" & VarIDAzienda & ") AND "
sSQL = sSQL & "(IDAttivitaAzienda=" & VarIDAttivitaAzienda & ") AND "
sSQL = sSQL & "(IDSezionale=" & ObjDoc.IDSezionale & ") AND "
sSQL = sSQL & "(DataEmissione=" & fnNormDate(ObjDoc.DataEmissione) & ")) "

Set rs = CnDMT.OpenResultset(sSQL)
If Not rs.EOF Then
    sSQL = "UPDATE RV_POLiquidazione SET "
    sSQL = sSQL & "IDOggetto=" & fnNotNullN(rs!IDOggetto) & ", "
    sSQL = sSQL & "Oggetto=" & fnNormString(fnNotNull(rs!Oggetto) & " N° " & fnNotNull(rs!Numero) & " del " & fnNotNull(rs!DataEmissione)) & ", "
    sSQL = sSQL & "PassaggioInFatturazione=" & fnNormBoolean(1) & " "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
    
    CnDMT.Execute sSQL

End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub fncAggiornaNumeroFattura(idSocio As Long, IDEsercizio As Long)
Dim sSQL As String
Dim rsCTRL As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PONumerazionePerSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & idSocio
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rsCTRL = CnDMT.OpenResultset(sSQL)

NumeroDocumento = NumeroDocumento + 1

If rsCTRL.EOF Then
    sSQL = "INSERT INTO RV_PONumerazionePerSocio ("
    sSQL = sSQL & "IDRV_PONumerazionePerSocio, IDAzienda, IDAnagrafica, IDEsercizio, Numero, Prefisso) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_PONumerazionePerSocio", "IDRV_PONumerazionePerSocio") & ", "
    sSQL = sSQL & VarIDAzienda & ", "
    sSQL = sSQL & idSocio & ", "
    sSQL = sSQL & IDEsercizio & ", "
    sSQL = sSQL & fnNormNumber(NumeroDocumento) & ", "
    sSQL = sSQL & fnNormString("") & ")"
Else
    sSQL = "UPDATE RV_PONumerazionePerSocio SET "
    sSQL = sSQL & "Numero=" & fnNotNullN(NumeroDocumento)
    sSQL = sSQL & " WHERE IDanagrafica=" & idSocio
    sSQL = sSQL & " AND IDAzienda= " & VarIDAzienda
    sSQL = sSQL & " AND IDEsercizio= " & IDEsercizio
End If

CnDMT.Execute sSQL



End Sub
Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(NomeReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = rs!IDReportTipoOggetto
Else
    fncTrovaReport = 0
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_fncTrovaReport:
    MsgBox Err.Description, vbCritical, "Trova report per stampa"
    fncTrovaReport = 0

End Function
Private Function fncImpostaDefaultReport(IDReportDefault As Long, IDTipoOggetto As Long, IDFiliale As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto
    sSQL = sSQL & " AND IDFiliale = " & IDFiliale
    
    CnDMT.Execute sSQL
    
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
    
    Set rs = CnDMT.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub fncReport()
    With Me.cboReport
        Set .Database = CnDMT
        .AddFieldKey "IDReportTipoOggetto"
        .DisplayField = "ReportTipoOggetto"
        .Sql = "SELECT * FROM ReportTipoOggetto WHERE IDTipoOggetto=" & Link_TipoOggetto
        .Fill
    End With
End Sub
Private Function fnDefaultReport(IDTipoOggetto As Long, IDFiliale As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDReportTipoOggetto FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fnDefaultReport = fnNotNullN(rs!IDReportTipoOggetto)
Else
    fnDefaultReport = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub StampaDocumento(IDOggetto As Long)
On Error GoTo ERR_StampaDocumento
Dim IDReportDefault As Long


Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
    
If MenuOptions.DBType = 1 Then
    'parametri di accesso al database ACCESS
    oReport.Password = "dmt192981046"
    oReport.User = "admin"
Else
    'parametri di accesso al database SQL Server
    oReport.Password = TheApp.Password '""
    oReport.User = TheApp.User '"sa"
End If

IDReportDefault = fnDefaultReport(Link_TipoOggetto, TheApp.Branch)

oReport.BranchID = TheApp.Branch 'IDFiliale

oReport.DocTypeID = Link_TipoOggetto


oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(Link_TipoOggetto) & ".IDOggetto = " & IDOggetto
oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser

oReport.Copies = FrmFine.txtNumeroCopie.Text

fncImpostaDefaultReport Me.cboReport.CurrentID, Link_TipoOggetto, TheApp.Branch

oReport.DoPrint oReport.PrinterName


fncImpostaDefaultReport IDReportDefault, Link_TipoOggetto, TheApp.Branch


Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "StampaDocumento"
End Sub
Private Sub fncSalvataggioImpostazioniDefaultFiliale()
Dim sSQL As String

sSQL = "UPDATE RV_POSchemaCoop SET "
sSQL = sSQL & "IDIvaFatturazione=" & Me.cboAliquotaIVA.CurrentID & ", "
sSQL = sSQL & "IDPagamentoFatturazione=" & Me.CboPagamento.CurrentID & ", "
sSQL = sSQL & "IDValutaFatturazione=" & Me.CboValuta.CurrentID & ", "
sSQL = sSQL & "IDCausaleContabileFatturazione=" & Me.cboCausaleContabile.CurrentID & ", "
sSQL = sSQL & "RaggruppaLiqPerSocio=" & Abs(Me.chkRaggrSocio.Value)
sSQL = sSQL & " WHERE IDFiliale=" & VarIDFiliale
sSQL = sSQL & " AND IDUtente=0"


CnDMT.Execute sSQL
End Sub

Private Function GET_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione As Long, NettoLiquidazionePerIVA As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Local_SommaTrattAgg As Double


'CALCOLO DEL TOTALE DELLE TRATTENUTE AGGIUNTIVE
sSQL = "SELECT SUM(ImportoTrattenuta) AS TotaleTrattenuteAggiuntive "
sSQL = sSQL & "FROM RV_POLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Local_SommaTrattAgg = 0
Else
    Local_SommaTrattAgg = fnNotNullN(rs!TotaleTrattenuteAggiuntive)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''CALCOLO DELLA PROPORZIONE DELLE TRATTENUTE AGGIUNTIVE'''''''''''''''''
sSQL = "SELECT NettoLiquidazione  "
sSQL = sSQL & "FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_TRATTENUTE_AGGIUNTIVE = 0
Else
    GET_IMPORTO_TRATTENUTE_AGGIUNTIVE = (Local_SommaTrattAgg / (fnNotNullN(rs!NettoLiquidazione) + Local_SommaTrattAgg)) * NettoLiquidazionePerIVA
End If

rs.CloseResultset
Set rs = Nothing

GET_IMPORTO_TRATTENUTE_AGGIUNTIVE = FormatNumber(GET_IMPORTO_TRATTENUTE_AGGIUNTIVE, 2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Function
Private Function GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE(IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Local_SommaTrattAgg As Double


'CALCOLO DEL TOTALE DELLE TRATTENUTE AGGIUNTIVE
sSQL = "SELECT SUM(ImportoTrattenuta) AS TotaleTrattenuteAggiuntive "
sSQL = sSQL & "FROM RV_POLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE = 0
Else
    GET_TOTALE_IMPORTO_TRATTENUTE_AGGIUNTIVE = fnNotNullN(rs!TotaleTrattenuteAggiuntive)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_TOTALE_TRATTENUTE_LIQUIDAZIONE(IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Local_SommaTrattAgg As Double


'CALCOLO DEL TOTALE DELLE TRATTENUTE AGGIUNTIVE
sSQL = "SELECT TotaleTrattenuta "
sSQL = sSQL & "FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_TRATTENUTE_LIQUIDAZIONE = 0
Else
    GET_TOTALE_TRATTENUTE_LIQUIDAZIONE = fnNotNullN(rs!TotaleTrattenuta)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_TOTALE_IMPONIBILE_LIQUIDAZIONE(IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Local_SommaTrattAgg As Double


'CALCOLO DEL TOTALE DELLE TRATTENUTE AGGIUNTIVE
sSQL = "SELECT TotaleDocumento "
sSQL = sSQL & "FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_IMPONIBILE_LIQUIDAZIONE = 0
Else
    GET_TOTALE_IMPONIBILE_LIQUIDAZIONE = fnNotNullN(rs!TotaleDocumento)
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Function GET_TIPO_CORPO(idSocio As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_TIPO_CORPO = 0

sSQL = "SELECT IDTipoCorpoFatturaSocio FROM RV_PO01_ConfigurazioneSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & idSocio
Set rs = CnDMT.OpenResultset(sSQL)
If Not rs.EOF Then
    GET_TIPO_CORPO = fnNotNullN(rs!IDTipoCorpoFatturaSocio)
End If

rs.CloseResultset
Set rs = Nothing

If GET_TIPO_CORPO > 0 Then Exit Function

sSQL = "SELECT IDTipoCorpoFatturaSocio FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_CORPO = 1
Else
    If fnNotNullN(rs!IDTipoCorpoFatturaSocio) = 0 Then
        GET_TIPO_CORPO = 1
    Else
        GET_TIPO_CORPO = fnNotNullN(rs!IDTipoCorpoFatturaSocio)
    End If
End If

rs.CloseResultset
Set rs = Nothing

If GET_TIPO_CORPO = 0 Then GET_TIPO_CORPO = 1

End Function
Private Function GET_LINK_IVA(CodiceArticolo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIvaAcquisto FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA = 0

Else
    GET_LINK_IVA = fnNotNullN(rs!IDIvaAcquisto)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CATEGORIA_MERC(CodiceArticolo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT  IDCategoriaMerceologica FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CATEGORIA_MERC = 0

Else
    GET_LINK_CATEGORIA_MERC = fnNotNullN(rs!IDCategoriaMerceologica)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraAcquisto, RV_POIDUnitaDiMisuraLiq"
sSQL = sSQL & " FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM = 0
Else
    If (fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq) = 0) Then
        GET_LINK_UM = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    Else
        GET_LINK_UM = GET_LINK_UM_PER_UM_COOP(fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq))
        If (GET_LINK_UM = 0) Then
            GET_LINK_UM = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM_PER_UM_COOP(IDUMCoop As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_LINK_UM_PER_UM_COOP = 0

sSQL = "SELECT * "
sSQL = sSQL & " FROM UnitaDiMisura "
sSQL = sSQL & " WHERE RV_POIDUnitaDiMisuraCoop=" & IDUMCoop


Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_PER_UM_COOP = fnNotNullN(rs!IDUnitaDiMisura)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TOTALE_ACCONTO_IN_FATTURA(TotaleDocumento As Double, idSocio As Long, IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim DaDataConf As String
Dim ADataConf As String

GET_TOTALE_ACCONTO_IN_FATTURA = 0

'''''''''''''''ACCONTI IN PARAMETRI FILIALE DI LIQUIDAZIONE'''''''''''''''''''
sSQL = "SELECT * FROM RV_POParametriTrattenuteAgg "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDRV_POTipoTrattenutaAggiuntiva=3"


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    GET_TOTALE_ACCONTO_IN_FATTURA = GET_TOTALE_ACCONTO_IN_FATTURA + ((TotaleDocumento / 100) * fnNotNullN(rs!Percentuale))
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT ImportoRestituito FROM RV_POAnticipazioniSocioRighe "
sSQL = sSQL & "WHERE IDAnagrafica=" & idSocio
sSQL = sSQL & " AND IDRV_POTipoStatoAnticipazioneCapitale=1"
sSQL = sSQL & " AND (IDIva IS NULL OR IDIva=0)"


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    GET_TOTALE_ACCONTO_IN_FATTURA = GET_TOTALE_ACCONTO_IN_FATTURA + fnNotNull(rs!ImportoRestituito)
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ACCONTI DA ADDEBITI CONFERIMENTI
'TROVO LA DATA DI INIZIO CONFERIMENTI PER QUESTA LIQUIDAZIONE
sSQL = "SELECT MIN(DataConferimento) as DataConferimento "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    DaDataConf = ""
Else
    DaDataConf = fnNotNull(rs!DataConferimento)
End If
rs.Close
Set rs = Nothing
'''TROVO LA DATA DI FINE CONFERIMENTI PER QUESTA LIQUIDAZIONE
sSQL = "SELECT MAX(DataConferimento) as DataConferimento "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    ADataConf = ""
Else
    ADataConf = fnNotNull(rs!DataConferimento)
End If
rs.Close
Set rs = Nothing

If DaDataConf = "" Then Exit Function
If ADataConf = "" Then Exit Function


'sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo, IDRV_POTipoTrattenutaAggiuntiva, "
sSQL = "SELECT SUM(TotaleRigaLordoIva) AS SommaTotaleRigaLordoIva, "
sSQL = sSQL & "SUM(ImpostaRiga) AS SommaImpostaRiga, "
sSQL = sSQL & "SUM(TotaleRigaNettoIva) AS SommaTotaleRigaNettoIva "
sSQL = sSQL & "FROM RV_POIECaricoMerceAddebiti "
sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
sSQL = sSQL & " AND IDAnagrafica = " & idSocio
sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DaDataConf)
sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(ADataConf)
sSQL = sSQL & " AND IDRV_POTipoTrattenutaAggiuntiva =3"
'sSQL = sSQL & " GROUP BY IDArticolo, CodiceArticolo, Articolo, IDRV_POTipoTrattenutaAggiuntiva"
'sSQL = sSQL & " HAVING IDRV_POTipoTrattenutaAggiuntiva =3 "

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection
If Not rs.EOF Then
    GET_TOTALE_ACCONTO_IN_FATTURA = GET_TOTALE_ACCONTO_IN_FATTURA + fnNotNullN(rs!SommaTotaleRigaNettoIva)
End If
rs.Close
Set rs = Nothing

End Function
Private Sub SCRIVI_ACCONTI_DA_FILIALE(NumeroDocumento As String, TotaleDocumento As Double)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggetto As Long
Dim rsNew As ADODB.Recordset


IDOggetto = ObjDoc.IDOggetto

'''''''''''''''''''''''''ACCONTI PER FILIALE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If TotaleDocumento > 0 Then
    sSQL = "SELECT * FROM RV_POParametriTrattenuteAgg "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDRV_POTipoTrattenutaAggiuntiva=3"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    Set rsNew = New ADODB.Recordset
    rsNew.Open "RV_POAccontiPerFattura", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rs.EOF
        rsNew.AddNew
            rsNew!IDRV_POAccontiPerFattura = fnGetNewKey("RV_POAccontiPerFattura", "IDRV_POAccontiPerFattura")
            rsNew!IDOggetto = IDOggetto
            rsNew!DescrizioneAcconto = fnNotNull(rs!DescrizioneTrattenutaAggiuntiva)
            rsNew!Percentuale = fnNotNullN(rs!Percentuale)
            rsNew!Importo = (TotaleDocumento / 100) * fnNotNullN(rs!Percentuale)
        rsNew.Update
    rs.MoveNext
    Wend
    rsNew.Close
    Set rsNew = Nothing
    rs.CloseResultset
    Set rs = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub SCRIVI_ACCONTI_PER_ANTICIPAZIONE(NumeroDocumento As String, TotaleDocumento As Double, IDAnagraficaSocio As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggetto As Long
Dim rsNew As ADODB.Recordset
Dim rsAnt As ADODB.Recordset


'sSQL = "SELECT * FROM Oggetto "
'sSQL = sSQL & "WHERE ("
'sSQL = sSQL & "(Numero=" & fnNormString(NumeroDocumento) & ") AND "
'sSQL = sSQL & "(IDTipoOggetto=" & Link_TipoOggettoPerDoc & ") AND "
'sSQL = sSQL & "(IDAzienda=" & TheApp.IDFirm & ") AND "
'sSQL = sSQL & "(IDAttivitaAzienda=" & VarIDAttivitaAzienda & ") AND "
'sSQL = sSQL & "(IDSezionale=" & ObjDoc.IDSezionale & ") AND "
'sSQL = sSQL & "(DataEmissione=" & fnNormDate(ObjDoc.DataEmissione) & ")) "

'Set rs = CnDMT.OpenResultset(sSQL)
'If Not rs.EOF Then
'    IDOggetto = fnNotNullN(rs!IDOggetto)
'Else

'    rs.CloseResultset
'    Set rs = Nothing
'    Exit Sub
'End If

'rs.CloseResultset
'Set rs = Nothing

IDOggetto = ObjDoc.IDOggetto
'''''''''''''''''''''''''ACCONTI PER FILIALE'''''''''''''''''''''''''
If TotaleDocumento > 0 Then
    sSQL = "SELECT * FROM RV_POAnticipazioniSocioRighe "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaSocio
    sSQL = sSQL & " AND IDRV_POTipoStatoAnticipazioneCapitale=1"
    
    Set rsAnt = New ADODB.Recordset
    rsAnt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    
    Set rsNew = New ADODB.Recordset
    rsNew.Open "RV_POAccontiPerFattura", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsAnt.EOF
        rsNew.AddNew
            rsNew!IDRV_POAccontiPerFattura = fnGetNewKey("RV_POAccontiPerFattura", "IDRV_POAccontiPerFattura")
            rsNew!IDOggetto = IDOggetto
            rsNew!DescrizioneAcconto = fnNotNull(rsAnt!DescrizioneTrattenutaAggiuntivaCapitale)
            rsNew!Percentuale = (fnNotNullN(rsAnt!ImportoRestituito) / TotaleDocumento) * 100
            rsNew!Importo = fnNotNullN(rsAnt!ImportoRestituito)
        rsNew.Update
        rsAnt!IDOggetto = IDOggetto
        rsAnt!NumeroDocumentoFattura = NumeroDocumento
        rsAnt!DataDocumentoFattura = ObjDoc.DataEmissione
        rsAnt!IDRV_POTipoStatoAnticipazioneCapitale = 2
        rsAnt!TipoStatoAnticipazioneCapitale = "Elaborato"
        rsAnt.Update
        
        GET_CALCOLA_TOTALI_ANTICIPAZIONE rsAnt!IDRV_POAnticipazioniSocio
    rsAnt.MoveNext
    Wend
    rsNew.Close
    Set rsNew = Nothing
    rsAnt.Close
    Set rsAnt = Nothing
End If

End Sub
Private Sub SCRIVI_ACCONTI_DA_CONFERIMENTO(NumeroDocumento As String, TotaleDocumento As Double, IDAnagraficaSocio As Long, IDLiquidazione As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDOggetto As Long
Dim rsNew As ADODB.Recordset
Dim rsAnt As ADODB.Recordset

Dim DaDataConf As String
Dim ADataConf As String

'sSQL = "SELECT * FROM Oggetto "
'sSQL = sSQL & "WHERE ("
'sSQL = sSQL & "(Numero=" & fnNormString(NumeroDocumento) & ") AND "
'sSQL = sSQL & "(IDTipoOggetto=" & Link_TipoOggettoPerDoc & ") AND "
'sSQL = sSQL & "(IDAzienda=" & TheApp.IDFirm & ") AND "
'sSQL = sSQL & "(IDAttivitaAzienda=" & VarIDAttivitaAzienda & ") AND "
'sSQL = sSQL & "(IDSezionale=" & ObjDoc.IDSezionale & ") AND "
'sSQL = sSQL & "(DataEmissione=" & fnNormDate(ObjDoc.DataEmissione) & ")) "

'Set rs = CnDMT.OpenResultset(sSQL)
'If Not rs.EOF Then
'    IDOggetto = fnNotNullN(rs!IDOggetto)
'Else
'    rs.CloseResultset
'    Set rs = Nothing
' '   Exit Sub
'End If

'rs.CloseResultset
'Set rs = Nothing

IDOggetto = ObjDoc.IDOggetto

If IDOggetto = 0 Then Exit Sub

'TROVO LA DATA DI INIZIO CONFERIMENTI PER QUESTA LIQUIDAZIONE
sSQL = "SELECT MIN(DataConferimento) as DataConferimento "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    DaDataConf = ""
Else
    DaDataConf = fnNotNull(rs!DataConferimento)
End If
rs.Close
Set rs = Nothing
'''TROVO LA DATA DI FINE CONFERIMENTI PER QUESTA LIQUIDAZIONE
sSQL = "SELECT MAX(DataConferimento) as DataConferimento "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    ADataConf = ""
Else
    ADataConf = fnNotNull(rs!DataConferimento)
End If
rs.Close
Set rs = Nothing

If DaDataConf = "" Then Exit Sub
If ADataConf = "" Then Exit Sub


'''''''''''''''''''''''''ACCONTI PER FILIALE'''''''''''''''''''''''''
If TotaleDocumento > 0 Then
    sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo, IDRV_POTipoTrattenutaAggiuntiva, "
    sSQL = sSQL & "SUM(TotaleRigaLordoIva) AS SommaTotaleRigaLordoIva, "
    sSQL = sSQL & "SUM(ImpostaRiga) AS SommaImpostaRiga, "
    sSQL = sSQL & "SUM(TotaleRigaNettoIva) AS SommaTotaleRigaNettoIva "
    sSQL = sSQL & "FROM RV_POIECaricoMerceAddebiti "
    sSQL = sSQL & "WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
    sSQL = sSQL & " AND IDAnagrafica = " & IDAnagraficaSocio
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(DaDataConf)
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(ADataConf)
    sSQL = sSQL & " GROUP BY IDArticolo, CodiceArticolo, Articolo, IDRV_POTipoTrattenutaAggiuntiva"
    sSQL = sSQL & " HAVING IDRV_POTipoTrattenutaAggiuntiva =3 "

    Set rsAnt = New ADODB.Recordset
    rsAnt.Open sSQL, CnDMT.InternalConnection
    
    
    sSQL = "SELECT * FROM RV_POAccontiPerFattura "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
    Set rsNew = New ADODB.Recordset
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsAnt.EOF
        rsNew.AddNew
            rsNew!IDRV_POAccontiPerFattura = fnGetNewKey("RV_POAccontiPerFattura", "IDRV_POAccontiPerFattura")
            rsNew!IDOggetto = IDOggetto
            rsNew!DescrizioneAcconto = fnNotNull(rsAnt!Articolo) & " - Conferimenti dal " & DaDataConf & " al " & ADataConf
            rsNew!Percentuale = (fnNotNullN(rsAnt!SommaTotaleRigaNettoIva) / TotaleDocumento) * 100
            rsNew!Importo = fnNotNullN(rsAnt!SommaTotaleRigaNettoIva)
        rsNew.Update
    rsAnt.MoveNext
    Wend
    rsNew.Close
    Set rsNew = Nothing
    rsAnt.Close
    Set rsAnt = Nothing
End If

End Sub
Private Sub GET_CALCOLA_TOTALI_ANTICIPAZIONE(IDTestataAnticipazione As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT * FROM RV_POAnticipazioniSocio "
sSQL = sSQL & "WHERE IDRV_POAnticipazioniSocio=" & IDTestataAnticipazione

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    rs!ImportoRestituito = GET_CAPITALE_RESTITUITO(IDTestataAnticipazione)
    rs!ImportoDaElaborare = GET_CAPITALE_DA_ELABORARE(IDTestataAnticipazione)
    rs!ImportoDaRestituire = rs!ImportoAnticipazione - (rs!ImportoRestituito + rs!ImportoDaElaborare)
    rs.Update
End If


rs.Close
Set rs = Nothing
End Sub
Private Function GET_CAPITALE_RESTITUITO(IDTestataAnticipazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(ImportoRestituito) AS Importo FROM RV_POAnticipazioniSocioRighe "
sSQL = sSQL & "WHERE IDRV_POAnticipazioniSocio=" & IDTestataAnticipazione
sSQL = sSQL & " AND IDRV_POTipoStatoAnticipazioneCapitale=2"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CAPITALE_RESTITUITO = 0
Else
    GET_CAPITALE_RESTITUITO = fnNotNullN(rs!Importo)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CAPITALE_DA_ELABORARE(IDTestataAnticipazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(ImportoRestituito) AS Importo FROM RV_POAnticipazioniSocioRighe "
sSQL = sSQL & "WHERE IDRV_POAnticipazioniSocio=" & IDTestataAnticipazione
sSQL = sSQL & " AND IDRV_POTipoStatoAnticipazioneCapitale=1"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CAPITALE_DA_ELABORARE = 0
Else
    GET_CAPITALE_DA_ELABORARE = fnNotNullN(rs!Importo)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Public Function GET_ALIQUOTA_IVA(IDIVA As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & "WHERE IDIva=" & IDIVA

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ALIQUOTA_IVA = 0
Else
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_TIPO_CORPO_DOCUMENTO(IDPeriodoLiquidazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_IMPORTO_DOCUMENTO_LIQ = 0
Else
    LINK_TIPO_IMPORTO_DOCUMENTO_LIQ = fnNotNullN(rs!IDTipoImportoDocumento)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento As Long, IDIvaCliente As Long) As Long
On Error GoTo ERR_GET_LINK_IVA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
Else
    If fnNotNullN(rs!IDIVA) > 0 Then
        GET_LINK_IVA_LETTERA_INTENTO = fnNotNullN(rs!IDIVA)
    Else
        GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_IVA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_IVA_LETTERA_INTENTO"
End Function
Private Sub GET_NUMERAZIONE_COOPERATIVA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NumerazioneCooperativaFACS, NonRiportareAccontiInFatturaPDC"
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    FLAG_NUMERAZIONE_COOPERATIVA = fnNotNullN(rs!NumerazioneCooperativaFACS)
    NON_RIPORTA_ACCONTI_IN_FATTURA = fnNotNullN(rs!NonRiportareAccontiInFatturaPDC)
Else
    FLAG_NUMERAZIONE_COOPERATIVA = 0
    NON_RIPORTA_ACCONTI_IN_FATTURA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_LINK_CONTO_PDC(CodiceArticolo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_CONTO_PDC = 0

'Contropartita articolo
sSQL = "SELECT IDPDCDare FROM Articolo "
sSQL = sSQL & " WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_LINK_CONTO_PDC = fnNotNullN(rs!IDPDCDare)
Else
    GET_LINK_CONTO_PDC = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub SetPDCArticolo(IDContoPDC As Long)
On Error Resume Next
Dim oNode As DmtPDC.INode
Dim oBranch As DmtPDC.Branch
Dim oPDC As DmtPDC.PDCServices

Dim CodiceConto As String
Dim DescrizioneConto As String
    
Set oPDC = New DmtPDC.PDCServices
'Imposta le proprietà dell'oggetto PDCServices
With oPDC
    'Viene fornita al controllo la connessione al database DMT.
    'La connessione è di tipo ADO.Connection quindi viene
    'passata la proprietà InternalConnection dell'oggetto Database
    Set .Connection = CnDMT.InternalConnection
    
    'Indica l'identificativo del Piano dei conti da visualizzare
    .IDPDC = Link_PianoDeiConti
    .HideAccounts = False
    .BranchType = btcAllBranchs
    '.BranchType = .BranchType + btcRevenuesBranch
    .AccountType = atcAllAccounts
    
    Set oNode = .GetAccount(IDContoPDC)
    
    
    'Codifica completa del Conto o del Ramo
    CodiceConto = oNode.CompletedCode
    DescrizioneConto = oNode.Description
    
End With

ObjDoc.Field "Art_CContropartita_codifica", CodiceConto, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
ObjDoc.Field "Art_CContropartita_descrizione", DescrizioneConto, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    
    
End Sub

Private Function GetPianoDeiConti() As Long
On Error GoTo ERR_GetPianoDeiConti
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
    sSQL = sSQL & "(IDAzienda = " & TheApp.IDFirm & ") AND "
    sSQL = sSQL & "(TipoPDC = " & 1 & ") AND "
    sSQL = sSQL & "(IDEsercizio= " & Me.cboEsercizio.CurrentID & "))"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GetPianoDeiConti = fnNotNullN(rs!IDPianoDeiConti)
    Else
        GetPianoDeiConti = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Exit Function
ERR_GetPianoDeiConti:
    
End Function
Private Function GET_LINK_IVA_L(IDCodiceArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIvaAcquisto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDCodiceArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_L = 0

Else
    GET_LINK_IVA_L = fnNotNullN(rs!IDIvaAcquisto)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CATEGORIA_MERC_L(IDCodiceArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT  IDCategoriaMerceologica FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDCodiceArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CATEGORIA_MERC_L = 0

Else
    GET_LINK_CATEGORIA_MERC_L = fnNotNullN(rs!IDCategoriaMerceologica)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CONTO_PDC_L(IDCodiceArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_CONTO_PDC_L = 0

'Contropartita articolo
sSQL = "SELECT IDPDCDare FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDCodiceArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_LINK_CONTO_PDC_L = fnNotNullN(rs!IDPDCDare)
Else
    GET_LINK_CONTO_PDC_L = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CATEGORIA_LIQ(IDCodiceArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT  RV_POIDCategoriaLiquidazione FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDCodiceArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CATEGORIA_LIQ = 0

Else
    GET_LINK_CATEGORIA_LIQ = fnNotNullN(rs!RV_POIDCategoriaLiquidazione)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub SCRIVI_RIFERIMENTI_CONFERIMENTI(IDAnagrafica As Long, IDAnagraficaSocio As Long)
On Error GoTo ERR_SCRIVI_RIFERIMENTI_CONFERIMENTI
Dim sSQL As String
Dim rsDoc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim DataInizio As String
Dim Mese As String
Dim DataFine As String
Dim rsPeriodo As DmtOleDbLib.adoResultset


Set rsDoc = New ADODB.Recordset
rsDoc.CursorLocation = adUseClient

rsDoc.Fields.Append "NumeroDocumento", adVarChar, 50, adFldIsNullable
rsDoc.Fields.Append "DataDocumento", adDBDate, , adFldIsNullable

rsDoc.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT RV_POLiquidazionePeriodo.DataInizio, RV_POLiquidazionePeriodo.DataFine, RV_POLiquidazione.IDOggetto "
sSQL = sSQL & "FROM RV_POLiquidazione INNER JOIN "
sSQL = sSQL & "RV_POLiquidazionePeriodo ON RV_POLiquidazione.IDRV_POLiquidazionePeriodo = RV_POLiquidazionePeriodo.IDRV_POLiquidazionePeriodo "
sSQL = sSQL & "GROUP BY RV_POLiquidazionePeriodo.DataInizio, RV_POLiquidazionePeriodo.DataFine, RV_POLiquidazione.IDOggetto "
sSQL = sSQL & "HAVING (RV_POLiquidazione.IDOggetto = " & ObjDoc.IDOggetto & ")"

Set rsPeriodo = CnDMT.OpenResultset(sSQL)

While Not rsPeriodo.EOF
    DataInizio = fnNotNull(rsPeriodo!DataInizio)
    DataFine = fnNotNull(rsPeriodo!DataFine)
    
    '''''RECUPERO DATI DALLA TESTA DEL DOCUMENTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT IDRV_POCaricoMerceTesta, IDTipoDocumentoCoop, NumeroDocumentoAcq, DataDocumentoAcq "
    sSQL = sSQL & "FROM RV_POCaricoMerceTesta "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND DataDocumentoAcq>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND DataDocumentoAcq<=" & fnNormDate(DataFine)
    'sSQL = sSQL & " AND IDAnagraficaFatturazione=" & IDAnagrafica
    sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaSocio
    'sSQL = sSQL & " AND IDTipoDocumentoCoop=1"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        If Len(fnNotNull(rs!NumeroDocumentoAcq)) > 0 Then
            rsDoc.Filter = "NumeroDocumento=" & fnNormString(rs!NumeroDocumentoAcq)
            If rsDoc.EOF Then
                rsDoc.AddNew
                    rsDoc!NumeroDocumento = rs!NumeroDocumentoAcq
                    rsDoc!DataDocumento = rs!DataDocumentoAcq
                rsDoc.Update
            End If
            rsDoc.Filter = vbNullString
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''RECUPERO DATI DALLE PESATURE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT RV_POCaricoMerceRighePes.IDRV_POCaricoMerceRighePes, RV_POCaricoMerceRighePes.NumeroDocumentoConsegna, RV_POCaricoMerceRighePes.DataDocumentoConsegna, "
    sSQL = sSQL & "RV_POCaricoMerceRighePes.IDRV_POCaricoMerceRighe , RV_POCaricoMerceRighePes.IDRV_POCaricoMerceTesta, RV_POCaricoMerceTesta.IDAzienda, RV_POCaricoMerceTesta.IDAnagrafica, "
    sSQL = sSQL & "RV_POCaricoMerceTesta.IDTipoDocumentoCoop, RV_POCaricoMerceTesta.IDAnagraficaFatturazione "
    sSQL = sSQL & "FROM RV_POCaricoMerceRighePes INNER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceRighePes.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe INNER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
    sSQL = sSQL & " WHERE RV_POCaricoMerceTesta.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND RV_POCaricoMerceRighePes.DataDocumentoConsegna>=" & fnNormDate(DataInizio)
    sSQL = sSQL & " AND RV_POCaricoMerceRighePes.DataDocumentoConsegna<=" & fnNormDate(DataFine)
    'sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDAnagraficaFatturazione=" & IDAnagrafica
    sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDAnagrafica=" & IDAnagraficaSocio
    'sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDTipoDocumentoCoop=1"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        If Len(fnNotNull(rs!NumeroDocumentoConsegna)) > 0 Then
            rsDoc.Filter = "NumeroDocumento=" & fnNormString(rs!NumeroDocumentoConsegna)
            If rsDoc.EOF Then
                rsDoc.AddNew
                    rsDoc!NumeroDocumento = rs!NumeroDocumentoConsegna
                    rsDoc!DataDocumento = rs!DataDocumentoConsegna
                rsDoc.Update
            End If
            rsDoc.Filter = vbNullString
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
rsPeriodo.MoveNext
Wend

rsPeriodo.CloseResultset
Set rsPeriodo = Nothing

If ((rsDoc.EOF) And (rsDoc.BOF)) Then
    rsDoc.Close
    Set rsDoc = Nothing
    Exit Sub
End If

rsDoc.MoveFirst

While Not rsDoc.EOF
    SCRIVI_DDT_RIF_XML fnNotNull(rsDoc!DataDocumento), fnNotNull(rsDoc!NumeroDocumento)
rsDoc.MoveNext
Wend

rsDoc.Close
Set rsDoc = Nothing
Exit Sub
ERR_SCRIVI_RIFERIMENTI_CONFERIMENTI:
    MsgBox Err.Description, vbCritical, "SCRIVI_RIFERIMENTI_CONFERIMENTI"
End Sub
Private Sub SCRIVI_DDT_RIF_XML(DataOrdine As String, NumeroOrdine As String)
On Error GoTo ERR_SCRIVI_ORD_CLI_RIF_XML
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
rsNew.AddNew
    rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
    rsNew!IDBloccoXML = 6
    rsNew!IDOggetto = ObjDoc.IDOggetto
    rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
    rsNew!RiferimentoNumeroLinea = 0
    rsNew!IDDocumento = 0
    If Len(DataOrdine) > 0 Then
        rsNew!Data = DataOrdine
    End If
    rsNew!NumItem = NumeroOrdine
rsNew.Update

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_SCRIVI_ORD_CLI_RIF_XML:
    MsgBox Err.Description, vbCritical, "SCRIVI_DDT_RIF_XML"

End Sub
