VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmFine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creazione liquidazione (4 di 4)"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
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
   ScaleHeight     =   6285
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Fine"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkUfficiale 
         Caption         =   "Imposta le liquidazioni ufficiali"
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
         TabIndex        =   12
         Top             =   1560
         Width           =   4455
      End
      Begin DMTDataCmb.DMTCombo cboReport 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   4815
         _ExtentX        =   8493
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
      Begin VB.OptionButton optRegistrazione 
         Caption         =   "Registrazione e stampa della liquidazione"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   4575
      End
      Begin VB.OptionButton optRegistrazione 
         Caption         =   "Registrazione della liquidazione"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Selezione report"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Registrazione della liquidazione"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      Picture         =   "FrmFine.frx":4781A
      ScaleHeight     =   4755
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   0
      Width           =   2745
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInfoDettaglio 
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   7815
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5400
      Width           =   7815
   End
End
Attribute VB_Name = "FrmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Link_TipoOggetto As Long
Private oReport As dmtReportLib.dmtReport

Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdAvanti_Click()
'On Error GoTo ERR_cmdAvanti_Click
Screen.MousePointer = 11

    Me.Enabled = False
    
    SalvaPeriodo Nuova_Liquidazione
    
    EliminazioDati
    
    Registrazione
    
    
    
    
    
    If TIPO_LIQUIDAZIONE = 3 Then
        AGGIORNA_DOCUMENTI_LIQUIDATI
    End If
    
    Me.Enabled = True
    
Screen.MousePointer = 0
Exit Sub
ERR_cmdAvanti_Click:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    Screen.MousePointer = 0
    Me.Enabled = True

End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    Me.optRegistrazione(0).Value = True
    
    Link_TipoOggetto = fncIDTipoOggettoPrg("RV_POLiquidazioneL")
    
    fncReport
    
    Me.cboReport.WriteOn fnDefaultReport
    Me.lblInfo.Caption = ""
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
On Error Resume Next
sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ReportTipoOggetto=" & fnNormString(NomeReport)
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = fnNotNullN(rs!IDReportTipoOggetto)
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
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & Link_TipoOggetto & " AND IDFiliale = " & TheApp.Branch
    
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
Private Function fnDefaultReport() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDReportTipoOggetto FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & Link_TipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fnDefaultReport = fnNotNullN(rs!IDReportTipoOggetto)
Else
    fnDefaultReport = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub EliminazioDati()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & LINK_PERIODO
If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & LINK_SOCIO
End If

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    'Eliminazione in RV_POLiquidazioneRighe
    sSQL = "DELETE FROM RV_POLiquidazioneRighe WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    CnDMT.Execute sSQL
    
    'Eliminazione in RV_POLiquidazioneRigheEla
    sSQL = "DELETE FROM RV_POLiquidazioneRigheEla WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    CnDMT.Execute sSQL
    
    'Eliminazione in RV_POLiquidazioneRigheFatt
    sSQL = "DELETE FROM RV_POLiquidazioneRigheFatt WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    CnDMT.Execute sSQL
   
    'Eliminazione in RV_POLiquidazioneRigheTratt
    sSQL = "DELETE FROM RV_POLiquidazioneRigheTratt WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    CnDMT.Execute sSQL
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Eliminazione in RV_POLiquidazione
sSQL = "DELETE FROM RV_POLiquidazione WHERE IDRV_POLiquidazionePeriodo=" & LINK_PERIODO
If LINK_SOCIO > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & LINK_SOCIO
End If

CnDMT.Execute sSQL


End Sub
Private Sub Registrazione()
Dim sSQL As String
Dim Link_Liquidazione As Long
Dim rsLiq As ADODB.Recordset
Dim Unita_Progresso As Double


Me.lblInfo.Caption = "START IN CORSO.........."
DoEvents
Me.lblInfoDettaglio.Caption = ""

sSQL = "SELECT * FROM RV_POTMPLiquidazione WHERE DaRegistrare= " & fnNormBoolean(1)
Set rsLiq = New ADODB.Recordset
rsLiq.Open sSQL, CnDMT.InternalConnection, adOpenKeyset
    
    
    
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 1000
If rsLiq.EOF = False Then
    rsLiq.MoveLast
    Unita_Progresso = Me.ProgressBar1.Max / rsLiq.RecordCount
Else
    Unita_Progresso = Me.ProgressBar1.Max
    MsgBox "Non risultano dati da registrare", vbInformation, "Registrazione liquidazione"
    Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"
    Me.lblInfoDettaglio.Caption = ""
    Exit Sub
End If

    
    rsLiq.MoveFirst
    GENERA_FILTRO_PER_TIPO_OGGETTO
    REGISTRAZIONE_PREZZI_MEDI LINK_PERIODO
    
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = 1000
    
    
    While Not rsLiq.EOF
        
        Me.lblInfo.Caption = "CREAZIONE LIQUIDAZIONE DEL SOCIO " & UCase(GET_RAGIONE_SOCIALE_SOCIO(fnNotNullN(rsLiq!IDAnagrafica)))
        DoEvents
        
        Link_Liquidazione = fnGetNewKey("RV_POLiquidazione", "IDRV_POLiquidazione")
        
        RegistrazioneTrattenuteElaborate rsLiq!IDRV_POTMPLiquidazione, Link_Liquidazione
        RegistrazioneNuoveTrattenute rsLiq!IDRV_POTMPLiquidazione, Link_Liquidazione
        RegistrazioneTipoTrattenuteUtilizzate Link_Liquidazione, fnNotNullN(rsLiq!IDAnagrafica), LINK_PERIODO
        
        sSQL = "INSERT INTO RV_POLiquidazione (IDRV_POLiquidazione, IDTipoOggetto, NumeroLiquidazione, NumeroProtInt, "
        sSQL = sSQL & "IDRV_POLiquidazionePeriodo, DataLiquidazione, IDAnagrafica, IDAnagraficaFatturazione, "
        sSQL = sSQL & "IDAzienda,IDFiliale, TotaleDocumento, TotaleIva, TotaleDocumentoLordoIva, TotaleTrattenuteConferimento, "
        sSQL = sSQL & "TotaleTrattenutaPerLavorazione, TotaleTrattenutaGenerale, TotaleTrattenuta, "
        sSQL = sSQL & "TotaleTrattenuteAggiuntive, TotaleTrattenuteRiepilogo, NettoLiquidazione, "
        sSQL = sSQL & "IDOggetto, Oggetto, PassaggioInFatturazione, Ufficiale, IDListino) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Link_Liquidazione & ", "
        sSQL = sSQL & fnGetTipoOggetto("RV_POLiquidazioneL") & ", "
        sSQL = sSQL & fnNotNullN(NUMERO_LIQUIDAZIONE) & ", "
        sSQL = sSQL & fnNotNullN(NUMERO_PROTOCOLLO) & ", "
        sSQL = sSQL & LINK_PERIODO & ", "
        sSQL = sSQL & fnNormDate(rsLiq!DataLiquidazione) & ", "
        sSQL = sSQL & fnNotNullN(rsLiq!IDAnagrafica) & ", "
        sSQL = sSQL & GET_LINK_ANA_FATT(fnNotNullN(rsLiq!IDAnagrafica)) & ", "
        sSQL = sSQL & TheApp.IDFirm & ", "
        sSQL = sSQL & TheApp.Branch & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleDocumento) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleIva) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleDocumentoLordoIva) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleTrattenuteConferimento) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TrattenutaPerLavorazione) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TrattenutaGenerale) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleTrattenuta) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleTrattenuteAggiuntive) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!TotaleTrattenuteRiepilogo) & ", "
        sSQL = sSQL & fnNormNumber(rsLiq!NettoLiquidazione) & ","
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & fnNormString("") & ", "
        sSQL = sSQL & fnNormBoolean(0) & ", "
        sSQL = sSQL & fnNormBoolean(Me.chkUfficiale.Value) & ", "
        sSQL = sSQL & fnNotNullN(rsLiq!IDListino) & ")"
        
        CnDMT.Execute sSQL
        
        If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
        End If
        If Me.optRegistrazione(1).Value = True Then
            StampaLiquidazione Link_Liquidazione
            DoEvents
        End If
        DoEvents
    rsLiq.MoveNext
    Wend
rsLiq.Close
Set rsLiq = Nothing

If Me.chkUfficiale.Value = vbChecked Then
    If LINK_TIPO_LIQ_CONF = 2 Then
        REGISTRA_RIGHE_CONFERIMENTO
    End If
End If

Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"
Me.lblInfoDettaglio.Caption = ""
Me.cmdAvanti.Enabled = False
End Sub
Private Sub RegistrazioneTrattenuteElaborate(IDTMPLiquidazione As Long, IDLiq As Long)
Dim sSQL As String
Dim rsEla As ADODB.Recordset
Dim rsUPD As ADODB.Recordset
Dim I As Integer
Dim ID As Long


Me.lblInfoDettaglio.Caption = "REGISTRAZIONE TRATTENUTE IN CORSO.........."
DoEvents

ID = fnGetNewKey("RV_POLiquidazioneRigheEla", "IDRV_POLiquidazioneRigheEla")


sSQL = "SELECT * FROM RV_POTMPLiquidazioneRigheEla WHERE IDRV_POTMPLiquidazione=" & IDTMPLiquidazione

Set rsEla = New ADODB.Recordset
Set rsUPD = New ADODB.Recordset

rsEla.Open sSQL, CnDMT.InternalConnection

sSQL = "SELECT * FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione = " & IDLiq

rsUPD.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsEla.EOF
    rsUPD.AddNew
    
        For I = 0 To rsUPD.Fields.Count - 1
            Select Case rsUPD.Fields(I).Name
                Case "IDRV_POLiquidazioneRigheEla"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = ID ' fnGetNewKey("RV_POLiquidazioneRigheEla", "IDRV_POLiquidazioneRigheEla")
                Case "IDRV_POLiquidazione"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = IDLiq
                Case "IDRV_POLiquidazionePeriodo"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = LINK_PERIODO
                Case "IDUtenteInserimento"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = TheApp.IDUser
                Case "DataInserimento"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = Date
                Case "PCInserimento"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = GET_NOMECOMPUTER
                Case "NomeUtentePCInserimento"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = GET_NOMEUTENTE
                Case "IDRV_POLiquidazioneRigheElaPM"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = GET_LINK_PREZZO_MEDIO_TMP(rsEla!IDRV_POTMPLiquidazionePrezzoMedio, LINK_PERIODO)
                Case "IDRV_POLiquidazioneConguaglio"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = LIQ_CONGUAGLIO
                Case "IDUtenteModifica"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = TheApp.IDUser
                Case "DataModifica"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = Date
                Case "PCModifica"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = GET_NOMECOMPUTER
                Case "NomeUtentePCModifica"
                    rsUPD.Fields(rsUPD.Fields(I).Name).Value = GET_NOMEUTENTE

                Case Else
                    rsUPD.Fields(rsEla.Fields(rsUPD.Fields(I).Name).Name).Value = rsEla.Fields(rsUPD.Fields(I).Name).Value
            End Select
        Next
        
        'rsUPD!IDRV_POLiquidazioneRigheEla = fnGetNewKey("RV_POLiquidazioneRigheEla", "IDRV_POLiquidazioneRigheEla")
        'rsUPD!IDRV_POLiquidazione = IDLiq
        'rsUPD!IDRV_POLiquidazionePeriodo = LINK_PERIODO
        
        '    rsUPD.Fields(rsEla.Fields(I).Name) = rsEla.Fields(I).Value
        'Next
        ID = ID + 1
    rsUPD.Update
rsEla.MoveNext
Wend
rsUPD.Close
Set rsUPD = Nothing

rsEla.Close
Set rs = Nothing
End Sub
Private Sub RegistrazioneNuoveTrattenute(IDTMPLiquidazione As Long, IDLiq As Long)
Dim sSQL As String
Dim rsEla As ADODB.Recordset
Dim ID As Long

Me.lblInfoDettaglio.Caption = "REGISTRAZIONE NUOVE TRATTENUTE........"
DoEvents

ID = fnGetNewKey("RV_POLiquidazioneRighe", "IDRV_POLiquidazioneRighe")

sSQL = "SELECT * FROM RV_POTMPLiquidazioneRighe WHERE IDRV_POTMPLiquidazione=" & IDTMPLiquidazione

Set rsEla = New ADODB.Recordset

rsEla.Open sSQL, CnDMT.InternalConnection
    
While Not rsEla.EOF
    sSQL = "INSERT INTO RV_POLiquidazioneRighe ("
    sSQL = sSQL & "IDRV_POLiquidazione, IDRV_POLiquidazioneRighe, DescrizioneAggiuntiva, "
    sSQL = sSQL & "IDRV_POTipoTrattenutaAggiuntiva, Percentuale, ImportoTrattenuta, "
    sSQL = sSQL & "IDRV_POAnticipazioniSocioRighe, IDRV_POAnticipazioniSocio, IDRV_POTipoRicalcoloComm, "
    sSQL = sSQL & "IDRV_POParametriTrattenuteAgg, IDRV_POCategoriaTrattenuteAggiuntive, CalcolaSuTrattenute) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDLiq & ", "
    'sSQL = sSQL & fnGetNewKey("RV_POLiquidazioneRighe", "IDRV_POLiquidazioneRighe") & ", "
    sSQL = sSQL & ID & ", "
    sSQL = sSQL & fnNormString(rsEla!DescrizioneAggiuntiva) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POTipoTrattenutaAggiuntiva) & ", "
    sSQL = sSQL & fnNormNumber(rsEla!Percentuale) & ", "
    sSQL = sSQL & fnNormNumber(rsEla!ImportoTrattenuta) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POAnticipazioniSocioRighe) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POAnticipazioniSocio) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POTipoRicalcoloComm) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POParametriTrattenuteAgg) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!IDRV_POCategoriaTrattenuteAggiuntive) & ", "
    sSQL = sSQL & fnNotNullN(rsEla!CalcolaSuTrattenute) & ")"
 
    CnDMT.Execute sSQL
    
    If fnNotNullN(rsEla!IDRV_POAnticipazioniSocioRighe) > 0 Then
        Me.lblInfoDettaglio.Caption = "REGISTRAZIONE ANTICIPAZIONI.........."
        DoEvents
        AGGIORNA_ANTICIPAZIONI IDLiq, fnNotNullN(rsEla!IDRV_POAnticipazioniSocioRighe), fnNotNullN(rsEla!IDRV_POAnticipazioniSocio)
    End If
    
    ID = ID + 1
rsEla.MoveNext
Wend
rsEla.Close
Set rs = Nothing
End Sub
Private Function AGGIORNA_ANTICIPAZIONI(IDLiquidazione As Long, IDAnticipazioneRighe As Long, IDAnticipazioneTesta As Long)
Dim sSQL As String
Dim rsAgg As ADODB.Recordset

sSQL = "SELECT * FROM RV_POAnticipazioniSocioRighe "
sSQL = sSQL & "WHERE IDRV_POAnticipazioniSocioRighe=" & IDAnticipazioneRighe

Set rsAgg = New ADODB.Recordset

rsAgg.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rsAgg.EOF Then
    rsAgg!IDRV_POLiquidazione = IDLiquidazione
    rsAgg!IDRV_POTipoStatoAnticipazioneInteresse = 2
    rsAgg!TipoStatoAnticipazioneInteresse = "Elaborato"
    rsAgg.Update
End If


rsAgg.Close
Set rsAgg = Nothing
End Function
Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdIndietro.Value = True Then
        FrmVisualizzaLiquidazione.Show
        Exit Sub
    End If
    
End Sub

Private Sub optRegistrazione_Click(Index As Integer)
If Index = 0 Then
    Me.cboReport.Enabled = False
Else
    Me.cboReport.Enabled = True
End If
End Sub
Private Sub StampaLiquidazione(IDLiquidazione As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POFiltro "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & "AND IDTipoOggetto=" & Link_TipoOggetto

CnDMT.Execute sSQL

GENERA_FILTRO_PER_TIPO_OGGETTO

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


'Imposta l'idfiliale di appartenenza del documento da stampare
    oReport.BranchID = TheApp.Branch 'IDFiliale
'Imposta l'identificativo del tipo di documento
    oReport.DocTypeID = Link_TipoOggetto
    'oReport.Where = "IDRegistro = " & Val(Me.Txt_Reg_IDRegistro)
    
    sSQL = "IDRV_POLiquidazione=" & IDLiquidazione
    sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser

    oReport.Where = sSQL
        
    IDReport = fncTrovaReport(Me.cboReport.Text, Link_TipoOggetto)

    If IDReport > 0 Then
        fncImpostaDefaultReport IDReport
        
        oReport.DoPrint oReport.PrinterName
        
    Else
        MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
    End If
Exit Sub

End Sub

Private Function fnGetTipoOggetto(Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub GENERA_FILTRO_PER_TIPO_OGGETTO()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim Filtro As String

sSQL = "DELETE FROM RV_POFiltro "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoOggetto=" & Link_TipoOggetto

CnDMT.Execute sSQL
    
Filtro = ""



sSQL = "SELECT * FROM RV_POFiltro"
Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenDynamic, adLockPessimistic

rs.AddNew
    rs!IDUtente = TheApp.IDUser
    rs!IDAzienda = TheApp.IDFirm
    rs!IDTipoOggetto = Link_TipoOggetto
    rs!Filtro = Filtro
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Sub RegistrazioneTipoTrattenuteUtilizzate(IDLiquidazione As Long, IDAnagraficaSocio As Long, IDPeriodoLiquidazione As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim ID As Long

Me.lblInfoDettaglio.Caption = "REGISTRAZIONE DEI TIPI DI TRATTENUTE UTILIZZATE......."
DoEvents

ID = fnGetNewKey("RV_POLiquidazioneRigheTratt", "IDRV_POLiquidazioneRigheTratt")

sSQL = "SELECT * FROM RV_POTMPLiquidazioneTrattConf "
sSQL = sSQL & "WHERE IDAnagraficaSocio=" & IDAnagraficaSocio

Set rs = CnDMT.OpenResultset(sSQL)

Set rsNew = New ADODB.Recordset
sSQL = "SELECT * FROM RV_POLiquidazioneRigheTratt "
sSQL = sSQL & "WHERE IDAnagraficaSocio=" & IDAnagraficaSocio
sSQL = sSQL & " AND IDRV_POLiquidazione=" & IDLiquidazione

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POLiquidazioneRigheTratt = ID ' fnGetNewKey("RV_POLiquidazioneRigheTratt", "IDRV_POLiquidazioneRigheTratt")
        rsNew!IDRV_POLiquidazione = IDLiquidazione
        rsNew!IDRV_POLiquidazionePeriodo = IDPeriodoLiquidazione
        rsNew!IDRV_POCaricoMerceRighe = rs!IDRV_POCaricoMerceRighe
        rsNew!IDArticoloVendita = rs!IDArticoloVendita
        rsNew!IDRV_POTrattenutaPerLiquidazione = rs!IDRV_POTrattenutaPerLiquidazione
        rsNew!ValoreTrattenuta1 = rs!ValoreTrattenuta1
        rsNew!ValoreTrattenuta2 = rs!ValoreTrattenuta2
        rsNew!PercTrattenuta1 = rs!PercTrattenuta1
        rsNew!PercTrattenuta2 = rs!PercTrattenuta2
        rsNew!ValoreTrattenuta1Conf = rs!ValoreTrattenuta1Conf
        rsNew!ValoreTrattenuta2Conf = rs!ValoreTrattenuta2Conf
        rsNew!IDTipoOggetto = rs!IDTipoOggetto
        rsNew!IDOggetto = rs!IDOggetto
        rsNew!IDValoriOggettoDettaglio = rs!IDValoriOggettoDettaglio
        rsNew!IDRV_POTipoTrattenuta = rs!IDRV_POTipoTrattenuta
        rsNew!SoloRigaConferimento = rs!SoloRigaConferimento
        rsNew!IDAnagraficaSocio = rs!IDAnagraficaSocio
    rsNew.Update
    ID = ID + 1
rs.MoveNext
Wend


rsNew.Close
Set rsNew = Nothing
rs.CloseResultset
Set rs = Nothing


End Sub
Private Function GET_RAGIONE_SOCIALE_SOCIO(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Anagrafica, Nome FROM Anagrafica "
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_RAGIONE_SOCIALE_SOCIO = ""
Else
    GET_RAGIONE_SOCIALE_SOCIO = fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AGGIORNA_DOCUMENTI_LIQUIDATI()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto, IDTipoOggetto "
sSQL = sSQL & " FROM RV_POTMPLiqFattureSel "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Registra=1"


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Select Case fnNotNullN(rs!IDTipoOggetto)
        Case 2
            sSQL = "UPDATE ValoriOggettoPerTipo0002 SET "
            sSQL = sSQL & "RV_POIDLiquidazionePeriodo=" & LINK_PERIODO
            sSQL = sSQL & " WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
            CnDMT.Execute sSQL
        Case 114
            sSQL = "UPDATE ValoriOggettoPerTipo0072 SET "
            sSQL = sSQL & "RV_POIDLiquidazionePeriodo=" & LINK_PERIODO
            sSQL = sSQL & " WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
            CnDMT.Execute sSQL
        Case 8
            sSQL = "UPDATE ValoriOggettoPerTipo0008 SET "
            sSQL = sSQL & "RV_POIDLiquidazionePeriodo=" & LINK_PERIODO
            sSQL = sSQL & " WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
            CnDMT.Execute sSQL
    End Select
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_LINK_ANA_FATT(IDAnagraficaSocio) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagraficaFatturazione "
sSQL = sSQL & "FROM RV_PO01_ConfigurazioneSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaSocio


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANA_FATT = IDAnagraficaSocio
Else
    If fnNotNullN(rs!IDAnagraficaFatturazione) = 0 Then
        GET_LINK_ANA_FATT = IDAnagraficaSocio
    Else
        GET_LINK_ANA_FATT = fnNotNullN(rs!IDAnagraficaFatturazione)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub REGISTRAZIONE_PREZZI_MEDI(IDPeriodoLiquidazione As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim NumeroRecord As Long
Dim Unita_Progresso As Double
Dim ID As Long

Me.lblInfoDettaglio.Caption = "CALCOLO NUMERO DEI PREZZI MEDI CALCOLATI......."
DoEvents

sSQL = "SELECT COUNT(IDRV_POTMPLiquidazionePrezzoMedio) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POTMPLiquidazionePrezzoMedio "
'sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

Me.lblInfoDettaglio.Caption = "REGISTRAZIONE DEI PREZZI MEDI CALCOLATI......."
DoEvents

Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)


ID = fnGetNewKey("RV_POLiquidazioneRigheElaPM", "IDRV_POLiquidazioneRigheElaPM")

sSQL = "SELECT * FROM RV_POLiquidazioneRigheElaPM "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POTMPLiquidazionePrezzoMedio "
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POLiquidazioneRigheElaPM = ID 'fnGetNewKey("RV_POLiquidazioneRigheElaPM", "IDRV_POLiquidazioneRigheElaPM")
        rsNew!IDRV_POLiquidazionePeriodo = IDPeriodoLiquidazione
        rsNew!IDRV_POTipoPrezzoMedio = fnNotNullN(rs!IDTipoPrezzoMedio)
        
        rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsNew!IDCategoriaMerceologica = fnNotNullN(rs!IDCategoriaMerceologica)
        rsNew!IDSocio = fnNotNullN(rs!IDSocio)
        
        rsNew!DaDataConferimento = rs!DaDataConferimento
        rsNew!ADataConferimento = rs!ADataConferimento
        rsNew!DaDataVendita = rs!DaDataVendita
        rsNew!ADataVendita = rs!ADataVendita
        rsNew!DaDataLavorazione = rs!DaDataLavorazione
        rsNew!ADataLavorazione = rs!ADataLavorazione
        
        
        rsNew!PrezzoMedio = fnNotNullN(rs!PrezzoMedio)
        rsNew!ImportoSconti = fnNotNullN(rs!PrezzoSconti)
        rsNew!ImportoCommissioni = fnNotNullN(rs!PrezzoCommissioni)
        rsNew!ImportoVarInclusoImballo = fnNotNullN(rs!PrezzoVarInclusoImballo)
        rsNew!PrezzoMedioNettoIva = fnNotNullN(rs!PrezzoMedioNettoIva)
    
        rsNew!IDRV_POTMPLiquidazionePrezzoMedio = fnNotNullN(rs!IDRV_POTMPLiquidazionePrezzoMedio)
        
        ID = ID + 1
    rsNew.Update
    
    REGISTRAZIONE_PREZZI_MEDI_RIGHE rsNew!IDRV_POLiquidazioneRigheElaPM, rsNew!IDRV_POTMPLiquidazionePrezzoMedio, IDPeriodoLiquidazione
    
    If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
    End If
    
    DoEvents
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
End Sub
Private Sub REGISTRAZIONE_PREZZI_MEDI_RIGHE(IDRigaPrezzoMedio As Long, IDRigaPrezzoMedioTMP As Long, IDPeriodoLiquidazione As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim ID As Long

ID = fnGetNewKey("RV_POLiquidazioneRigheElaPMRighe", "IDRV_POLiquidazioneRigheElaPMRighe")

sSQL = "SELECT * FROM RV_POLiquidazioneRigheElaPMRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POTMPLiquidazionePrezzoMedioRighe "
sSQL = sSQL & " WHERE IDRV_POTMPLiquidazionePrezzoMedio=" & IDRigaPrezzoMedioTMP
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POLiquidazioneRigheElaPMRighe = ID ' fnGetNewKey("RV_POLiquidazioneRigheElaPMRighe", "IDRV_POLiquidazioneRigheElaPMRighe")
        rsNew!IDRV_POLiquidazioneRigheElaPM = IDRigaPrezzoMedio
        rsNew!IDRV_POLiquidazionePeriodo = IDPeriodoLiquidazione
        
        rsNew!IDTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
        rsNew!IDOggetto = fnNotNullN(rs!IDOggetto)
        rsNew!IDValoriOggettoDettaglio = fnNotNullN(rs!IDValoriOggettoDettaglio)
        rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsNew!IDCategoriaMerceologica = fnNotNullN(rs!IDCategoriaMerceologica)
        rsNew!Quantita = fnNotNullN(rs!Quantita)
        rsNew!QuantitaDocumento = fnNotNullN(rs!QuantitaDocumento)
        rsNew!ImportoNettoVendita = fnNotNullN(rs!ImportoNettoVendita)
        rsNew!ImportoSconti = fnNotNullN(rs!ImportoSconti)
        rsNew!ImportoVariazionePrezzoImballo = fnNotNullN(rs!ImportoVarImballo)
        rsNew!ImportoCommissioni = fnNotNullN(rs!ImportoCommissioni)
        rsNew!ImportoLiquidazione = fnNotNullN(rs!ImportoLiquidazioni)
        rsNew!IDRV_POTMPLiquidazionePrezzoMedio = fnNotNullN(rs!IDRV_POTMPLiquidazionePrezzoMedio)
        rsNew!IDRV_POTMPLiquidazionePrezzoMedioRighe = fnNotNullN(rs!IDRV_POTMPLiquidazionePrezzoMedioRighe)
        rsNew!DataVendita = rs!DataVendita
        rsNew!DataConferimento = rs!DataConferimento
        rsNew!DataLavorazione = rs!DataLavorazione
        
        ID = ID + 1
    rsNew.Update
    
        
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing



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
Private Function GET_LINK_PREZZO_MEDIO_TMP(IDTmpPrezzoMedio As Long, IDPeriodoLiquidazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POLiquidazioneRigheElaPM "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheElaPM "
sSQL = sSQL & "WHERE IDRV_POTMPLiquidazionePrezzoMedio=" & IDTmpPrezzoMedio
sSQL = sSQL & " AND IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PREZZO_MEDIO_TMP = 0
Else
    GET_LINK_PREZZO_MEDIO_TMP = fnNotNullN(rs!IDRV_POLiquidazioneRigheElaPM)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub REGISTRA_RIGHE_CONFERIMENTO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.lblInfoDettaglio.Caption = "CAMBIO STATO RIGHE CONFERIMENTI.........."
DoEvents


sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POTMPLiquidazioneRigheConf "

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    sSQL = "UPDATE RV_POCaricoMerceRighe SET "
    sSQL = sSQL & "IDRV_POTipoConfLiquidazione=2 "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    
    CnDMT.Execute sSQL
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
