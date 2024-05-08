VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form FrmVisualizzaLiquidazione 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creazione liquidazione (3 di 4)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmVisualizzaLiquidazione.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   14310
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GridCastelletoIva 
      Height          =   1575
      Left            =   0
      TabIndex        =   9
      Top             =   6360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2778
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
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   5535
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9763
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
   Begin VB.CommandButton cmdStampa 
      Caption         =   "Stampa riepilogo (F10)"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdNuovoDettaglio 
      Caption         =   "Nuove trattenute (F9)"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdDettaglioTrattenute 
      Caption         =   "Dettaglio trattenute (F8)"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   11760
      TabIndex        =   2
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   13080
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "CASTELLETTO IVA PER LIQUIDAZIONE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Soci da liquidare"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "FrmVisualizzaLiquidazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Link_TipoOggettoLocal As Long
Private oReport As dmtReportLib.dmtReport

Private rsGriglia As ADODB.Recordset
Private rsGrigliaIva As ADODB.Recordset

Public Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    sSQL = "SELECT * FROM RV_POTMPLiquidazione "
    sSQL = sSQL & "WHERE IDRV_POPeriodoLiquidazione=" & LINK_PERIODO
    sSQL = sSQL & " ORDER BY DaRegistrare DESC, Anagrafica"
    
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
          
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                        
                .ColumnsHeader.Add "IDRV_POTMPLiquidazione", "ID", dgInteger, False, 500, dgAlignleft
                Set cl = .ColumnsHeader.Add("DaRegistrare", "Da registrare", dgBoolean, True, 1200, dgAlignleft)
                    cl.Editable = True
                .ColumnsHeader.Add "DataLiquidazione", "Data liq.", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "Anagrafica", "Socio/Fornitore", dgchar, True, 3000, dgAlignleft
                Set cl = .ColumnsHeader.Add("TotaleDocumento", "Tot. Doc. netto", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("TotaleIva", "Tot. Iva", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("TotaleDocumentoLordoIva", "Tot. Doc. Lordo IVA", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("TotaleTrattenuta", "Tot. Tratt.", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("TotaleTrattenuteAggiuntive", "Tot. Tratt. Agg.", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("TotaleTrattenuteRiepilogo", "Tot. Tratt. Riep.", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                
                Set cl = .ColumnsHeader.Add("NettoLiquidazione", "Tot. Liq", dgCurrency, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    cl.BackColor = &HF7C173
                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
End Sub
Public Sub fncGrigliaIva()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
'If TIPO_IMPORTO_ARTICOLO = 1 Then
    sSQL = "SELECT IDIva_per_Imp_Vend as IDIva , AliquotaIva_per_Imp_Vend As AliquotaIva, Iva_per_Imp_Vend as Iva, CodiceIva_per_Imp_Vend as CodiceIva,"
    sSQL = sSQL & "SUM(ImportoUnitarioDaReg) AS ImportoUnitarioDaReg, SUM(ImponibileDaReg) AS ImponibileDaReg, SUM(ImpostaDaReg) AS ImpostaDaReg,"
    sSQL = sSQL & "SUM(ImportoLordoDaReg) As ImportoLordoDaReg "
    sSQL = sSQL & "FROM RV_POTMPLiquidazioneRigheEla "
    sSQL = sSQL & "WHERE (RV_POTMPLiquidazioneRigheEla.IDAnagrafica = " & Me.Griglia.AllColumns("IDAnagrafica").Value & ") "
    sSQL = sSQL & " AND (RV_POTMPLiquidazioneRigheEla.IDRV_POLiquidazionePeriodo = " & LINK_PERIODO & ") "
    sSQL = sSQL & " AND (NOT (IDIva_per_imp_vend IS NULL))"
    sSQL = sSQL & " GROUP BY IDIva_per_Imp_Vend, AliquotaIva_per_Imp_Vend, Iva_per_Imp_Vend, CodiceIva_per_Imp_Vend"
'Else
'    sSQL = "SELECT IDIva_per_Imp_Medio as IDIva, AliquotaIva_per_Imp_Medio as AliquotaIva, Iva_per_Imp_Medio as Iva, CodiceIva_per_Imp_Medio as CodiceIva, SUM(ImponibileMedio) AS Imponibile, "
'    sSQL = sSQL & "SUM(ImpostaImponibileMedio) AS ImpostaImponibile, SUM(ImportoLordoMedio) AS ImportoLordo "
'    sSQL = sSQL & "From RV_POTMPLiquidazioneRigheEla "
'    sSQL = sSQL & "WHERE (RV_POTMPLiquidazioneRigheEla.IDAnagrafica = " & Me.Griglia.AllColumns("IDAnagrafica").Value & ") AND (RV_POTMPLiquidazioneRigheEla.IDRV_POLiquidazionePeriodo = " & LINK_PERIODO & ") "
'    sSQL = sSQL & " AND (NOT (IDIva_per_imp_vend IS NULL)) "
'
'    sSQL = sSQL & "GROUP BY IDIva_per_Imp_Medio, AliquotaIva_per_Imp_Medio, Iva_per_Imp_Medio, CodiceIva_per_Imp_Medio"
'End If

'sSQL = "SELECT IDRV_POLiquidazione, IDAnagrafica, IDRV_POLiquidazionePeriodo, Iva_per_Imp_Vend, CodiceIva_per_Imp_Vend, AliquotaIva_per_Imp_Vend, IDIva_per_Imp_Vend, "
'sSQL = sSQL & "IDIva_per_Imp_Medio , AliquotaIva_per_Imp_Medio, Iva_per_Imp_Medio, CodiceIva_per_Imp_Medio, SUM(ImpostaImponibileMedio)"
'sSQL = sSQL & "AS ImpostaImponibileMedio, SUM(ImpostaImponibileVenduto) AS ImpostaImponibileVenduto, SUM(ImponibileMedio) AS ImponibileMedio,"
'sSQL = sSQL & "SUM(ImponibileVenduto) AS ImponibileVenduto, SUM(ImportoLordoVenduto) AS ImportoLordoVenduto, SUM(ImportoLordoMedio) AS ImportoLordoMedio,"
'sSQL = sSQL & "SUM(ImportoUnitarioDaReg) AS ImportoUnitarioDaReg, SUM(ImponibileDaReg) AS ImponibileDaReg, SUM(ImpostaDaReg) AS ImpostaDaReg,"
'sSQL = sSQL & "SUM(ImportoLordoDaReg) As ImportoLordoDaReg "
'sSQL = sSQL & "From dbo.RV_POLiquidazioneRigheEla "
'sSQL = sSQL & "GROUP BY IDRV_POLiquidazione, IDAnagrafica, IDRV_POLiquidazionePeriodo, CodiceIva_per_Imp_Vend, AliquotaIva_per_Imp_Vend, IDIva_per_Imp_Vend, "
'sSQL = sSQL & "IDIva_per_Imp_Medio , AliquotaIva_per_Imp_Medio, Iva_per_Imp_Medio, CodiceIva_per_Imp_Medio, Iva_per_Imp_Vend "
'sSQL = sSQL & "HAVING (RV_POTMPLiquidazioneRigheEla.IDAnagrafica = " & Me.Griglia.AllColumns("IDAnagrafica").Value & ") "
'sSQL = sSQL & " AND (RV_POTMPLiquidazioneRigheEla.IDRV_POLiquidazionePeriodo = " & LINK_PERIODO & ") "
'sSQL = sSQL & " AND (NOT (IDIva_per_imp_vend IS NULL)) "
    
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        Set rsGrigliaIva = New ADODB.Recordset
        rsGrigliaIva.CursorLocation = adUseClient
        rsGrigliaIva.Open sSQL, CnDMT.InternalConnection
            
        With Me.GridCastelletoIva
                    .SelectionMode = dgSelectRow
                    .ColumnsHeader.Clear

                    .ColumnsHeader.Add "IDIva", "IDIva", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceIva", "Codice", dgchar, True, 700, dgAlignleft
                    .ColumnsHeader.Add "Iva", "Iva", dgchar, True, 1500, dgAlignleft
                    Set cl = .ColumnsHeader.Add("ImponibileDaReg", "Imponibile", dgCurrency, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("ImpostaDaReg", "Imposta", dgCurrency, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("ImportoLordoDaReg", "Imposta", dgCurrency, True, 1500, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."

                        
            Set .Recordset = rsGrigliaIva
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della liquidazione?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdAvanti_Click()
    rsGriglia.UpdateBatch
    Unload Me
End Sub

Private Sub cmdDettaglioTrattenute_Click()
Dim NumeroRecord As Long

    LINK_DOCUMENTO_TMP_LIQ = Me.Griglia.AllColumns("IDRV_POTMPLiquidazione").Value
    FrmDettaglioRigaElaborata.Show vbModal
        

End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub

Private Sub cmdNuovoDettaglio_Click()
Dim NumeroRecord As Long

    LINK_DOCUMENTO_TMP_LIQ = Me.Griglia.AllColumns("IDRV_POTMPLiquidazione").Value
    FrmInserimentoRighe.Show vbModal

    NumeroRecord = Me.Griglia.ListIndex - 1
    
    fncGriglia
    
    Me.Griglia.Recordset.Move NumeroRecord
End Sub

Private Sub cmdStampa_Click()
On Error GoTo ERR_cmdStampa_Click
Dim sSQLWHERE As String

    Link_TipoOggettoLocal = fncIDTipoOggettoPrg(App.EXEName)

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
    oReport.DocTypeID = Link_TipoOggettoLocal
    
    
    
    sSQLWHERE = "IDRV_POPeriodoLiquidazione = " & LINK_PERIODO
    sSQLWHERE = sSQLWHERE & " AND IDSocio = " & Me.Griglia.AllColumns("IDAnagrafica").Value
    
    
    oReport.Where = sSQLWHERE
    
        
    IDReport = fncTrovaReport("RV_PO01RiepilogoRegistrazione.rpt", Link_TipoOggettoLocal)

    If IDReport > 0 Then
        fncImpostaDefaultReport (IDReport)
        'Effettua l'anteprima di stampa
        'Settare il nome della stampante per questo tipo di stampa
        
        'oReport.PrinterName = fncTrovaStampante(IDReport)
        oReport.Preview 0, 0, 0
        'oReport.DoPrint "", 2
        
    Else
        MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
    End If
Exit Sub
ERR_cmdStampa_Click:
    MsgBox Err.Description, vbCritical, "stampa riepilogo "
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF8 Then
        LINK_DOCUMENTO_TMP_LIQ = Me.Griglia.AllColumns("IDRV_POTMPLiquidazione").Value
        FrmDettaglioRigaElaborata.Show vbModal
    End If
    
    If KeyCode = vbKeyF9 Then
        LINK_DOCUMENTO_TMP_LIQ = Me.Griglia.AllColumns("IDRV_POTMPLiquidazione").Value
        FrmInserimentoRighe.Show vbModal
    End If

    If KeyCode = vbKeyF10 Then
        cmdStampa_Click
    End If


End Sub


Private Sub Form_Load()
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fncGriglia
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        FrmFine.Show
        Exit Sub
    End If
    
    If Me.cmdIndietro.Value = True Then
        FrmNuovoPeriodo.Show
        Exit Sub
    End If
    
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
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("DaRegistrare").Value), 2
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
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("DaRegistrare").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
            rsGriglia.Fields("DaRegistrare").Value = Abs(CLng(Selected))
            'sbCheckSelected
            Me.Griglia.Refresh
        End If
End Sub

Private Sub Griglia_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    fncGrigliaIva
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
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & Link_TipoOggettoLocal & " AND IDFiliale = " & TheApp.Branch
    
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

