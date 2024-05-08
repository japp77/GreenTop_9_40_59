VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.0#0"; "DMTDataCmb.OCX"
Begin VB.Form FrmStampaEtichette 
   Caption         =   "Stampa etichette"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12930
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Height          =   8055
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.TextBox txtUltimaStampaPedana 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   5640
         Width           =   12735
      End
      Begin VB.TextBox txtNumeroUltimaStampa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox txtDataOraUltimaStampa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2640
         Width           =   7095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   3240
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Configurazione stampa"
         Height          =   1935
         Left            =   0
         TabIndex        =   12
         Top             =   6000
         Width           =   12735
         Begin VB.CommandButton cmdStampaPedana 
            Caption         =   "Stampa pedana"
            Height          =   375
            Left            =   9120
            TabIndex        =   15
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtCommentoStampaPed 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H0000FFFF&
            Height          =   1095
            Left            =   5280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   600
            Width           =   3615
         End
         Begin VB.ComboBox cboStampantePed 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   5055
         End
         Begin DMTDataCmb.DMTCombo cboReportPed 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   5055
            _ExtentX        =   8916
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
         Begin VB.Label Label1 
            Caption         =   "Descrizione stampa"
            Height          =   255
            Index           =   5
            Left            =   5280
            TabIndex        =   19
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Stampante"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label1 
            Caption         =   "Report"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Configurazione stampa"
         Height          =   1815
         Left            =   0
         TabIndex        =   1
         Top             =   3480
         Width           =   12735
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stampa etichetta pedana"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9120
            TabIndex        =   25
            Top             =   240
            Width           =   2775
         End
         Begin VB.ComboBox cboStampa 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtCommentoStampaLav 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H0000FFFF&
            Height          =   1125
            Left            =   5280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   480
            Width           =   3615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Stampa tutte le etichette"
            Height          =   375
            Left            =   9120
            TabIndex        =   3
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton cmdStampaSingola 
            Caption         =   "Stampa riga selezionata"
            Height          =   375
            Left            =   9120
            TabIndex        =   2
            Top             =   600
            Width           =   2295
         End
         Begin DMTDataCmb.DMTCombo cboReport 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   5055
            _ExtentX        =   8916
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
         Begin VB.Label Label1 
            Caption         =   "Report"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label Label1 
            Caption         =   "Stampante"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label1 
            Caption         =   "Descrizione stampa"
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   7
            Top             =   240
            Width           =   3615
         End
      End
      Begin DmtGridCtl.DmtGrid GridEtichette 
         Height          =   2295
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4048
         BackColor       =   12582912
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMove      =   0   'False
         UpdatePosition  =   0   'False
         UseUserSettings =   0   'False
         ColumnsHeaderHeight=   20
      End
      Begin VB.Label lblInfoEtichetteLav 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Top             =   3000
         Width           =   11535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LAVORAZIONI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   12735
      End
      Begin VB.Label lblPedana 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   5280
         Width           =   12735
      End
   End
End
Attribute VB_Name = "FrmStampaEtichette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function GetProfileString Lib "kernel32" _
'    Alias "GetProfileStringA" _
'    (ByVal lpAppName As String, _
'    ByVal lpKeyName As String, _
'    ByVal lpDefault As String, _
'    ByVal lpReturnedString As String, _
'    ByVal nSize As Long) As Long
 
'Private Declare Function WriteProfileString Lib "kernel32" _
'    Alias "WriteProfileStringA" _
'    (ByVal lpszSection As String, _
'    ByVal lpszKeyName As String, _
'    ByVal lpszString As String) As Long
 
'Private Declare Function SendMessage Lib "user32" _
'    Alias "SendMessageA" _
'    (ByVal hwnd As Long, _
'    ByVal wMsg As Long, _
'    ByVal wParam As Long, _
'    lparam As String) As Long
Private rsGriglia As ADODB.Recordset
Private oReport As dmtReportLib.dmtReport
Private ColliStampati As Long
Private ColliDaStampare As Long
Private Link_TipoOggettoLocal As Long
Private NuovoRecordPerStampa As Long

Private CODICE_PEDANA As String


Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
            rsGriglia.Fields("DaStampare").Value = Abs(CLng(Selected))
            
            Me.GridEtichette.Refresh
        End If
End Sub



Private Sub cboReport_Click()
    
    Me.cboStampa.Text = Get_TrovaStampante
    
End Sub
Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub

Private Sub cmdStampaPedana_Click()
'On Error GoTo ERR_cmdStampaPedana_Click
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsEti As ADODB.Recordset


If Me.cboReportPed.CurrentID = 0 Then
    MsgBox "Selezionare una stampa per etichette di pedana", vbInformation, "Stampa etichette"
    Exit Sub
End If
If Len(Me.cboStampantePed.Text) = 0 Then
    MsgBox "Selezionare la stampante per le etichette di pedana", vbInformation, "Stampa etichette"
    Exit Sub
End If



sSQL = "DELETE FROM RV_POTMPStampaEtichetteRighe "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
Cn.Execute sSQL

Set rs = New ADODB.Recordset
Set rsEti = New ADODB.Recordset


rsEti.Open "RV_POTMPStampaEtichetteRighe", Cn.InternalConnection, adOpenDynamic, adLockPessimistic

rs.Open "RV_POIEStampaEtichette", Cn.InternalConnection, adOpenDynamic, adLockPessimistic
rs.Filter = "CodicePedana = " & CODICE_PEDANA
If rs.EOF = False Then
    Me.lblInfoEtichetteLav.Caption = "STAMPA PEDANA " & Me.GridEtichette.AllColumns("CodicePedana").Value

    
    DoEvents
    While Not rs.EOF
        rsEti.AddNew
            rsEti!IDUtente = TheApp.IDUser
            rsEti!IDAzienda = TheApp.IDFirm
            rsEti!IDFiliale = TheApp.Branch
            rsEti!IDRV_POPedana = fnNotNullN(rs!IDRV_POPedana)
            rsEti!CodicePedana = fnNotNull(rs!CodicePedana)
            rsEti!DescrizionePedana = fnNotNull(rs!DescrizionePedana)
            rsEti!IDAnagraficaCliente = fnNotNullN(rs!IDCliente)
            rsEti!AnagraficaCliente = fnNotNull(rs!Anagrafica)
            rsEti!IDRV_POAssegnazioneMerce = fnNotNullN(rs!IDRV_POAssegnazioneMerce)
            rsEti!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsEti!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
            rsEti!Articolo = fnNotNull(rs!Articolo)
            rsEti!IDArticoloImballo = fnNotNullN(rs!IDImballoVendita)
            rsEti!CodiceImballo = fnNotNull(rs!CodiceImballoVendita)
            rsEti!Imballo = fnNotNull(rs!ImballoVendita)
            rsEti!Colli = fnNotNullN(rs!Colli)
            rsEti!PesoLordo = fnNotNullN(rs!PesoLordo)
            rsEti!Tara = fnNotNullN(rs!Tara)
            rsEti!PesoNetto = fnNotNullN(rs!PesoNetto)
            rsEti!Pezzi = fnNotNullN(rs!Pezzi)
            rsEti!Qta_UM = fnNotNullN(rs!Qta_UM)
            rsEti!CodiceLotto = fnNotNull(rs!CodiceLottoVendita)
            rsEti!IDAnagraficaSocio = fnNotNullN(rs!IDAnagraficaSocio)
            rsEti!AnagraficaSocio = fnNotNull(rs!AnagraficaSocio)
            rsEti!NomeSocio = fnNotNull(rs!NomeSocio)
            rsEti!CodiceSocio = fnNotNull(rs!CodiceSocio)
            rsEti!IDRV_POCalibro = fnNotNullN(rs!IDRV_POCalibro)
            rsEti!Calibro = fnNotNull(rs!Calibro)
            rsEti!IDRV_POTipoCategoria = fnNotNullN(rs!IDRV_POTipoCategoria)
            rsEti!TipoCategoria = fnNotNull(rs!TipoCategoria)
            rsEti!IDRegione = fnNotNullN(rs!IDRegione)
            rsEti!Regione = fnNotNull(rs!Regione)
            rsEti!IDComune = fnNotNullN(rs!IDComune)
            rsEti!IDNazione = fnNotNullN(rs!IDNazione)
            rsEti!Nazione = fnNotNull(rs!Nazione)
            rsEti!Comune = fnNotNull(rs!Comune)
            rsEti!IDProvincia = fnNotNullN(rs!IDProvincia)
            rsEti!LottoDiConferimento = fnNotNull(rs!LottoDiConferimento)
            rsEti!BNDOO = fnNotNull(GET_BNDOO(TheApp.Branch))
            rsEti!IDLottoDiCampagna = fnNotNullN(rs!IDLottoDiCampagna)
            rsEti!CodiceCertificazioneLotto = CODICE_CERTIFICAZIONE_LOTTO_ETI
            rsEti!DescrizioneCertificazioneLotto = DESCRIZIONE_CERTIFICAZIONE_LOTTO_ETI
            rsEti!ProtocolloCertificazioneLotto = PROTOCOLLO_CERTIFICAZIONE_LOTTO_ETI
            rsEti!EnteCertificatoreLotto = ENTE_CERTIFICAZIONE_LOTTO_ETI
            rsEti!CodiceCertificazioneSocioPred = CODICE_CERTIFICAZIONE_SOCIO_ETI
            rsEti!DescrizioneCertificazioneSocioPred = DESCRIZIONE_CERTIFICAZIONE_SOCIO_ETI
            rsEti!ProtocolloCertificazioneSocioPred = PROTOCOLLO_CERTIFICAZIONE_SOCIO_ETI
            rsEti!EnteCertificatoreSocioPred = ENTE_CERTIFICAZIONE_SOCIO_ETI
            rsEti!LottoCliente = fnNotNull(rs!LottoCliente)
            rsEti!AltreAnnotazioniCliente = fnNotNull(rs!AltreAnnotazioniPerCliente)
            rsEti!IDCategoriaMerceologica = fnNotNullN(rs!IDCategoriaMerceologica)
            rsEti!CategoriaMerceologica = fnNotNull(rs!CategoriaMerceologica)
            rsEti!IDTipoProdotto = fnNotNullN(rs!IDTipoProdotto)
            rsEti!TipoProdotto = fnNotNull(rs!TipoProdotto)
            rsEti!CodiceLottoEntrata = fnNotNull(rs!CodiceLottoEntrata)
            rsEti!CodiceAssociato = fnNotNull(GET_CODICEFORNITORE(TheApp.Branch))
            rsEti!CodiceABarreArticolo = fnNotNull(GET_CODICEABARRE(rs!IDArticolo))
            rsEti!IDVarietaLottoCampagna = LINK_VARIETA_LOTTO_CAMPAGNA
            rsEti!IDFamigliaLottoCampagna = LINK_FAMIGLIA_LOTTO_CAMPAGNA
            rsEti!VarietaLottoCampagna = VARIETA_LOTTO_CAMPAGNA
            rsEti!FamigliaLottoCampagna = FAMIGLIA_LOTTO_CAMPAGNA
        rsEti.Update
        
        AggiornaEtichettePedanaStampate fnNotNullN(rs!IDRV_POPedana)
    rs.MoveNext
    Wend
End If
rs.Close
Set rs = Nothing

rsEti.Close
Set rsEti = Nothing


StampaEtichettePedana
End Sub

Private Sub cmdStampaSingola_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsEti As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim I As Integer
Dim Unita_progresso
Dim NumeroRecord As Long

If Len(Me.cboStampa.Text) = 0 Then
    MsgBox "Selezionare la stampante per le etichette di lavorazione", vbInformation, "Stampa etichette"
    Exit Sub
End If
    
If Me.cboReport.CurrentID = 0 Then
    MsgBox "Selezionare una stampa per etichette di lavorazione", vbInformation, "Stampa etichette"
    Exit Sub
End If
    
If Me.Check1.Value = vbChecked Then
    If Me.cboReportPed.CurrentID = 0 Then
        MsgBox "Selezionare una stampa per etichette di pedana", vbInformation, "Stampa etichette"
        Exit Sub
    End If
    If Len(Me.cboStampantePed.Text) = 0 Then
        MsgBox "Selezionare la stampante per le etichette di pedana", vbInformation, "Stampa etichette"
        Exit Sub
    End If
        
End If
    
    rsGriglia.UpdateBatch
    
    sSQL = "SELECT Count(IDRV_POEtichette) AS TotaleRecord FROM RV_POEtichette "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    sSQL = sSQL & " AND DaStampare=" & fnNormBoolean(1)
    sSQL = sSQL & " AND ColliDaStampare > 0"
    sSQL = sSQL & " AND IDRV_POAssegnazioneMerce=" & fnNotNullN(Me.GridEtichette.AllColumns("IDRV_POAssegnazioneMerce").Value)
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic
    
    If rs.EOF Then
        Unita_progresso = 0
    Else
        If fnNotNullN(rs!TotaleRecord) > 0 Then
            Unita_progresso = Me.ProgressBar1.Max / fnNotNullN(rs!TotaleRecord)
        Else
            Unita_progresso = 0
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    If Unita_progresso = 0 Then
        MsgBox "Non risulta nessuna etichetta da stampare", vbInformation, "Stampa etichette"
        Exit Sub
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    sSQL = "SELECT * FROM RV_POEtichette "
    sSQL = sSQL & "WHERE IDUtente = " & TheApp.IDUser
    sSQL = sSQL & " AND DaStampare = " & fnNormBoolean(1)
    sSQL = sSQL & " AND ColliDaStampare > 0 "
    sSQL = sSQL & " AND IDRV_POAssegnazioneMerce=" & fnNotNullN(Me.GridEtichette.AllColumns("IDRV_POAssegnazioneMerce").Value)

    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic


    sSQL = "DELETE FROM RV_POTMPStampaEtichetteRighe "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    Cn.Execute sSQL
    
    
    Set rsEti = New ADODB.Recordset
    rsEti.Open "RV_POIEStampaEtichette", Cn.InternalConnection, adOpenDynamic, adLockPessimistic
    
    
    
    While Not rs.EOF

        rsEti.Filter = "IDRV_POAssegnazioneMerce = " & fnNotNullN(rs!IDRV_POAssegnazioneMerce)
        
        Set rsNew = New ADODB.Recordset
        rsNew.Open "RV_POTMPStampaEtichetteRighe", Cn.InternalConnection, adOpenDynamic, adLockPessimistic

        Screen.MousePointer = 11
        'For I = 1 To fnNotNullN(rs!ColliDaStampare)
            rsNew.AddNew
                rsNew!IDUtente = TheApp.IDUser
                rsNew!IDAzienda = TheApp.IDFirm
                rsNew!IDFiliale = TheApp.Branch
                rsNew!IDRV_POPedana = fnNotNullN(rsEti!IDRV_POPedana)
                rsNew!CodicePedana = fnNotNull(rsEti!CodicePedana)
                rsNew!DescrizionePedana = fnNotNull(rsEti!DescrizionePedana)
                rsNew!IDAnagraficaCliente = fnNotNullN(rsEti!IDCliente)
                rsNew!AnagraficaCliente = fnNotNull(rsEti!Anagrafica)
                rsNew!IDRV_POAssegnazioneMerce = fnNotNullN(rsEti!IDRV_POAssegnazioneMerce)
                rsNew!IDArticolo = fnNotNullN(rsEti!IDArticolo)
                rsNew!CodiceArticolo = fnNotNull(rsEti!CodiceArticolo)
                rsNew!Articolo = fnNotNull(rsEti!Articolo)
                rsNew!IDArticoloImballo = fnNotNullN(rsEti!IDImballoVendita)
                rsNew!CodiceImballo = fnNotNull(rsEti!CodiceImballoVendita)
                rsNew!Imballo = fnNotNull(rsEti!ImballoVendita)
                rsNew!Colli = fnNotNullN(rsEti!Colli)
                rsNew!PesoLordo = fnNotNullN(rsEti!PesoLordo)
                rsNew!Tara = fnNotNullN(rsEti!Tara)
                rsNew!PesoNetto = fnNotNullN(rsEti!PesoNetto)
                rsNew!Pezzi = fnNotNullN(rsEti!Pezzi)
                rsNew!Qta_UM = fnNotNullN(rsEti!Qta_UM)
                rsNew!CodiceLotto = fnNotNull(rsEti!CodiceLottoVendita)
                rsNew!IDAnagraficaSocio = fnNotNullN(rsEti!IDAnagraficaSocio)
                rsNew!AnagraficaSocio = fnNotNull(rsEti!AnagraficaSocio)
                rsNew!NomeSocio = fnNotNull(rsEti!NomeSocio)
                rsNew!CodiceSocio = fnNotNull(rsEti!CodiceSocio)
                rsNew!IDRV_POCalibro = fnNotNullN(rsEti!IDRV_POCalibro)
                rsNew!Calibro = fnNotNull(rsEti!Calibro)
                rsNew!IDRV_POTipoCategoria = fnNotNullN(rsEti!IDRV_POTipoCategoria)
                rsNew!TipoCategoria = fnNotNull(rsEti!TipoCategoria)
                rsNew!IDRegione = fnNotNullN(rsEti!IDRegione)
                rsNew!Regione = fnNotNull(rsEti!Regione)
                rsNew!IDNazione = fnNotNullN(rsEti!IDNazione)
                rsNew!Nazione = fnNotNull(rsEti!Nazione)
                rsNew!IDComune = fnNotNullN(rsEti!IDComune)
                rsNew!Comune = fnNotNull(rsEti!Comune)
                rsNew!IDProvincia = fnNotNullN(rsEti!IDProvincia)
                rsNew!LottoDiConferimento = fnNotNull(rsEti!LottoDiConferimento)
                rsNew!BNDOO = fnNotNull(GET_BNDOO(TheApp.Branch))
                rsNew!IDLottoDiCampagna = fnNotNullN(rsEti!IDLottoDiCampagna)
                rsNew!CodiceCertificazioneLotto = CODICE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!DescrizioneCertificazioneLotto = DESCRIZIONE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!ProtocolloCertificazioneLotto = PROTOCOLLO_CERTIFICAZIONE_LOTTO_ETI
                rsNew!EnteCertificatoreLotto = ENTE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!CodiceCertificazioneSocioPred = CODICE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!DescrizioneCertificazioneSocioPred = DESCRIZIONE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!ProtocolloCertificazioneSocioPred = PROTOCOLLO_CERTIFICAZIONE_SOCIO_ETI
                rsNew!EnteCertificatoreSocioPred = ENTE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!LottoCliente = fnNotNull(rsEti!LottoCliente)
                rsNew!AltreAnnotazioniCliente = fnNotNull(rsEti!AltreAnnotazioniPerCliente)
                rsNew!IDCategoriaMerceologica = fnNotNullN(rsEti!IDCategoriaMerceologica)
                rsNew!CategoriaMerceologica = fnNotNull(rsEti!CategoriaMerceologica)
                rsNew!IDTipoProdotto = fnNotNullN(rsEti!IDTipoProdotto)
                rsNew!TipoProdotto = fnNotNull(rsEti!TipoProdotto)
                rsNew!CodiceLottoEntrata = fnNotNull(rsEti!CodiceLottoEntrata)
                rsNew!CodiceAssociato = fnNotNull(GET_CODICEFORNITORE(TheApp.Branch))
                rsNew!CodiceABarreArticolo = fnNotNull(GET_CODICEABARRE(rsEti!IDArticolo))
                rsNew!IDVarietaLottoCampagna = LINK_VARIETA_LOTTO_CAMPAGNA
                rsNew!IDFamigliaLottoCampagna = LINK_FAMIGLIA_LOTTO_CAMPAGNA
                rsNew!VarietaLottoCampagna = VARIETA_LOTTO_CAMPAGNA
                rsNew!FamigliaLottoCampagna = FAMIGLIA_LOTTO_CAMPAGNA
            rsNew.Update
                                                
            'Me.lblInfoEtichetteLav.Caption = "Numero etichette elaborate " & I & " di " & fnNotNullN(rs!ColliDaStampare)
            DoEvents
        'Next
        
        rsNew.Close
        Set rsNew = Nothing
        Screen.MousePointer = 0
        
        AggiornaEtichetteStampate fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!ColliDaStampare)
        
        StampaEtichette fnNotNullN(rs!ColliDaStampare)
        
        If Me.Check1.Value = vbChecked Then
            CODICE_PEDANA = fnNotNull(rsEti!CodicePedana)
            cmdStampaPedana_Click
        End If
        
        
        
        rsGriglia!ColliStampati = rsGriglia!ColliStampati + rsGriglia!ColliDaStampare
        rsGriglia.Update
        Me.GridEtichette.Refresh
        
        
        
        If (Me.ProgressBar1.Value + Unita_progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_progresso
        End If
        
        DoEvents
        'MsgBox "StampaEtichette"
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
        
    rsEti.Close
    Set rsEti = Nothing
        
    
    
    
End Sub
    
    
    

Private Sub Command1_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsEti As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim I As Integer
Dim Unita_progresso

If Len(Me.cboStampa.Text) = 0 Then
    MsgBox "Selezionare la stampante per le etichette di lavorazione", vbInformation, "Stampa etichette"
    Exit Sub
End If
    
If Me.cboReport.CurrentID = 0 Then
    MsgBox "Selezionare una stampa per etichette di lavorazione", vbInformation, "Stampa etichette"
    Exit Sub
End If
    
If Me.Check1.Value = vbChecked Then
    If Me.cboReportPed.CurrentID = 0 Then
        MsgBox "Selezionare una stampa per etichette di pedana", vbInformation, "Stampa etichette"
        Exit Sub
    End If
    If Len(Me.cboStampantePed.Text) = 0 Then
        MsgBox "Selezionare la stampante per le etichette di pedana", vbInformation, "Stampa etichette"
        Exit Sub
    End If
        
End If
    
    rsGriglia.Update
    
    sSQL = "SELECT Count(IDRV_POEtichette) AS TotaleRecord FROM RV_POEtichette "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    sSQL = sSQL & " AND DaStampare=" & fnNormBoolean(1)
    sSQL = sSQL & " AND ColliDaStampare > 0"
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic
    
    If rs.EOF Then
        Unita_progresso = 0
    Else
        If fnNotNullN(rs!TotaleRecord) > 0 Then
            Unita_progresso = Me.ProgressBar1.Max / fnNotNullN(rs!TotaleRecord)
        Else
            Unita_progresso = 0
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    If Unita_progresso = 0 Then
        MsgBox "Non risulta nessuna etichetta da stampare", vbInformation, "Stampa etichette"
        Exit Sub
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    sSQL = "SELECT * FROM RV_POEtichette "
    sSQL = sSQL & "WHERE IDUtente = " & TheApp.IDUser
    sSQL = sSQL & " AND DaStampare = " & fnNormBoolean(1)
    sSQL = sSQL & " AND ColliDaStampare > 0"
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic


    
    
    Set rsEti = New ADODB.Recordset
    rsEti.Open "RV_POIEStampaEtichette", Cn.InternalConnection, adOpenDynamic, adLockPessimistic
    
    
    
    While Not rs.EOF
        sSQL = "DELETE FROM RV_POTMPStampaEtichetteRighe "
        sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
        Cn.Execute sSQL

        rsEti.Filter = "IDRV_POAssegnazioneMerce = " & fnNotNullN(rs!IDRV_POAssegnazioneMerce)
        
        Set rsNew = New ADODB.Recordset
        rsNew.Open "RV_POTMPStampaEtichetteRighe", Cn.InternalConnection, adOpenDynamic, adLockPessimistic

        Screen.MousePointer = 11
        For I = 1 To fnNotNullN(rs!ColliDaStampare)
            rsNew.AddNew
                rsNew!IDUtente = TheApp.IDUser
                rsNew!IDAzienda = TheApp.IDFirm
                rsNew!IDFiliale = TheApp.Branch
                rsNew!IDRV_POPedana = fnNotNullN(rsEti!IDRV_POPedana)
                rsNew!CodicePedana = fnNotNull(rsEti!CodicePedana)
                rsNew!DescrizionePedana = fnNotNull(rsEti!DescrizionePedana)
                rsNew!IDAnagraficaCliente = fnNotNullN(rsEti!IDCliente)
                rsNew!AnagraficaCliente = fnNotNull(rsEti!Anagrafica)
                rsNew!IDRV_POAssegnazioneMerce = fnNotNullN(rsEti!IDRV_POAssegnazioneMerce)
                rsNew!IDArticolo = fnNotNullN(rsEti!IDArticolo)
                rsNew!CodiceArticolo = fnNotNull(rsEti!CodiceArticolo)
                rsNew!Articolo = fnNotNull(rsEti!Articolo)
                rsNew!IDArticoloImballo = fnNotNullN(rsEti!IDImballoVendita)
                rsNew!CodiceImballo = fnNotNull(rsEti!CodiceImballoVendita)
                rsNew!Imballo = fnNotNull(rsEti!ImballoVendita)
                rsNew!Colli = fnNotNullN(rsEti!Colli)
                rsNew!PesoLordo = fnNotNullN(rsEti!PesoLordo)
                rsNew!Tara = fnNotNullN(rsEti!Tara)
                rsNew!PesoNetto = fnNotNullN(rsEti!PesoNetto)
                rsNew!Pezzi = fnNotNullN(rsEti!Pezzi)
                rsNew!Qta_UM = fnNotNullN(rsEti!Qta_UM)
                rsNew!CodiceLotto = fnNotNull(rsEti!CodiceLottoVendita)
                rsNew!IDAnagraficaSocio = fnNotNullN(rsEti!IDAnagraficaSocio)
                rsNew!AnagraficaSocio = fnNotNull(rsEti!AnagraficaSocio)
                rsNew!NomeSocio = fnNotNull(rsEti!NomeSocio)
                rsNew!CodiceSocio = fnNotNull(rsEti!CodiceSocio)
                rsNew!IDRV_POCalibro = fnNotNullN(rsEti!IDRV_POCalibro)
                rsNew!Calibro = fnNotNull(rsEti!Calibro)
                rsNew!IDRV_POTipoCategoria = fnNotNullN(rsEti!IDRV_POTipoCategoria)
                rsNew!TipoCategoria = fnNotNull(rsEti!TipoCategoria)
                rsNew!IDRegione = fnNotNullN(rsEti!IDRegione)
                rsNew!Regione = fnNotNull(rsEti!Regione)
                rsNew!IDNazione = fnNotNullN(rsEti!IDNazione)
                rsNew!Nazione = fnNotNull(rsEti!Nazione)
                rsNew!IDComune = fnNotNullN(rsEti!IDComune)
                rsNew!Comune = fnNotNull(rsEti!Comune)
                rsNew!IDProvincia = fnNotNullN(rsEti!IDProvincia)
                rsNew!LottoDiConferimento = fnNotNull(rsEti!LottoDiConferimento)
                rsNew!BNDOO = fnNotNull(GET_BNDOO(TheApp.Branch))
                rsNew!IDLottoDiCampagna = fnNotNullN(rsEti!IDLottoDiCampagna)
                rsNew!CodiceCertificazioneLotto = CODICE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!DescrizioneCertificazioneLotto = DESCRIZIONE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!ProtocolloCertificazioneLotto = PROTOCOLLO_CERTIFICAZIONE_LOTTO_ETI
                rsNew!EnteCertificatoreLotto = ENTE_CERTIFICAZIONE_LOTTO_ETI
                rsNew!CodiceCertificazioneSocioPred = CODICE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!DescrizioneCertificazioneSocioPred = DESCRIZIONE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!ProtocolloCertificazioneSocioPred = PROTOCOLLO_CERTIFICAZIONE_SOCIO_ETI
                rsNew!EnteCertificatoreSocioPred = ENTE_CERTIFICAZIONE_SOCIO_ETI
                rsNew!LottoCliente = fnNotNull(rsEti!LottoCliente)
                rsNew!AltreAnnotazioniCliente = fnNotNull(rsEti!AltreAnnotazioniPerCliente)
                rsNew!IDCategoriaMerceologica = fnNotNullN(rsEti!IDCategoriaMerceologica)
                rsNew!CategoriaMerceologica = fnNotNull(rsEti!CategoriaMerceologica)
                rsNew!IDTipoProdotto = fnNotNullN(rsEti!IDTipoProdotto)
                rsNew!TipoProdotto = fnNotNull(rsEti!TipoProdotto)
                rsNew!CodiceLottoEntrata = fnNotNull(rsEti!CodiceLottoEntrata)
                rsNew!CodiceAssociato = fnNotNull(GET_CODICEFORNITORE(TheApp.Branch))
                rsNew!CodiceABarreArticolo = fnNotNull(GET_CODICEABARRE(rsEti!IDArticolo))
                rsNew!IDVarietaLottoCampagna = LINK_VARIETA_LOTTO_CAMPAGNA
                rsNew!IDFamigliaLottoCampagna = LINK_FAMIGLIA_LOTTO_CAMPAGNA
                rsNew!VarietaLottoCampagna = VARIETA_LOTTO_CAMPAGNA
                rsNew!FamigliaLottoCampagna = FAMIGLIA_LOTTO_CAMPAGNA
            
            rsNew.Update
                                                
            Me.lblInfoEtichetteLav.Caption = "Numero etichette elaborate " & I & " di " & fnNotNullN(rs!ColliDaStampare)
            DoEvents
        Next
        rsNew.Close
        Set rsNew = Nothing
        Screen.MousePointer = 0
        AggiornaEtichetteStampate fnNotNullN(rs!IDRV_POAssegnazioneMerce), fnNotNullN(rs!ColliDaStampare)
        
        StampaEtichette fnNotNullN(rs!ColliDaStampare)
        If Me.Check1.Value = vbChecked Then
            CODICE_PEDANA = fnNotNull(rsEti!CodicePedana)
            cmdStampaPedana_Click
        End If
        
        If (Me.ProgressBar1.Value + Unita_progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_progresso
        End If
        
        DoEvents
        
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
        
    rsEti.Close
    Set rsEti = Nothing
        
        
    
End Sub

Private Sub Form_Activate()
    Me.GridEtichette.SetFocus
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)


'RecuperaDati

fncGriglia
Link_TipoOggettoLocal = fncIDTipoOggettoPrg("RV_POEtichetteLavorazione")
fncStampanti
fncStampantiPedana
fncReport
fncReportPedana


Me.cboReport.WriteOn fnDefaultReport
Me.cboReportPed.WriteOn fnDefaultReportPedana

    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With
    

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    

        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If
End Sub
Public Sub RecuperaDati()

End Sub


Public Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    

    sSQL = "SELECT * FROM RV_POEtichette WHERE IDUtente=" & TheApp.IDUser
    sSQL = sSQL & "ORDER BY CodiceSocio, DataLavorazione DESC"
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            'Set rsEvent = rsGriglia2.Data
    
        
    
        With Me.GridEtichette
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ForeColor = vbYellow
            .ColumnsHeader.Clear
                    .ColumnsHeader.Add "IDRV_POEtichette", "ID", dgInteger, False, 500, dgAlignleft
                    Set cl = .ColumnsHeader.Add("DaStampare", "Da stampare", dgBoolean, True, 1200, dgAlignleft)
                        cl.Editable = True
                    Set cl = .ColumnsHeader.Add("DataLavorazione", "Data lav.", dgDate, True, 1100, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("CodiceSocio", "Cod. Socio", dgchar, True, 1000, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("DataConferimento", "Data Conf.", dgDate, True, 1100, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 800, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("ColliStampati", "Stampate", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("ColliDaStampare", "Da stampare", dgDouble, True, 1200, dgAlignRight)
                        cl.Editable = True
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                        cl.ForeColor = vbBlack
                        cl.BackColor = vbRed
                        
                    Set cl = .ColumnsHeader.Add("IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("CodiceArticolo", "Codice articolo", dgchar, True, 2000, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("Articolo", "Articolo", dgchar, False, 2500, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("CodiceLotto", "Codice lotto", dgchar, True, 2000, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("CodicePedana", "Pedana", dgchar, True, 1300, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("Cliente", "Cliente", dgchar, True, 2000, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("NumeroOrdine", "N° Ord.", dgInteger, True, 800, dgAlignleft)
                        cl.ForeColor = vbYellow
                    Set cl = .ColumnsHeader.Add("DataOrdine", "Data Ord.", dgDate, True, 1100, dgAlignleft)
                        cl.ForeColor = vbYellow
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
        
    rsGriglia.Close
    Set rsGriglia = Nothing
        
End Sub

Private Sub GridEtichette_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GridEtichette.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Dastampare").Value), 2
        End If
    End If

End Sub

Private Sub GridEtichette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GridEtichette.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GridEtichette.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GridEtichette.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("DaStampare").Value), 2
            End If
        End If
    End If

End Sub
Private Sub StampaEtichettePedana()
On Error GoTo ERR_CmdStampaEtichette_Click
Dim IDReport As Long
Dim Stampante As String
Dim Pr As Printer
Dim NomeStampantePredefinita As String
Dim IDTipoOggettoPrg As Long
        
        IDTipoOggettoPrg = fncIDTipoOggettoPrg("RV_POEtichetteLavorazione")
        
        NomeStampantePredefinita = Printer.DeviceName
            
        Set oReport = New dmtReportLib.dmtReport
            Set oReport.Connection = Cn
            If MenuOptions.DBType = 1 Then
                'parametri di accesso al database ACCESS
                oReport.Password = "dmt192981046"
                oReport.User = "admin"
            Else
                'parametri di accesso al database SQL Server
                oReport.Password = TheApp.Password
                oReport.User = TheApp.User
            End If
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = IDTipoOggettoPrg
                
            oReport.Where = "IDUtente=" & TheApp.IDUser
            
            IDReport = fncTrovaReport(Me.cboReportPed.Text, IDTipoOggettoPrg)
            
            If IDReport > 0 Then
                fncImpostaDefaultReport IDReport, IDTipoOggettoPrg
                
                    
                Stampante = Me.cboStampantePed.Text
                If Stampante <> NomeStampantePredefinita Then
                    For Each Pr In Printers
                        If Pr.DeviceName = Stampante Then
                            If WinNTSetDefaultPrinter(Pr.DeviceName) = True Then
                                Exit For
                            Else
                                MsgBox "Problemi con la stampante " & Me.cboStampantePed.Text
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                oReport.Copies = 1
                oReport.DoPrint Stampante
                If Stampante <> NomeStampantePredefinita Then
                    WinNTSetDefaultPrinter NomeStampantePredefinita
                End If
                
            Else
                
                MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
            
            End If
          
            
      


Exit Sub
ERR_CmdStampaEtichette_Click:
MsgBox Err.Description, vbCritical, "Errore nella stampa delle etichette"
If Stampante <> NomeStampantePredefinita Then
    WinNTSetDefaultPrinter NomeStampantePredefinita
End If

  
    
End Sub

Private Sub StampaEtichette(NumeroEtichette As Long)
On Error GoTo ERR_CmdStampaEtichette_Click
Dim IDReport As Long
Dim Stampante As String
Dim Pr As Printer
Dim NomeStampantePredefinita As String
Dim IDTipoOggettoPrg As Long




        IDTipoOggettoPrg = fncIDTipoOggettoPrg("RV_POEtichetteLavorazione")
        
        NomeStampantePredefinita = Printer.DeviceName
            
        Set oReport = New dmtReportLib.dmtReport
            Set oReport.Connection = Cn
            If MenuOptions.DBType = 1 Then
                'parametri di accesso al database ACCESS
                oReport.Password = "dmt192981046"
                oReport.User = "admin"
            Else
                'parametri di accesso al database SQL Server
                oReport.Password = TheApp.Password
                oReport.User = TheApp.User
            End If
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = IDTipoOggettoPrg
                
            oReport.Where = "IDUtente=" & TheApp.IDUser
            
            IDReport = fncTrovaReport(Me.cboReport.Text, IDTipoOggettoPrg)
            
            If IDReport > 0 Then
                fncImpostaDefaultReport IDReport, IDTipoOggettoPrg
                
                    
                Stampante = Me.cboStampa.Text
                If Stampante <> NomeStampantePredefinita Then
                    For Each Pr In Printers
                        If Pr.DeviceName = Stampante Then
                            If WinNTSetDefaultPrinter(Pr.DeviceName) = True Then
                                Exit For
                            Else
                                MsgBox "Problemi con la stampante " & Me.cboStampa.Text
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                
                oReport.Copies = NumeroEtichette
                
                oReport.DoPrint Stampante
                
                If Stampante <> NomeStampantePredefinita Then
                    WinNTSetDefaultPrinter NomeStampantePredefinita
                End If
                
            Else
                
                MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
            
            End If
          
            
      


Exit Sub
ERR_CmdStampaEtichette_Click:
MsgBox Err.Description, vbCritical, "Errore nella stampa delle etichette"
If Stampante <> NomeStampantePredefinita Then
    WinNTSetDefaultPrinter NomeStampantePredefinita
End If

  
    
End Sub

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

Private Function fncTrovaReport(IDReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(IDReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = fnNotNullN(rs!IDReportTipoOggetto)
Else
    fncTrovaReport = 1
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_fncTrovaReport:
    MsgBox Err.Description, vbCritical, "Trova report per stampa"
    fncTrovaReport = 0

End Function
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto & " AND IDFiliale = " & TheApp.Branch
    
    Cn.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function GET_CodiceSocio(IDAnagrafica As Long, IDAzienda As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Codice FROM Fornitore WHERE "
sSQL = sSQL & "IDAzienda=" & IDAzienda & " AND "
sSQL = sSQL & "IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_CodiceSocio = ""
Else
    GET_CodiceSocio = fnNotNull(rs!Codice)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function AggiornaEtichetteStampate(IDLavorazione As Long, ColliDaStampare As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POEtichetteLavorazione WHERE IDRV_POLavorazione=" & IDLavorazione

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDRV_POLavorazione = IDLavorazione
    rs!ColliStampati = ColliDaStampare
    rs!DataUltimaStampa = Date
    rs!OraUltimaStampa = time
    rs!NumeroUltimaStampa = ColliDaStampare
Else
    rs!ColliStampati = fnNotNullN(rs!ColliStampati) + ColliDaStampare
    rs!DataUltimaStampa = Date
    rs!OraUltimaStampa = time
    rs!NumeroUltimaStampa = ColliDaStampare

End If
    
    rs.Update
        
    
rs.Close
Set rs = Nothing
End Function
Private Function AggiornaEtichettePedanaStampate(IDPedana As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POEtichettePedana WHERE IDRV_POPedana=" & IDPedana

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDRV_POPedana = IDPedana
    rs!DataUltimaStampa = Date
    rs!OraUltimaStampa = time
Else
    rs!DataUltimaStampa = Date
    rs!OraUltimaStampa = time
End If
    
    rs.Update
        
    
rs.Close
Set rs = Nothing
End Function

Private Sub GET_ColliStampati(IDLavorazione As Long, ColliLavorazione As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ColliStampati FROM RV_POEtichetteLavorazione WHERE IDRV_POlavorazione=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    ColliStampati = 0
   
Else
    ColliStampati = fnNotNullN(rs!ColliStampati)
   
End If

ColliDaStampare = ColliLavorazione - ColliStampati

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub fncStampanti()
Dim prn As Printer
For Each prn In Printers
    Me.cboStampa.AddItem prn.DeviceName
Next


End Sub
Private Sub fncStampantiPedana()
Dim prn As Printer
For Each prn In Printers
    Me.cboStampantePed.AddItem prn.DeviceName
Next


End Sub

Private Sub fncReport()
Dim sSQL As String
    With Me.cboReport
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDReportTipoOggetto"
        .DisplayField = "ReportTipoOggetto"
        sSQL = "SELECT * FROM RV_POEtichetteDefault INNER JOIN "
        sSQL = sSQL & "ReportTipoOggetto ON RV_POEtichetteDefault.IDReportPerTipoOggetto = ReportTipoOggetto.IDReportTipoOggetto "
        sSQL = sSQL & "WHERE IDRV_POTipoEtichetta=1"
        sSQL = sSQL & " AND NomeComputer=" & fnNormString(NOME_COMPUTER)
        sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
        .SQL = sSQL
        
        .Fill
    End With
End Sub
Private Sub fncReportPedana()
Dim sSQL As String
    With Me.cboReportPed
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDReportTipoOggetto"
        .DisplayField = "ReportTipoOggetto"
        sSQL = "SELECT * FROM RV_POEtichetteDefault INNER JOIN "
        sSQL = sSQL & "ReportTipoOggetto ON RV_POEtichetteDefault.IDReportPerTipoOggetto = ReportTipoOggetto.IDReportTipoOggetto "
        sSQL = sSQL & "WHERE IDRV_POTipoEtichetta=2"
        sSQL = sSQL & " AND NomeComputer=" & fnNormString(NOME_COMPUTER)
        sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
        .SQL = sSQL
        
        .Fill
    End With
End Sub

Private Function fnDefaultReport() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDReportPerTipoOggetto, Stampante FROM RV_POEtichetteDefault "
sSQL = sSQL & "WHERE Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND NomeComputer=" & fnNormString(NOME_COMPUTER)
sSQL = sSQL & " AND IDRV_POTipoEtichetta=1"
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fnDefaultReport = fnNotNullN(rs!IDReportPerTipoOggetto)
    Me.cboStampa.Text = fnNotNull(rs!Stampante)
Else
    fnDefaultReport = 0
    Me.cboStampa.Text = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function fnDefaultReportPedana() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDReportPerTipoOggetto, Stampante FROM RV_POEtichetteDefault "
sSQL = sSQL & "WHERE Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND NomeComputer=" & fnNormString(NOME_COMPUTER)
sSQL = sSQL & " AND IDRV_POTipoEtichetta=2"
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fnDefaultReportPedana = fnNotNullN(rs!IDReportPerTipoOggetto)
    Me.cboStampantePed.Text = fnNotNull(rs!Stampante)
Else
    fnDefaultReportPedana = 0
    Me.cboStampantePed.Text = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function Get_TrovaStampante() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT Stampante FROM RV_POEtichetteDefault "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDReportPerTipoOggetto=" & Me.cboReport.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Get_TrovaStampante = fnNotNull(rs!Stampante)
    NuovoRecordPerStampa = 0
Else
    Get_TrovaStampante = ""
    NuovoRecordPerStampa = 1
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AggiornaDefaultStampantePerUtente()
Dim sSQL As String
If NuovoRecordPerStampa = 0 Then
    sSQL = "UPDATE RV_POEtichetteDefault SET "
    sSQL = sSQL & "Stampante=" & fnNormString(Me.cboStampa.Text)
    sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser & " AND "
    sSQL = sSQL & "IDReportPerTipoOggetto=" & Me.cboReport.CurrentID
Else
    sSQL = "INSERT INTO RV_POEtichetteDefault ("
    sSQL = sSQL & "IDRV_POEtichetteDefault, IDUtente, IDReportPerTipoOggetto, Stampante) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_POEtichetteDefault", "IDRV_POEtichetteDefault") & ", "
    sSQL = sSQL & TheApp.Branch & ", "
    sSQL = sSQL & Me.cboReport.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cboStampa.Text) & ")"

End If
    Cn.Execute sSQL
End Sub
Private Function GET_BNDOO(IDFiliale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT BNDOO FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_BNDOO = ""
Else
    GET_BNDOO = Trim(fnNotNull(rs!BNDOO))
End If
    
    
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICEFORNITORE(IDFiliale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
sSQL = "SELECT CodiceAssociato FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICEFORNITORE = ""
Else
    GET_CODICEFORNITORE = Trim(fnNotNull(rs!CodiceAssociato))
End If
    
    
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CERTIFICAZIONE_SOCIO(IDSocio As Long, NomeCampo As String) As String
On Error GoTo ERR_GET_CERTIFICAZIONE_SOCIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo & " "
sSQL = sSQL & "FROM RV_PO01_CertificazioneSocio INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON "
sSQL = sSQL & "RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_EnteCertificazione ON RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione = RV_PO01_EnteCertificazione.IDRV_PO01_EnteCertificazione "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDSocio
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CERTIFICAZIONE_SOCIO = ""
Else
    GET_CERTIFICAZIONE_SOCIO = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function


ERR_GET_CERTIFICAZIONE_SOCIO:
    GET_CERTIFICAZIONE_SOCIO = ""
End Function

Private Function GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(IDLottoDiCampagna As Long, NomeCampo As String) As String
On Error GoTo ERR_GET_CERTIFICAZIONE_LOTTO_CAMPAGNA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PO01_LottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_CertificazioneSocio ON "
sSQL = sSQL & "RV_PO01_LottoCampagna.IDRV_PO01_CertificazioneSocio = RV_PO01_CertificazioneSocio.IDRV_PO01_CertificazioneSocio INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON "
sSQL = sSQL & "RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_EnteCertificazione ON RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione = RV_PO01_EnteCertificazione.IDRV_PO01_EnteCertificazione "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoDiCampagna

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CERTIFICAZIONE_LOTTO_CAMPAGNA = ""
Else
    GET_CERTIFICAZIONE_LOTTO_CAMPAGNA = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function


ERR_GET_CERTIFICAZIONE_LOTTO_CAMPAGNA:
    GET_CERTIFICAZIONE_LOTTO_CAMPAGNA = ""
End Function

Private Sub GridEtichette_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POEtichetteLavorazione "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & fnNotNullN(Me.GridEtichette.AllColumns("IDRV_POAssegnazioneMerce").Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtDataOraUltimaStampa.Text = "NESSUNA ETICHETTA STAMPATA"
    Me.txtNumeroUltimaStampa.Text = "ETICHETTE ULTIMA STAMPA: 0"
Else
    Me.txtDataOraUltimaStampa.Text = "DATA E ORA ULTIMA STAMPA: " & fnNotNull(rs!DataUltimaStampa) & "  " & fnNotNull(rs!OraUltimaStampa)
    Me.txtNumeroUltimaStampa.Text = "ETICHETTE ULTIMA STAMPA: " & fnNotNullN(rs!NumeroUltimaStampa)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.lblPedana(0).Caption = "PEDANA: " & fnNotNull(Me.GridEtichette.AllColumns("CodicePedana").Value)
CODICE_PEDANA = fnNotNull(Me.GridEtichette.AllColumns("CodicePedana").Value)
sSQL = "SELECT * FROM RV_POEtichettePedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & fnNotNullN(Me.GridEtichette.AllColumns("IDRV_POPedana").Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtUltimaStampaPedana.Text = "NESSUNA ETICHETTA STAMPATA"
Else
    Me.txtUltimaStampaPedana.Text = "DATA E ORA ULTIMA STAMPA: " & fnNotNull(rs!DataUltimaStampa) & "  " & fnNotNull(rs!OraUltimaStampa)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function GET_COLLI_STAMPATI(IDLavorazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POEtichetteLavorazione "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_COLLI_STAMPATI = 0
Else
    GET_COLLI_STAMPATI = fnNotNullN(rs!ColliStampati)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function WinNTSetDefaultPrinter(ByRef DeviceName As String) As Boolean
    Dim Buffer As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim R As Long
   
    If DeviceName <> "" Then
        'Preleva le informazioni sulla stampante desiderata dal file WIN.INI file.
        Buffer = Space(1024)
        R = GetProfileString("PrinterPorts", DeviceName, "", Buffer, Len(Buffer))
       
        'Esegue il parsing della stringa ricavata dalla chiamata precedente...
        Call GetDriverAndPort(Buffer, DriverName, PrinterPort)
        If DriverName <> "" And PrinterPort <> "" Then
            WinNTSetDefaultPrinter = SetDefaultPrinter(DeviceName, DriverName, PrinterPort)
        Else
            WinNTSetDefaultPrinter = False
        End If
    End If
End Function
Private Sub GetDriverAndPort(ByVal Buffer As String, ByRef DriverName As String, ByRef PrinterPort As String)
    Dim iDriver As Integer
    Dim iPort As Integer
   
    DriverName = ""
    PrinterPort = ""
 
    'Il nome del driver è la prima stringa delimitata dalla ","
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then
        DriverName = Left(Buffer, iDriver - 1)
        'Il numero di porta è tra le due ","
        iPort = InStr(iDriver + 1, Buffer, ",")
        If iPort > 0 Then
            PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
        End If
    End If
End Sub
Private Function SetDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal PrinterPort As String) As Boolean
    Dim DeviceLine As String
    Dim R As Long
    Dim l As Long
   
    DeviceLine = DeviceName & "," & DriverName & "," & PrinterPort
    'Memorizza le informazioni della stampante nella sezione  [WINDOWS] del file Win.ini alla voce DEVICE
    R = WriteProfileString("windows", "Device", DeviceLine)
   
    If R Then
        'Avverte tutte le applicazioni attive che il file .INI è cambiato:
        l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
        SetDefaultPrinter = True
    Else
        SetDefaultPrinter = False
    End If
End Function

Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Function GET_CODICEABARRE(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceBarre FROM ArticoloPerTipoCodiceBarre "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDTipoCodiceBarre=13"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICEABARRE = ""
Else
    GET_CODICEABARRE = FF_EAN13(fnNotNull(Trim(rs!CodiceBarre)))
End If

rs.CloseResultset
Set rs = Nothing
End Function
