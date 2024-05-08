VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProtocollo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protocolli di certificazione "
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProtocollo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Certificazioni per socio"
      TabPicture(0)   =   "frmProtocollo.frx":4781A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GrigliaCertSocio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Certificazioni per famiglia prodotto"
      TabPicture(1)   =   "frmProtocollo.frx":47836
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrigliaCertFamiglia"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Certificazione azienda"
      TabPicture(2)   =   "frmProtocollo.frx":47852
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GrigliaCertAzienda"
      Tab(2).ControlCount=   1
      Begin DmtGridCtl.DmtGrid GrigliaCertSocio 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7646
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
      Begin DmtGridCtl.DmtGrid GrigliaCertFamiglia 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7646
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
      Begin DmtGridCtl.DmtGrid GrigliaCertAzienda 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7646
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
End
Attribute VB_Name = "frmProtocollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGrigliaSocio As ADODB.Recordset
Private rsGrigliaFamiglia As ADODB.Recordset
Private rsGrigliaAzienda As ADODB.Recordset

Private Sub GET_GRIGLIA_CERT_SOCIO()
On Error GoTo ERR_GET_GRIGLIA_CERT_SOCIO
Dim OLD_Cursor As Long
Dim cl As dgColumnHeader
Dim sSQL As String

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_PO01_CertificazioneSocio.IDRV_PO01_CertificazioneSocio, RV_PO01_CertificazioneSocio.IDAnagrafica, "
    sSQL = sSQL & "RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione, RV_PO01_CertificazioneSocio.ProtocolloCertificazione,"
    sSQL = sSQL & "RV_PO01_CertificazioneSocio.Predefinito, RV_PO01_CertificazioneSocio.CertificatoDal, RV_PO01_CertificazioneSocio.CertificatoAl,"
    sSQL = sSQL & "RV_PO01_Certificazione.CodiceCertificazione, RV_PO01_Certificazione.DescrizioneCertificazione, RV_PO01_EnteCertificazione.EnteCertificazione, "
    sSQL = sSQL & "RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione "
    sSQL = sSQL & "FROM RV_PO01_EnteCertificazione RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Certificazione ON "
    sSQL = sSQL & "RV_PO01_EnteCertificazione.IDRV_PO01_EnteCertificazione = RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_CertificazioneSocio ON RV_PO01_Certificazione.IDRV_PO01_Certificazione = RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione "
    sSQL = sSQL & "WHERE RV_PO01_CertificazioneSocio.IDAnagrafica=" & frmMain.CDSocio.KeyFieldID
    If frmMain.CDCertificazione.KeyFieldID > 0 Then
        sSQL = sSQL & " AND RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione=" & frmMain.CDCertificazione.KeyFieldID
    End If
    
    
    Set rsGrigliaSocio = New ADODB.Recordset
    
    rsGrigliaSocio.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaCertSocio
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDRV_PO01_CertificazioneSocio", "IDRV_PO01_CertificazioneSocio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_Certificazione", "IDRV_PO01_Certificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "ProtocolloCertificazione", "Protocollo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "CodiceCertificazione", "Codice", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "DescrizioneCertificazione", "Descrizione", dgchar, True, 1800, dgAlignleft
                .ColumnsHeader.Add "Predefinito", "Predefinito", dgBoolean, True, 1800, dgAligncenter
                .ColumnsHeader.Add "CertificatoDal", "Certificato dal", dgDate, False, 1800, dgAligncenter
                .ColumnsHeader.Add "CertificatoAl", "Certificato al", dgDate, False, 1800, dgAligncenter
                .ColumnsHeader.Add "IDRV_PO01_EnteCertificazione", "IDRV_PO01_EnteCertificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "EnteCertificazione", "Ente certificatore", dgchar, False, 2000, dgAlignleft
                                
                
        Set .Recordset = rsGrigliaSocio
        .LoadUserSettings
        .Refresh
    End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA_CERT_SOCIO:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_CERT_SOCIO"
End Sub
Private Sub GET_GRIGLIA_CERT_FAMIGLIA()
On Error GoTo ERR_GET_GRIGLIA_CERT_FAMIGLIA
Dim OLD_Cursor As Long
Dim cl As dgColumnHeader
Dim sSQL As String

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_PO01_Certificazione.CodiceCertificazione, RV_PO01_Certificazione.DescrizioneCertificazione, RV_PO01_EnteCertificazione.EnteCertificazione, "
    sSQL = sSQL & "RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione, RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_CertificazioneSocioFamiglia,"
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.IDAnagrafica, RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_FamigliaProdotti,"
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_Certificazione, RV_PO01_CertificazioneSocioFamiglia.ProtocolloCertificazione,"
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.Predefinito, RV_PO01_CertificazioneSocioFamiglia.CertificatoDal,"
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.CertificatoAl , RV_PO01_FamigliaProdotti.FamigliaProdotti "
    sSQL = sSQL & "FROM RV_PO01_CertificazioneSocioFamiglia LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_FamigliaProdotti ON "
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Certificazione ON "
    sSQL = sSQL & "RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_EnteCertificazione ON RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione = RV_PO01_EnteCertificazione.IDRV_PO01_EnteCertificazione "
    sSQL = sSQL & "WHERE RV_PO01_CertificazioneSocioFamiglia.IDAnagrafica=" & frmMain.CDSocio.KeyFieldID
    sSQL = sSQL & " AND RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_FamigliaProdotti=" & frmMain.cboFamigliaProdotti.CurrentID
    If frmMain.CDCertificazione.KeyFieldID > 0 Then
        sSQL = sSQL & " AND RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_Certificazione=" & frmMain.CDCertificazione.KeyFieldID
    End If
    
    
    Set rsGrigliaFamiglia = New ADODB.Recordset
    
    rsGrigliaFamiglia.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaCertFamiglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDRV_PO01_CertificazioneSocio", "IDRV_PO01_CertificazioneSocio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_FamigliaProdotti", "IDRV_PO01_FamigliaProdotti", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "FamigliaProdotti", "Famiglia prodotto", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_Certificazione", "IDRV_PO01_Certificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "ProtocolloCertificazione", "Protocollo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "CodiceCertificazione", "Codice", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "DescrizioneCertificazione", "Descrizione", dgchar, True, 1800, dgAlignleft
                .ColumnsHeader.Add "Predefinito", "Predefinito", dgBoolean, False, 1800, dgAligncenter
                .ColumnsHeader.Add "IDRV_PO01_EnteCertificazione", "IDRV_PO01_EnteCertificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "EnteCertificazione", "Ente certificatore", dgchar, False, 2000, dgAlignleft
                
        Set .Recordset = rsGrigliaFamiglia
        .LoadUserSettings
        .Refresh
    End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA_CERT_FAMIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_CERT_FAMIGLIA"
End Sub
Private Sub GET_GRIGLIA_CERT_AZIENDA()
On Error GoTo ERR_GET_GRIGLIA_CERT_AZIENDA
Dim OLD_Cursor As Long
Dim cl As dgColumnHeader
Dim sSQL As String

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_PO01_ParametriFiliale.IDRV_PO01_ParametriFiliale, RV_PO01_ParametriFiliale.IDAzienda, RV_PO01_ParametriFiliale.IDFiliale, RV_PO01_ParametriFiliale.IDRV_PO01_Certificazione, "
    sSQL = sSQL & "RV_PO01_ParametriFiliale.ProtocolloCertificazione, RV_PO01_Certificazione.CodiceCertificazione, RV_PO01_Certificazione.DescrizioneCertificazione, RV_PO01_ParametriFiliale.ProtocolloCertificazionePredefinito, "
    sSQL = sSQL & "RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione , RV_PO01_EnteCertificazione.EnteCertificazione "
    sSQL = sSQL & "FROM RV_PO01_ParametriFiliale INNER JOIN "
    sSQL = sSQL & "RV_PO01_Certificazione ON RV_PO01_ParametriFiliale.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_EnteCertificazione ON RV_PO01_Certificazione.IDRV_PO01_EnteCertificazione = RV_PO01_EnteCertificazione.IDRV_PO01_EnteCertificazione "
    sSQL = sSQL & " WHERE RV_PO01_ParametriFiliale.IDAzienda=" & TheApp.IDFirm
    If frmMain.CDCertificazione.KeyFieldID > 0 Then
        sSQL = sSQL & " AND RV_PO01_ParametriFiliale.IDRV_PO01_Certificazione=" & frmMain.CDCertificazione.KeyFieldID
    End If
    
    
    Set rsGrigliaAzienda = New ADODB.Recordset
    
    rsGrigliaAzienda.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaCertAzienda
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDRV_PO01_ParametriFiliale", "IDRV_PO01_ParametriFiliale", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_Certificazione", "IDRV_PO01_Certificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "ProtocolloCertificazione", "Protocollo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "CodiceCertificazione", "Codice", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "DescrizioneCertificazione", "Descrizione", dgchar, True, 1800, dgAlignleft
                .ColumnsHeader.Add "ProtocolloCertificazionePredefinito", "Predefinito", dgBoolean, False, 1800, dgAligncenter
                .ColumnsHeader.Add "IDRV_PO01_EnteCertificazione", "IDRV_PO01_EnteCertificazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "EnteCertificazione", "Ente certificatore", dgchar, False, 2000, dgAlignleft
        Set .Recordset = rsGrigliaAzienda
        .LoadUserSettings
        .Refresh
    End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA_CERT_AZIENDA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_CERT_AZIENDA"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        If Me.SSTab1.Tab = 0 Then
            GrigliaCertSocio_DblClick
        End If
        
        If Me.SSTab1.Tab = 1 Then
            GrigliaCertFamiglia_DblClick
        End If
    End If
End Sub


Private Sub Form_Load()
    GET_GRIGLIA_CERT_SOCIO
    GET_GRIGLIA_CERT_FAMIGLIA
    GET_GRIGLIA_CERT_AZIENDA
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rsGrigliaSocio.Close
    Set rsGrigliaSocio = Nothing
    
    rsGrigliaFamiglia.Close
    Set rsGrigliaFamiglia = Nothing
End Sub

Private Sub GrigliaCertAzienda_DblClick()
On Error GoTo ERR_GrigliaCertAzienda_DblClick
    If Not ((rsGrigliaAzienda.EOF) And (rsGrigliaAzienda.BOF)) Then
        frmMain.txtProtocolloCertificazione.Text = Me.GrigliaCertAzienda.AllColumns("ProtocolloCertificazione").Value
        Unload Me
    End If
Exit Sub
ERR_GrigliaCertAzienda_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCertAzienda_DblClick"
End Sub

Private Sub GrigliaCertFamiglia_DblClick()
On Error GoTo ERR_GrigliaCertFamiglia_DblClick
    If Not ((rsGrigliaFamiglia.EOF) And (rsGrigliaFamiglia.BOF)) Then
        frmMain.txtProtocolloCertificazione.Text = Me.GrigliaCertFamiglia.AllColumns("ProtocolloCertificazione").Value
        Unload Me
    End If
Exit Sub
ERR_GrigliaCertFamiglia_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCertFamiglia_DblClick"
End Sub
Private Sub GrigliaCertSocio_DblClick()
On Error GoTo ERR_GrigliaCertSocio_DblClick
    If Not ((rsGrigliaSocio.EOF) And (rsGrigliaSocio.BOF)) Then
        frmMain.txtProtocolloCertificazione.Text = Me.GrigliaCertSocio.AllColumns("ProtocolloCertificazione").Value
        Unload Me
    End If
Exit Sub
ERR_GrigliaCertSocio_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCertSocio_DblClick"
End Sub

