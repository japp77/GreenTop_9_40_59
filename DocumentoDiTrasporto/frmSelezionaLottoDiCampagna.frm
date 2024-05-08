VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaLottoDiCampagna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELEZIONE LOTTO DI PRODUZIONE"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   16890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelezionaLottoDiCampagna.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   16890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRicerca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   12135
   End
   Begin VB.CheckBox chkLottoChiuso 
      Caption         =   "Visualizza anche i lotti di produzione chiusi"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   12938
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
Attribute VB_Name = "frmSelezionaLottoDiCampagna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As ADODB.Recordset

Public gPaintNotify As PaintNotify
Private Sub chkLottoChiuso_Click()
    SettaggioGriglia
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
    Set gPaintNotify = New PaintNotify
    SettaggioGriglia
    Me.Griglia.Recordset.Requery
Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
    Unload Me
End Sub


Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    If (LINK_SOCIO_PER_LOTTO = 0) Then
        If (GET_CONTROLLO_SOCI_PER_CLIENTE(frmMain.cdAnagrafica.KeyFieldID) = 0) Then
            sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagnaPerSelezione "
            sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
            sSQL = sSQL & " AND (Provvisorio=0 OR Provvisorio IS NULL)"
            sSQL = sSQL & " AND (VirtualDelete=0 OR VirtualDelete IS NULL)"
            If Me.chkLottoChiuso.Value = vbUnchecked Then
                sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(Me.chkLottoChiuso.Value)
            End If
            If Len(Trim(Me.txtRicerca.Text)) > 0 Then
                sSQL = sSQL & " AND ("
                sSQL = sSQL & " (Anagrafica LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (Nome LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (CodiceSocio LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (CodiceLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (DescrizioneLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " )"
            End If
            sSQL = sSQL & " ORDER BY Anagrafica, Nome"
        Else
            sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagnaPerCliente "
            sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
            sSQL = sSQL & " AND IDAnagraficaCliente=" & frmMain.cdAnagrafica.KeyFieldID
            sSQL = sSQL & " AND (Provvisorio=0 OR Provvisorio IS NULL)"
            sSQL = sSQL & " AND (VirtualDelete=0 OR VirtualDelete IS NULL)"
            If Me.chkLottoChiuso.Value = vbUnchecked Then
                sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(Me.chkLottoChiuso.Value)
            End If
            If Len(Trim(Me.txtRicerca.Text)) > 0 Then
                sSQL = sSQL & " AND ("
                sSQL = sSQL & " (Anagrafica LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (Nome LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (CodiceSocio LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (CodiceLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " OR (DescrizioneLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
                sSQL = sSQL & " )"
            End If
            sSQL = sSQL & " ORDER BY Anagrafica, Nome"
        End If
    Else
        sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagna "
        sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND IDSocio=" & LINK_SOCIO_PER_LOTTO
        sSQL = sSQL & " AND (Provvisorio=0 OR Provvisorio IS NULL)"
        sSQL = sSQL & " AND (VirtualDelete=0 OR VirtualDelete IS NULL)"
        If Me.chkLottoChiuso.Value = vbUnchecked Then
            sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(Me.chkLottoChiuso.Value)
        End If
        If Len(Trim(Me.txtRicerca.Text)) > 0 Then
            sSQL = sSQL & " AND ("
            sSQL = sSQL & " (CodiceLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
            sSQL = sSQL & " OR (DescrizioneLotto LIKE " & fnNormString("%" & txtRicerca.Text & "%") & ")"
            sSQL = sSQL & " )"
        End If
        sSQL = sSQL & " ORDER BY CodiceLotto"
    End If
    Set rsArt = New ADODB.Recordset
    rsArt.Open sSQL, Cn.InternalConnection
    
    With Me.Griglia
    .UpdatePosition = True
    Set .PaintNotifyObj = gPaintNotify
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna", dgNumeric, False, 1000, dgAlignleft
            .ColumnsHeader.Add "IDSocio", "IDSocio", dgNumeric, False, 1000, dgAlignleft
            .ColumnsHeader.Add "Anagrafica", "Socio/Fornitore", dgchar, LINK_SOCIO_PER_LOTTO = 0, 2000, dgAlignleft
            .ColumnsHeader.Add "Nome", "Nome", dgchar, LINK_SOCIO_PER_LOTTO = 0, 2000, dgAlignleft
            .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, LINK_SOCIO_PER_LOTTO = 0, 1500, dgAlignleft
            .ColumnsHeader.Add "CodiceLotto", "Codice lotto di campagna", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DescrizioneLotto", "Descrizione lotto di campagna", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "StatoLotto", "Stato", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "DataSbloccoLotto", "Data di sblocco", dgDate, 1500, dgAlignleft
            .ColumnsHeader.Add "FamigliaProdotti", "Famiglia prodotti", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Varieta", "Varietà", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Serre", "Appezzamenti", dgchar, True, 5500, dgAlignleft, , , True
        Set .Recordset = rsArt
        
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    Cn.CursorLocation = OLDCursor
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
'If KeyCode = vbKeyReturn Then
'    Griglia_DblClick
'End If

End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub
Private Sub Griglia_DblClick()
On Error GoTo ERR_Griglia_DblClick:
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If (LINK_SOCIO_PER_LOTTO = 0) Then
    LINK_LOTTO_PROD_LAV = fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_LottoCampagna").Value)
    frmMain.txtRaggrRigaOrdine.Text = frmMain.GET_DESCRIZIONE_LOTTO_PROD_LAV(LINK_LOTTO_PROD_LAV)
Else
    frmMain.txtLottoCampagna.Text = fnNotNull(Me.Griglia.AllColumns("CodiceLotto").Value)
    sSQL = "SELECT IDRV_PO01_LottoCampagna, Acquistato "
    sSQL = sSQL & " FROM RV_PO01_LottoCampagna "
    sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_LottoCampagna").Value)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        If (fnNotNullN(rs!Acquistato) = 1) Then
            frmMain.cboTipoDocumentoCoop.WriteOn 2
        Else
            frmMain.cboTipoDocumentoCoop.WriteOn 1
        End If
    End If

    rs.CloseResultset
    Set rs = Nothing
End If

Unload Me
    
Exit Sub
ERR_Griglia_DblClick:
    MsgBox Err.Description, vbCritical, "Griglia_DblClick"

End Sub
Private Function GET_CONTROLLO_SOCI_PER_CLIENTE(IDCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_SOCI_PER_CLIENTE = 0
sSQL = "SELECT COUNT(IDRV_POConfigurazioneClienteSoci) AS Numero "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteSoci "
sSQL = sSQL & "WHERE IDAnagraficaCliente=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_SOCI_PER_CLIENTE = fnNotNullN(rs!Numero)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Griglia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Griglia_DblClick
    End If
End Sub

Private Sub txtRicerca_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        SettaggioGriglia
    End If
End Sub

