VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaLottoDiCampagna 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEZIONE LOTTO DI PRODUZIONE"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLottoChiuso 
      Caption         =   "Visualizza anche i lotti di produzione chiusi"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9135
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7011
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
Private rsArt As DmtOleDbLib.adoResultset
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
'On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_PO01_LottoCampagna "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    If frmMain.chkSelLottoProdAnaFatt.Value = vbUnchecked Then
        sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocio.KeyFieldID
    Else
        sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocioFatt.KeyFieldID
    End If
    If Me.chkLottoChiuso.Value = vbUnchecked Then
        sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(Me.chkLottoChiuso.Value)
    End If
    
    Set rsArt = Cn.OpenResultset(sSQL)
        Set rsEvent = rsArt.Data
    
    With Me.Griglia
    Set .PaintNotifyObj = gPaintNotify
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_PO01_LottoCampagna", "ID", dgNumeric, False, 1000, dgAlignleft
            .ColumnsHeader.Add "CodiceLotto", "Codice lotto di campagna", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DescrizioneLotto", "Descrizione lotto di campagna", dgchar, True, 2000, dgAlignleft
            
            .ColumnsHeader.Add "Stato", "Stato", dgchar, True, 1500, dgAlignleft, , , True
            .ColumnsHeader.Add "DataSbloccoLotto", "Data di sblocco", dgDate, 1500, 0
            .ColumnsHeader.Add "Famiglia", "Famiglia prodotti", dgchar, True, 2500, dgAlignleft, , , True
            .ColumnsHeader.Add "Varieta", "Varietà", dgchar, True, 2000, dgAlignleft, , , True
            .ColumnsHeader.Add "Serre", "Appezzamenti", dgchar, True, 5500, dgAlignleft, , , True
            
        Set .Recordset = rsArt.Data
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = vbKeyReturn Then
    Griglia_DblClick
End If

End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub
Private Sub Griglia_DblClick()
    frmMain.txtIDLottoCampagna.Value = fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_LottoCampagna").Value)
    frmMain.txtLottoDiConferimento.Text = fnNotNull(Me.Griglia.AllColumns("CodiceLotto").Value)
    frmMain.txtDataSbloccoLotto.Value = fnNotNullN(Me.Griglia.AllColumns("DataSbloccoLotto").Value)
    Unload Me
End Sub
