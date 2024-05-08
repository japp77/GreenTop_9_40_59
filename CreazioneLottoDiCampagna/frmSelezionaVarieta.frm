VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaVarieta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEZIONE VARIETA' PRODOTTI"
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
   Begin VB.CheckBox chkAllVarieta 
      Caption         =   "Visualizza tutte le varietà"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12615
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
Attribute VB_Name = "frmSelezionaVarieta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset
Public gPaintNotify As PaintNotify

Private Sub chkAllVarieta_Click()
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
    
    sSQL = "SELECT RV_PO01_Varieta.IDRV_PO01_Varieta, RV_PO01_Varieta.Varieta, RV_PO01_Varieta.IDRV_PO01_FamigliaProdotti, "
    sSQL = sSQL & "RV_PO01_Varieta.CodiceImEx, RV_PO01_FamigliaProdotti.FamigliaProdotti "
    sSQL = sSQL & "FROM RV_PO01_Varieta LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_FamigliaProdotti ON RV_PO01_Varieta.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti "
    
    If Me.chkAllVarieta.Value = vbUnchecked Then
        sSQL = sSQL & "WHERE (RV_PO01_Varieta.IDRV_PO01_FamigliaProdotti=" & frmMain.cboFamigliaProdotti.CurrentID & ") "
    End If
    
    sSQL = sSQL & "ORDER BY RV_PO01_Varieta.Varieta"
    
    
    
    
    
    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.Data
        
        With Me.Griglia
        Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_PO01_Varieta", "ID", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "Varieta", "Varieta", dgchar, True, 4000, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_FamigliaProdotti", "ID Famiglia Prodotti", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "FamigliaProdotti", "Famiglia Prodotti", dgchar, True, 4000, dgAlignleft
                .ColumnsHeader.Add "CodiceImEx", "Codice Import/Export", dgchar, True, 4000, dgAlignleft, , , True
            Set .Recordset = rsArt.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Varieta"
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
    frmMain.txtIDVarieta.Value = fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_Varieta").Value)
    frmMain.txtVarieta.Text = fnNotNull(Me.Griglia.AllColumns("Varieta").Value)
    Unload Me
End Sub
