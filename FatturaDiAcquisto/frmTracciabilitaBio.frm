VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmTracciabilitaBio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TRACCIABILITA' LOTTO DI CAMPAGNA"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
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
Attribute VB_Name = "frmTracciabilitaBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset
Public gPaintNotify As PaintNotify


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
    
    sSQL = "SELECT RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna, RV_PO01_LottoCampagna.IDRV_PO01_Schema, RV_PO01_SerraPerLotto.IDRV_PO01_Serra, "
    sSQL = sSQL & "RV_PO01_Serra.Codice , RV_PO01_SerraPerLotto.DimensioneMQ, RV_PO01_SerraPerLotto.DimensioneHA "
    sSQL = sSQL & "FROM RV_PO01_SerraPerLotto LEFT OUTER JOIN  "
    sSQL = sSQL & "RV_PO01_Serra ON RV_PO01_SerraPerLotto.IDRV_PO01_Serra = RV_PO01_Serra.IDRV_PO01_Serra RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_LottoCampagna ON RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna = RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna "
    sSQL = sSQL & "WHERE RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna=" & fnNormNumber(frmMain.txtIDLottoCampagna.Text)
    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.Data
        
        With Me.Griglia
        Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_PO01_LottoCampagna", "ID", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "Codice", "Serra", dgchar, True, 1500, dgAlignleft
                .ColumnsHeader.Add "DimensioneMQ", "Dimensione in Mq", dgDouble, True, 1500, dgAlignleft
                .ColumnsHeader.Add "DimensioneHA", "Dimensione in Ha", dgchar, True, 1500, dgAlignleft
                .ColumnsHeader.Add "Settore", "Settore (Terreno)", dgchar, True, 8000, dgAlignleft, , , True
                
            Set .Recordset = rsArt.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub
Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub
Private Sub Griglia_DblClick()
    'frmMain.txtIDLottoCampagna.Value = fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_LottoCampagna").Value)
    'frmMain.txtLottoDiConferimento.Text = fnNotNull(Trim(Me.Griglia.AllColumns("Codice").Value))
    'Unload Me
End Sub

