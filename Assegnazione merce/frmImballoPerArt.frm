VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmImballoPerArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONA IMBALLO"
   ClientHeight    =   3015
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImballoPerArt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5318
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
Attribute VB_Name = "frmImballoPerArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub GET_GRIGLIA_PROCESSI()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3



sSQL = "SELECT IDArticoloImballo, "
sSQL = sSQL & "Articolo.CodiceArticolo, "
sSQL = sSQL & "Articolo.Articolo "
sSQL = sSQL & "FROM dbo.RV_PODistintaBaseRigheConf INNER JOIN "
sSQL = sSQL & "dbo.RV_PODistintaBase ON dbo.RV_PODistintaBaseRigheConf.IDRV_PODistintaBase = dbo.RV_PODistintaBase.IDRV_PODistintaBase INNER JOIN "
sSQL = sSQL & "dbo.Articolo ON dbo.RV_PODistintaBaseRigheConf.IDArticoloImballo = dbo.Articolo.IDArticolo "
sSQL = sSQL & "WHERE dbo.RV_PODistintaBase.IDArticolo=" & frmMain.CDArticolo.KeyFieldID
sSQL = sSQL & " GROUP BY IDArticoloImballo, Articolo.CodiceArticolo, Articolo.Articolo "

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection


With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    
   
    .ColumnsHeader.Add "IDArticoloImballo", "IDArticoloImballo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceArticolo", "Codice", dgchar, True, 2500, dgAlignleft
    .ColumnsHeader.Add "Articolo", "Imballo", dgchar, True, 3500, dgAlignleft
        
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
    
End With

Cn.CursorLocation = OLDCursor

Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_PROCESSI"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrigliaCorpo_DblClick
    End If
    
End Sub

Private Sub Form_Load()
        GET_GRIGLIA_PROCESSI
End Sub

Private Sub GrigliaCorpo_DblClick()
    If ((Me.GrigliaCorpo.Recordset.EOF) And (Me.GrigliaCorpo.Recordset.BOF)) Then Exit Sub
    
    frmMain.CDCodiceImballo.Load Me.GrigliaCorpo.AllColumns("IDArticoloImballo").Value

    Unload Me
    
    
End Sub

