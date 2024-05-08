VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRicerca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ricerca lotto per l'articolo"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10950
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
   ScaleHeight     =   5910
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVisualizzaLottoScad 
      Caption         =   "Visualizza lotti scaduto"
      Height          =   255
      Left            =   40
      TabIndex        =   1
      Top             =   120
      Width           =   10815
   End
   Begin DmtGridCtl.DmtGrid GridRicerca 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10935
      _ExtentX        =   19288
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
      UpdatePosition  =   0   'False
      UseUserSettings =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmRicerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As DmtOleDbLib.adoResultset

Private Sub chkVisualizzaLottoScad_Click()
   GetGrigliaRicerca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GridRicerca_DblClick
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    GetGrigliaRicerca

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub GetGrigliaRicerca()
On Error GoTo ERR_GetGrigliaLavorazione
Dim sSQL As String
Dim cl As dgColumnHeader

sSQL = "SELECT *"
sSQL = sSQL & "FROM LottoArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & frmMain.CDArticoloSemi.KeyFieldID
If Me.chkVisualizzaLottoScad.Value = 0 Then
    sSQL = sSQL & " AND DataScadenza>=" & fnNormDate(Date)
End If
sSQL = sSQL & " ORDER BY DataScadenza"

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    Set rsGriglia = Cn.OpenResultset(sSQL)
            Set rsEvent = rsGriglia.Data
        
    
        With Me.GridRicerca
                .ColumnsHeader.Clear
                    .ColumnsHeader.Add "IDLottoArticolo", "IDRiga", dgNumeric, False, 500, dgAlignleft
                    .ColumnsHeader.Add "DataScadenza", "Data scadenza", dgDate, True, 2000, dgAlignlef
                    .ColumnsHeader.Add "Codice", "Codice lotto", dgchar, True, 1500, dgAlignlef
                    .ColumnsHeader.Add "LottoArticolo", "Descrizione", dgchar, True, 2000, dgAlignleft
                    '.ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignlef
                    '.ColumnsHeader.Add "Articolo", "Descrizione articolo", dgchar, True, 2000, dgAlignleft
            Set .Recordset = rsGriglia.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GetGrigliaLavorazione:
    MsgBox Err.Description, vbCritical, "Griglia ricerca"
End Sub



Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Me.GridRicerca.Width = Me.Width - 150
        Me.GridRicerca.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsGriglia.CloseResultset
    Set rsGriglia = Nothing
End Sub



Private Sub GridRicerca_DblClick()
On Error GoTo ERR_GridRicerca
    frmMain.txtIDLottoArticolo.Value = Me.GridRicerca.AllColumns("IDLottoArticolo").Value
'    frmMain.txtCodiceLottoSeme.Text = Me.GridRicerca.AllColumns("Codice").Value
'    frmMain.txtDescrizioneLottoSeme.Text = Me.GridRicerca.AllColumns("LottoArticolo").Value
    Unload Me
Exit Sub
ERR_GridRicerca:
    MsgBox Err.Description, vbCritical, "GridRicerca_DblClick"
End Sub
