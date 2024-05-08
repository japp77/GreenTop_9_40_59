VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmElencoNote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ELENCO NOTE DISPONIBILI"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmElencoNote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   16845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
Attribute VB_Name = "frmElencoNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset



Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = 3
    
sSQL = "SELECT * FROM RV_POIENotePerDocumentoTipoOggetto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoNotePerDocumento=" & LINK_TIPO_NOTA_SEL
sSQL = sSQL & " AND IDTipoOggetto=" & FrmFine.CboTipoDocumento.CurrentID
Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, CnDMT.InternalConnection


With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "ID", "IDRV_PONotePerDocumentoPerTipoOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_PONotePerDocumento", "IDRV_PONotePerDocumento", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POTipoNotePerDocumento", "IDRV_POTipoNotePerDocumento", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 2000, dgAlignleft
    .ColumnsHeader.Add "Annotazione", "Codice tipo pedana", dgchar, True, 5000, dgAlignleft

    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
    
CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Griglia_DblClick
    End If
End Sub

Private Sub Form_Load()
    CONFERMA_RIGA_NOTA = 0
    RETURN_RIGA_NOTA = ""
    
    GET_GRIGLIA
End Sub

Private Sub Griglia_DblClick()
    
    If ((Me.Griglia.Recordset.EOF) And (Me.Griglia.Recordset.BOF)) Then Exit Sub
    
    CONFERMA_RIGA_NOTA = 1
    
    RETURN_RIGA_NOTA = Me.Griglia.AllColumns("Annotazione").Value
    
    Unload Me
End Sub
