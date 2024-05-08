VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmCollegamentiConferimento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "COLLEGAMENTI APERTI PER QUESTO CONFERIMENTO"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12810
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11880
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
Attribute VB_Name = "frmCollegamentiConferimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub SettaggioGriglia()
'On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_POTMPCollegamentiConferimento "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & LINK_TESTA_DOCUMENTO
    
    
    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.Data
        
        With Me.Griglia
        
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_POTMPCollegamentiConferimento", "IDRV_POTMPCollegamentiConferimento", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "Oggetto", "Oggetto", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "NumeroDocumento", "N° Doc.", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "DataDocumento", "Data Doc.", dgDate, True, 1000, dgAlignleft
                .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "MacchinaPC", "P.C.", dgchar, True, 2000, dgAlignleft
            
            Set .Recordset = rsArt.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub

Private Sub Form_Load()
    SettaggioGriglia
End Sub
