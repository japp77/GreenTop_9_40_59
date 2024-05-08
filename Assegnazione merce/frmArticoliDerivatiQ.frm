VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmArticoliDerivatiQ 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmArticoliDerivatiQ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   13125
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7223
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
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmArticoliDerivatiQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub Form_Activate()
    If rsArt.EOF = False Then
        Me.GrigliaCorpo.Enabled = True
        Me.GrigliaCorpo.SetFocus
    Else
        Me.GrigliaCorpo.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        GrigliaCorpo_Click
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Me.Top = frmMain.PicForm.Top
    Me.Left = frmMain.PicForm.Left
    
    SettaggioGriglia
     
End Sub
Private Sub SettaggioGriglia()
'On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_POArticoloFiglio.IDRV_POArticoloFiglio, RV_POArticoloFiglio.IDArticoloFiglio, RV_POArticoloFiglio.IDArticolo, RV_POArticoloFiglio.PesoPerOrdinamento, "
    sSQL = sSQL & "RV_POArticoloFiglio.IDRV_POTipoLavorazione, RV_POTipoLavorazione.TipoLavorazione, Articolo.CodiceArticolo, Articolo.Articolo,"
    sSQL = sSQL & "RV_POTipoCategoria.IDRV_POTipoCategoria , RV_POTipoCategoria.TipoCategoria, RV_POCalibro.IDRV_POCalibro, RV_POCalibro.Calibro,  Articolo.DescrizioneArticoloRidotta "
    sSQL = sSQL & "FROM RV_POTipoCategoria RIGHT OUTER JOIN "
    sSQL = sSQL & "Articolo ON RV_POTipoCategoria.IDRV_POTipoCategoria = Articolo.RV_POIDTipoCategoria LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POCalibro ON Articolo.RV_POIDCalibro = RV_POCalibro.IDRV_POCalibro RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_POArticoloFiglio ON Articolo.IDArticolo = RV_POArticoloFiglio.IDArticoloFiglio LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoLavorazione ON RV_POArticoloFiglio.IDRV_POTipoLavorazione = RV_POTipoLavorazione.IDRV_POTipoLavorazione "
    sSQL = sSQL & "WHERE RV_POArticoloFiglio.IDArticolo=" & Link_ArticoloPadre
    sSQL = sSQL & " AND ((Articolo.IDTipoProdotto=" & Link_TipoAumentoPeso & ")"
    sSQL = sSQL & " OR (Articolo.IDTipoProdotto=" & Link_TipoScarto & ")"
    sSQL = sSQL & " OR (Articolo.IDTipoProdotto=" & Link_TipoCaloPeso & "))"
    sSQL = sSQL & " ORDER BY RV_POArticoloFiglio.PesoPerOrdinamento, Articolo.CodiceArticolo "
    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.Data
        
        With Me.GrigliaCorpo
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "PesoPerOrdinamento", "Peso", dgNumeric, True, 800, dgAlignRight
                .ColumnsHeader.Add "IDArticolo", "ID", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "IDArticoloFiglio", "IDFiglio", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "CodiceArticolo", "Codice Articolo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "DescrizioneArticoloRidotta", "Descrizione", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "TipoCategoria", "Categoria", dgchar, True, 1500, dgAlignleft
                .ColumnsHeader.Add "Calibro", "Calibro", dgchar, True, 1500, dgAlignleft
                .ColumnsHeader.Add "TipoLavorazione", "Tipo lav.", dgchar, True, 1500, dgAlignleft
            Set .Recordset = rsArt.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsArt Is Nothing) Then
        rsArt.CloseResultset
        Set rsArt = Nothing
    End If
End Sub

Private Sub GrigliaCorpo_Click()

    frmMain.CDArticolo_Q.Load fnNotNullN(Me.GrigliaCorpo.AllColumns("IDArticoloFiglio").Value)

    Unload Me
End Sub

