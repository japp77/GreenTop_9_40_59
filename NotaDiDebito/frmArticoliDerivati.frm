VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmArticoliDerivati 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7223
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
Attribute VB_Name = "frmArticoliDerivati"
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
     SettaggioGriglia
     
End Sub
Private Sub SettaggioGriglia()
'On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_POArticoloFiglio.IDRV_POArticoloFiglio, RV_POArticoloFiglio.IDArticoloFiglio, RV_POArticoloFiglio.IDArticolo, "
    sSQL = sSQL & "RV_POArticoloFiglio.IDRV_POTipoLavorazione , RV_POTipoLavorazione.TipoLavorazione, Articolo.CodiceArticolo, Articolo.Articolo "
    sSQL = sSQL & "FROM RV_POArticoloFiglio LEFT OUTER JOIN "
    sSQL = sSQL & "Articolo ON RV_POArticoloFiglio.IDArticoloFiglio = Articolo.IDArticolo LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoLavorazione ON RV_POArticoloFiglio.IDRV_POTipoLavorazione = RV_POTipoLavorazione.IDRV_POTipoLavorazione "
    sSQL = sSQL & "WHERE RV_POArticoloFiglio.IDArticolo=" & Link_ArticoloPadre
    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.Data
        
        With Me.GrigliaCorpo
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDArticolo", "ID", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "IDArticoloFiglio", "IDFiglio", dgNumeric, False, 1000, dgAlignleft
                .ColumnsHeader.Add "CodiceArticolo", "Codice Articolo", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 3500, dgAlignleft
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

    Unload Me
End Sub

Private Function GetTipoLavorazioneStandard(IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDTipoLavorazione FROM Articolo WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetTipoLavorazioneStandard = 0
Else
    GetTipoLavorazioneStandard = fnNotNullN(rs!RV_POIDTipoLavorazione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
