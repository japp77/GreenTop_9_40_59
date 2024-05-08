VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmTrovaPedana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TROVA PEDANA"
   ClientHeight    =   4800
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
   Icon            =   "frmTrovaPedana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GridPedana 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
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
   Begin VB.ComboBox cboAnno 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lbInfo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   300
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "PEDANE DEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmTrovaPedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private rsGriglia As DmtOleDbLib.adoResultset
Private rsGriglia As ADODB.Recordset


Private Sub cboAnno_Click()
    If IsNumeric(Me.cboAnno.Text) Then
        Me.lbInfo.Caption = "CARICAMENTO IN CORSO..................."
        DoEvents
            
        fncGriglia
        Me.lbInfo.Caption = ""
        DoEvents
    End If
End Sub

Private Sub Form_Activate()

If fncAnnoPedane = False Then
    MsgBox "Non ci sono pedane da gestire", vbInformation, "Trova pedana"
    Unload Me
Else
    Me.cboAnno.ListIndex = 0
End If

End Sub

Private Sub fncGriglia()
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
sSQL_WHERE = ""

sSQL = "SELECT RV_POPedana.IDRV_POPedana, RV_POPedana.Anno, RV_POPedana.Codice, RV_POPedana.Mese, RV_POPedana.Giorno, RV_POPedana.IDFiliale, "
sSQL = sSQL & "RV_POPedana.IDAzienda, RV_POAssegnazioneMerce.IDOggettoOrdine, RV_POAssegnazioneMerce.NumeroOrdine,"
sSQL = sSQL & "RV_POAssegnazioneMerce.DataOrdine, RV_POAssegnazioneMerce.IDCliente, Anagrafica.Anagrafica, Anagrafica.Nome,"
sSQL = sSQL & "RV_POPedana.IDRV_POTipoPedana , RV_POPedana.Descrizione, RV_POTipoPedana.TipoPedana, RV_POAssegnazioneMerce.NumeroListaPrelievo, RV_POPedana.PesoPedana "
sSQL = sSQL & "FROM RV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POTipoPedana.IDRV_POTipoPedana = RV_POPedana.IDRV_POTipoPedana LEFT OUTER JOIN "
sSQL = sSQL & "Anagrafica INNER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON Anagrafica.IDAnagrafica = RV_POAssegnazioneMerce.IDCliente ON "
sSQL = sSQL & "RV_POPedana.IDRV_POPedana = RV_POAssegnazioneMerce.IDRV_POPedana "
sSQL = sSQL & "GROUP BY RV_POPedana.IDRV_POPedana, RV_POPedana.Anno, RV_POPedana.Codice, RV_POPedana.Mese, RV_POPedana.Giorno, RV_POPedana.IDFiliale,"
sSQL = sSQL & "RV_POPedana.IDAzienda, RV_POAssegnazioneMerce.IDOggettoOrdine, RV_POAssegnazioneMerce.NumeroOrdine,"
sSQL = sSQL & "RV_POAssegnazioneMerce.DataOrdine, RV_POAssegnazioneMerce.IDCliente, Anagrafica.Anagrafica, Anagrafica.Nome,"
sSQL = sSQL & "RV_POPedana.IDRV_POTipoPedana , RV_POPedana.Descrizione, RV_POTipoPedana.TipoPedana, RV_POAssegnazioneMerce.NumeroListaPrelievo, RV_POPedana.PesoPedana "
sSQL = sSQL & "Having (RV_POPedana.Anno = " & Me.cboAnno.Text & ")"
sSQL = sSQL & " AND (RV_POPedana.IDFiliale = " & TheApp.Branch & ") "
sSQL = sSQL & " AND (RV_POPedana.IDAzienda = " & TheApp.IDFirm & ")"
sSQL = sSQL & " AND (RV_POAssegnazioneMerce.IDCliente=" & frmMain.ACSCliente.IDAnagrafica & ")"

If WHERE_TROVA_PEDANA = 1 Then
    If Len(Trim(frmMain.txtCodicePedana.Text)) > 0 Then
        sSQL = sSQL & " AND RV_POPedana.Codice LIKE " & fnNormString("%" & frmMain.txtCodicePedana.Text & "%")
    End If
End If
    
sSQL = sSQL & " ORDER BY RV_POPedana.Anno DESC, RV_POPedana.Mese DESC, RV_POPedana.Giorno DESC, RV_POPedana.IDRV_POPedana DESC"

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = adUseClient

    Set rsGriglia = New ADODB.Recordset 'Cn.OpenResultset(sSQL)
        rsGriglia.Open sSQL, Cn.InternalConnection
        'Set rsEvent = rsGriglia.Data

    With Me.GridPedana
        .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDRV_POPedana", "ID", dgNumeric, False, 800, dgAlignleft
        .ColumnsHeader.Add "Anno", "Anno", dgNumeric, True, 1000, dgAlignleft
        .ColumnsHeader.Add "Mese", "M", dgNumeric, True, 500, dgAlignleft
        .ColumnsHeader.Add "Giorno", "G", dgNumeric, True, 500, dgAlignleft
        .ColumnsHeader.Add "Codice", "Codice pedana", dgchar, True, 4000, dgAlignleft
        .ColumnsHeader.Add "IDRV_POTipoPedana", "IDRV_POTipoPedana", dgInteger, True, 500, dgAlignleft
        .ColumnsHeader.Add "TipoPedana", "TipoPedana", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "IDOggettoOrdine", "IDOggettoOrdine", dgInteger, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDCliente", "IDCliente", dgInteger, False, 1500, dgAlignleft
        .ColumnsHeader.Add "Anagrafica", "Cliente", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "Nome", "Nome", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "DataOrdine", "Data Ord.", dgDate, True, 2500, dgAlignleft
        .ColumnsHeader.Add "NumeroOrdine", "N° Ord.", dgNumeric, True, 1500, dgAlignRight
        .ColumnsHeader.Add "NumeroListaPrelievo", "N° lista", dgNumeric, True, 1500, dgAlignRight
        
        Set .Recordset = rsGriglia 'rsGriglia.Data
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor

End Sub
Private Function fncAnnoPedane() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = " SELECT Anno FROM RV_POPedana "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " GROUP BY Anno ORDER BY Anno DESC"

Set rs = Cn.OpenResultset(sSQL)

Me.cboAnno.Clear

If rs.EOF = False Then
    While Not rs.EOF
        Me.cboAnno.AddItem fnNotNullN(rs!Anno)
    rs.MoveNext
    Wend
    
    fncAnnoPedane = True
Else
    fncAnnoPedane = False
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GridPedana_DblClick
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rsGriglia Is Nothing) Then
    rsGriglia.Close
    Set rsGriglia = Nothing
End If
End Sub

Private Sub GridPedana_DblClick()
    
    LINK_PEDANA = Me.GridPedana.AllColumns("IDRV_POPedana").Value
    frmMain.txtCodicePedana.Text = fnNotNull(Me.GridPedana.AllColumns("Codice").Value)
    frmMain.CDTipoPedana.Load fnNotNullN(Me.GridPedana.AllColumns("IDRV_POTipoPedana").Value)
    frmMain.txtPesoPedana.Value = fnNotNullN(Me.GridPedana.AllColumns("PesoPedana").Value)

    
    Unload Me


End Sub
