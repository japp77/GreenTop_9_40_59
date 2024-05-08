VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmTrovaPedana 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TROVA PEDANA"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11205
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
   ScaleHeight     =   4635
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRicerca 
      Caption         =   "Ricerca"
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtCodicePedana 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin DmtGridCtl.DmtGrid GridPedana 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Codice pedana"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmTrovaPedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As DmtOleDbLib.adoResultset

Private Sub cmdRicerca_Click()
    
    fncGriglia

End Sub

Private Sub Form_Activate()
    
    If WHERE_TROVA_PEDANA = 1 Then
        Me.txtCodicePedana.Text = "%" & frmMain.txtCodicePedana.Text & "%"
    End If
    
    If WHERE_TROVA_PEDANA = 2 Then
        Me.txtCodicePedana.Text = "%" & frmMain.txtCodicePedana.Text & "%"
    End If
    
    cmdRicerca_Click
    
End Sub

Private Sub fncGriglia()
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
        
    
'sSQL_WHERE = ""
'sSQL = "SELECT IDRV_POPedana, Mese, Giorno, Codice, Anno "
'sSQL = sSQL & "FROM RV_POPedana "
'sSQL = sSQL & "WHERE Codice LIKE " & fnNormString(Me.txtCodicePedana.Text & "%")
'sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
'sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

sSQL = "SELECT RV_POPedana.IDRV_POPedana, RV_POPedana.Anno, RV_POPedana.Codice, RV_POPedana.Mese, RV_POPedana.Giorno, RV_POPedana.IDFiliale, "
sSQL = sSQL & "RV_POPedana.IDAzienda, RV_POAssegnazioneMerce.IDOggettoOrdine, RV_POAssegnazioneMerce.NumeroOrdine,"
sSQL = sSQL & "RV_POAssegnazioneMerce.DataOrdine, RV_POAssegnazioneMerce.IDCliente, Anagrafica.Anagrafica, Anagrafica.Nome,"
sSQL = sSQL & "RV_POPedana.IDRV_POTipoPedana , RV_POPedana.Descrizione, RV_POTipoPedana.TipoPedana "
sSQL = sSQL & "FROM RV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POTipoPedana.IDRV_POTipoPedana = RV_POPedana.IDRV_POTipoPedana LEFT OUTER JOIN "
sSQL = sSQL & "Anagrafica INNER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON Anagrafica.IDAnagrafica = RV_POAssegnazioneMerce.IDCliente ON "
sSQL = sSQL & "RV_POPedana.IDRV_POPedana = RV_POAssegnazioneMerce.IDRV_POPedana "
sSQL = sSQL & "GROUP BY RV_POPedana.IDRV_POPedana, RV_POPedana.Anno, RV_POPedana.Codice, RV_POPedana.Mese, RV_POPedana.Giorno, RV_POPedana.IDFiliale,"
sSQL = sSQL & "RV_POPedana.IDAzienda, RV_POAssegnazioneMerce.IDOggettoOrdine, RV_POAssegnazioneMerce.NumeroOrdine,"
sSQL = sSQL & "RV_POAssegnazioneMerce.DataOrdine, RV_POAssegnazioneMerce.IDCliente, Anagrafica.Anagrafica, Anagrafica.Nome,"
sSQL = sSQL & "RV_POPedana.IDRV_POTipoPedana , RV_POPedana.Descrizione, RV_POTipoPedana.TipoPedana "
sSQL = sSQL & "Having (RV_POPedana.Codice LIKE " & fnNormString(Me.txtCodicePedana.Text & "%") & ")"
sSQL = sSQL & " AND  (RV_POPedana.IDFiliale = " & TheApp.Branch & ") "
sSQL = sSQL & " AND  (RV_POPedana.IDAzienda = " & TheApp.IDFirm & ")"



If WHERE_TROVA_PEDANA = 2 Then
    sSQL = sSQL & " AND  (RV_POAssegnazioneMerce.IDOggettoOrdine = " & frmAssegnazioneMerce.txtIDOrdineCliente.Value & ")"
End If

sSQL = sSQL & " ORDER BY ANNO Desc, Mese DESC, Giorno DESC, Codice DESC"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        Set rsGriglia = CnDMT.OpenResultset(sSQL)
            Set rsevent = rsGriglia.Data
    
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
            .ColumnsHeader.Add "NumeroOrdine", "N° Ord.", dgNumeric, True, 1500, dgAlignleft
            
            Set .Recordset = rsGriglia.Data
            .Refresh
        End With
        
    CnDMT.CursorLocation = OLDCursor

End Sub

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
    rsGriglia.CloseResultset
    Set rsGriglia = Nothing
End Sub

Private Sub GridPedana_DblClick()
    If WHERE_TROVA_PEDANA = 1 Then
        frmMain.txtCodicePedana.Text = Me.GridPedana.AllColumns("Codice").Value

    End If
    If WHERE_TROVA_PEDANA = 2 Then
        frmAssegnazioneMerce.txtIDPedana.Value = Me.GridPedana("IDRV_POPedana")
        frmAssegnazioneMerce.txtCodicePedana.Text = Me.GridPedana.AllColumns("Codice").Value
    End If
    
    Unload Me
End Sub
