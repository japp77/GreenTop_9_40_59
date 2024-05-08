VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmProtICEPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIGURAZIONE PROTOCOLLO ICE PER PERIODO"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13545
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProtICEPeriodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElimina 
      Caption         =   "ELIMINA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "SALVA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "NUOVO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13455
      Begin DMTEDITNUMLib.dmtNumber txtID 
         Height          =   315
         Left            =   11040
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTDATETIMELib.dmtDate txtDaData 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   360
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.TextBox txtProtocolloICE 
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   360
         Width           =   5775
      End
      Begin DMTDATETIMELib.dmtDate txtAData 
         Height          =   315
         Left            =   7920
         TabIndex        =   9
         Top             =   360
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtProgressivo 
         Height          =   315
         Left            =   9840
         TabIndex        =   12
         Top             =   360
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTDATETIMELib.dmtDate txtMaxData 
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Progressivo"
         Height          =   255
         Index           =   2
         Left            =   9840
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fine periodo"
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Inizio periodo"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblProtocolloICE 
         Caption         =   "Protocollo ICE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5775
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7646
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
Attribute VB_Name = "frmProtICEPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private IDConfigurazione As Long

Private Sub Nuovo_Record()
On Error GoTo ERR_Nuovo_Record
    IDConfigurazione = 0
    Me.txtID.Value = 0
    Me.txtDaData.Value = GetMaxDataFinePeriodo
    Me.txtAData.Value = 0
    Me.txtProgressivo.Value = 1

Exit Sub
ERR_Nuovo_Record:
    MsgBox Err.Description, vbCritical, "Nuovo_Record"

End Sub
Private Sub Salva_Record()
On Error GoTo ERR_Salva_Record
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim NumeroRecordSel As Long

    If Permesso_Salvataggio = False Then Exit Sub
    
    sSQL = "SELECT * FROM RV_POProgProtocolloICEPeriodo "
    sSQL = sSQL & "WHERE IDRV_POProgProtocolloICEPeriodo=" & IDConfigurazione
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

    If rs.EOF Then
        rs.AddNew
        rs!IDRV_POProgProtocolloICEPeriodo = fnGetNewKey("RV_POProgProtocolloICEPeriodo", "IDRV_POProgProtocolloICEPeriodo")
        rs!IDAzienda = TheApp.IDFirm
        rs!IDFiliale = frmMain.cboFiliale.CurrentID
        NumeroRecordSel = Me.Griglia.ListCount
    Else
        NumeroRecordSel = Me.Griglia.ListIndex - 1
    End If
    
    rs!DaData = Me.txtDaData.Text
    rs!AData = Me.txtAData.Text
    rs!IDRV_POProtocolloICE = Me.txtID.Value
    rs!Progressivo = Me.txtProgressivo.Value
    
    rs.Update

    rs.Close
    Set rs = Nothing
    
    GET_GRIGLIA
    
    Me.Griglia.Recordset.Move NumeroRecordSel

Exit Sub
ERR_Salva_Record:
    MsgBox Err.Description, vbCritical, "Salva_Record"
End Sub
Private Sub Elimina_Record()
On Error GoTo ERR_Elimina_Record
Dim sSQL As String
Dim Testo As String

    If IDConfigurazione = 0 Then Exit Sub
    
    'CONTROLLO UTILIZZO PROTOCOLLO NEI DOCUMENTI DI VENDITA
        
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Prima di eliminare la configurazione selezionata assicurarsi che non sia stata utilizzata in nessun documento di vendita" & vbCrLf
    Testo = Testo & "Vuoi procedere?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub
    
    sSQL = "DELETE FROM RV_POProgProtocolloICEPeriodo    "
    sSQL = sSQL & " WHERE IDRV_POProgProtocolloICEPeriodo=" & IDConfigurazione
    
    Cn.Execute sSQL
    
    GET_GRIGLIA
Exit Sub
ERR_Elimina_Record:
    MsgBox Err.Description, vbCritical, "Elimina_Record"
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader

sSQL = "SELECT RV_POProgProtocolloICEPeriodo.IDRV_POProgProtocolloICEPeriodo, RV_POProgProtocolloICEPeriodo.IDFiliale, RV_POProgProtocolloICEPeriodo.IDAzienda, "
sSQL = sSQL & "RV_POProgProtocolloICEPeriodo.IDRV_POProtocolloICE, RV_POProgProtocolloICEPeriodo.DaData, RV_POProgProtocolloICEPeriodo.AData, RV_POProgProtocolloICEPeriodo.Progressivo, "
sSQL = sSQL & "RV_POProtocolloICE.ProtocolloICE "
sSQL = sSQL & "FROM RV_POProgProtocolloICEPeriodo INNER JOIN "
sSQL = sSQL & "RV_POProtocolloICE ON RV_POProgProtocolloICEPeriodo.IDRV_POProtocolloICE = RV_POProtocolloICE.IDRV_POProtocolloICE "
sSQL = sSQL & "WHERE RV_POProgProtocolloICEPeriodo.IDFiliale=" & frmMain.cboFiliale.CurrentID

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient
rsGriglia.Open sSQL, Cn.InternalConnection
    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POProgProtocolloICEPeriodo", "IDRV_POProgProtocolloICEPeriodo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDFiliale", "IDFiliale", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POProtocolloICE", "IDRV_POProtocolloICE", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "ProtocolloICE", "Descrizione protocollo ICE", dgchar, True, 3500, dgAlignleft

            .ColumnsHeader.Add "DaData", "Inizio periodo", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "AData", "Fine periodo", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Progressivo", "Progressivo", dgNumeric, True, 1500, dgAlignRight
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Function GetMaxDataFinePeriodo() As Long
On Error GoTo ERR_GetMaxDataFinePeriodo
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(AData) as MaxData "
sSQL = sSQL & "FROM RV_POProgProtocolloICEPeriodo "
sSQL = sSQL & "WHERE IDFiliale=" & frmMain.cboFiliale.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    txtMaxData.Value = Date
Else
    If fnNotNullN(rs!MaxData) = 0 Then
        txtMaxData.Value = Date
    Else
        txtMaxData.Value = DateAdd("d", 1, rs!MaxData)
    End If
End If

rs.CloseResultset
Set rs = Nothing

GetMaxDataFinePeriodo = Me.txtMaxData.Value

Exit Function
ERR_GetMaxDataFinePeriodo:
    MsgBox Err.Description, vbCritical, "GetMaxDataFinePeriodo"
End Function

Private Sub cmdElimina_Click()
    Elimina_Record
End Sub

Private Sub cmdNuovo_Click()
    Nuovo_Record
End Sub

Private Sub cmdSalva_Click()
    Salva_Record
End Sub

Private Sub Form_Load()
    GET_GRIGLIA

    Nuovo_Record
End Sub

Private Sub Griglia_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
On Error GoTo ERR_Griglia_Reposition
    IDConfigurazione = fnNotNullN(Me.Griglia.AllColumns("IDRV_POProgProtocolloICEPeriodo").Value)
    Me.txtID.Value = fnNotNullN(Me.Griglia.AllColumns("IDRV_POProtocolloICE").Value)
    Me.txtDaData.Value = fnNotNullN(Me.Griglia.AllColumns("DaData").Value)
    Me.txtAData.Value = fnNotNullN(Me.Griglia.AllColumns("AData").Value)
    Me.txtProgressivo.Value = fnNotNullN(Me.Griglia.AllColumns("Progressivo").Value)

Exit Sub
ERR_Griglia_Reposition:
    MsgBox Err.Description, vbCritical, "Griglia_Reposition"
End Sub

Private Sub lblProtocolloICE_Click()
On Error GoTo ERR_lblProtocolloICE_Click
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset


Set oSearch = New dmtFind.Find
oSearch.Database = Cn
oSearch.Caption = "Protocollo ICE"
oSearch.AddDisplayField "Protocollo", "ProtocolloICE", 1

oSearch.Filters.Add "ProtocolloICE", Me.txtProtocolloICE

oSearch.Start = ""

sSQL = "SELECT * FROM RV_POProtocolloICE"

oSearch.SQL = fnAnsi2Jet(sSQL)
                        
Set oRes = oSearch.Exec

If Not oRes.EOF Then
    Me.txtID.Value = fnNotNullN(oRes!IDRV_POProtocolloICE)

End If

Set oRes = Nothing
Set oSearch = Nothing
Exit Sub
ERR_lblProtocolloICE_Click:
    MsgBox Err.Description, vbCritical, "lblProtocolloICE_Click"
End Sub

Private Sub txtID_Change()
On Error GoTo ERR_txtID_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POProtocolloICE "
sSQL = sSQL & "WHERE IDRV_POProtocolloICE=" & Me.txtID.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtProtocolloICE.Text = ""
Else
    Me.txtProtocolloICE.Text = fnNotNull(rs!ProtocolloICE)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_txtID_Change:
    MsgBox Err.Description, vbCritical, "txtID_Change"
End Sub
Private Function Permesso_Salvataggio() As Boolean
On Error GoTo ERR_Permesso_Salvataggio
Dim Testo As String

Permesso_Salvataggio = False

If Me.txtID.Value = 0 Then
    Testo = "Inserire il protocollo ICE"
    MsgBox Testo, vbCritical, "Validazione dati"
    Exit Function
End If

If Me.txtDaData.Value = 0 Then
    Testo = "Inserire la data di inizio periodo"
    MsgBox Testo, vbCritical, "Validazione dati"
    Exit Function
End If

If Me.txtAData.Value = 0 Then
    Testo = "Inserire la data di fine periodo"
    MsgBox Testo, vbCritical, "Validazione dati"
    Exit Function
End If

If Me.txtDaData.Value > Me.txtAData.Value Then
    Testo = "la data di fine periodo deve maggiore o uguale alla data di inizio periodo"
    MsgBox Testo, vbCritical, "Validazione dati"
    Exit Function

End If

If Me.txtProgressivo.Value = 0 Then
    Testo = "Inserire il progressivo iniziale"
    MsgBox Testo, vbCritical, "Validazione dati"
    Exit Function
End If

Permesso_Salvataggio = True
Exit Function
ERR_Permesso_Salvataggio:
    MsgBox Err.Description, vbCritical, "Permesso_Salvataggio"
End Function
