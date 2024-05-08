VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmElencoTrattAgg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elenco trattenute aggiuntive"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmElencoTrattAgg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4683
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
Attribute VB_Name = "frmElencoTrattAgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    CONFERMA_PAR_TRATT_AGG = False
    CREA_RECORDSET
End Sub
Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "ID", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Descrizione", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Percentuale", adDouble, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
        
sSQL = "SELECT * FROM RV_POParametriTrattenuteAgg "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND CalcolaSuTrattenute=1"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If (CONTROLLO_ESISTENZA_RIGA(LINK_LIQUIDAZIONE, fnNotNullN(rs!IDRV_POParametriTrattenuteAgg)) = False) Then
        rsGriglia.AddNew
            rsGriglia!ID = fnNotNullN(rs!IDRV_POParametriTrattenuteAgg)
            rsGriglia!Descrizione = rs!DescrizioneTrattenutaAggiuntiva
            rsGriglia!Percentuale = fnNotNullN(rs!Percentuale)
        rsGriglia.Update
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
        
GET_GRIGLIA
Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Function CONTROLLO_ESISTENZA_RIGA(idliquidazione As Long, idtrattagg As Long)
On Error GoTo ERR_CONTROLLO_ESISTENZA_RIGA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLiquidazioneRighe "
sSQL = sSQL & " WHERE IDRV_POLiquidazione=" & idliquidazione
sSQL = sSQL & " AND IDRV_POParametriTrattenuteAgg=" & idtrattagg

Set rs = Cn.OpenResultset(sSQL)

CONTROLLO_ESISTENZA_RIGA = Not rs.EOF

rs.CloseResultset
Set rs = Nothing
Exit Function

ERR_CONTROLLO_ESISTENZA_RIGA:
    MsgBox Err.Description, vbCritical, "CONTROLLO_ESISTENZA_RIGA"
End Function

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long
Dim sSQL As String

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    .LoadUserSettings
    
    
    .ColumnsHeader.Add "ID", "ID", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "Descrizione", "Trattenuta", dgchar, True, 7000, dgAlignleft
    
    Set cl = .ColumnsHeader.Add("Percentuale", "%", dgDouble, True, 2000, dgAlignRight)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."

                
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub GrigliaCorpo_DblClick()
    LINK_PAR_TRATT_AGG_RETURN = fnNotNullN(Me.GrigliaCorpo.AllColumns("ID").Value)
    CONFERMA_PAR_TRATT_AGG = True
    Unload Me
End Sub
