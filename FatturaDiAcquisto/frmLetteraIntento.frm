VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmLetteraIntento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lettere d'intento"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7335
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
   ScaleHeight     =   5265
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTutte 
      Caption         =   "Tutte"
      Height          =   255
      Left            =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3240
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8916
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
Attribute VB_Name = "frmLetteraIntento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Sub chkTutte_Click()
    GET_GRIGLIA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn Then Griglia_DblClick
    
End Sub

Private Sub Form_Load()
     GET_GRIGLIA
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader


sSQL = "SELECT LetteraIntento.*, SitoPerAnagrafica.SitoPerAnagrafica , TipoLetteraIntento.TipoLetteraIntento, TipoOperazioneEsenzione.TipoOperazioneEsenzione "
sSQL = sSQL & "FROM LetteraIntento LEFT OUTER JOIN "
sSQL = sSQL & "TipoOperazioneEsenzione ON LetteraIntento.IDTipoOperazioneEsenzione = TipoOperazioneEsenzione.IDTipoOperazioneEsenzione LEFT OUTER JOIN "
sSQL = sSQL & "SitoPerAnagrafica ON LetteraIntento.IDSitoPerAnagrafica = SitoPerAnagrafica.IDSitoPerAnagrafica LEFT OUTER JOIN "
sSQL = sSQL & "TipoLetteraIntento ON LetteraIntento.IDTipoLetteraIntento = TipoLetteraIntento.IDTipoLetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica_CF=" & frmMain.CDSocio.KeyFieldID
sSQL = sSQL & " AND IDTipoAnagrafica_CF=3"
'If Me.chkTutte.Value = vbUnchecked Then
'    sSQL = sSQL & " AND ((Anno=" & Year(frmMain.dtData.Text) & ") OR (AnnoOperazione=" & Year(frmMain.dtData.Text) & "))"
'End If

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
        .ColumnsHeader.Add "IDLetteraIntento", "IDLetteraIntento", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoLetteraIntento", "IDTipoLetteraIntento", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoLetteraIntento", "Tipo", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "Annullata", "Annullata", dgBoolean, True, 1500, dgAligncenter
        .ColumnsHeader.Add "Anno", "Anno", dgInteger, True, 1800, dgAlignRight
        .ColumnsHeader.Add "Data", "Data Reg.", dgDate, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Numero", "Numero Prog.", dgInteger, True, 1800, dgAlignRight
        
        .ColumnsHeader.Add "NumeroCliFor", "N° Reg. cliente", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "DataEmissione", "Data emissione", dgDate, True, 1800, dgAlignleft
        
        .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione", dgchar, False, 1200, dgAlignleft
        
        .ColumnsHeader.Add "IDIva", "IDIva", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneArticoloEsenzione", "Descrizione esenzione", dgchar, False, 1200, dgAlignleft
                
        .ColumnsHeader.Add "IDTipoOperazioneEsenzione", "IDTipoOperazioneEsenzione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoOperazioneEsenzione", "Tipo operazione", dgchar, False, 1800, dgAlignleft
        .ColumnsHeader.Add "AnnoOperazione", "Anno operazione", dgInteger, True, 1800, dgAlignRight
        .ColumnsHeader.Add "DaDataOperazione", "Da data Ope.", dgDate, True, 1800, dgAlignleft
        .ColumnsHeader.Add "ADataOperazione", "A data Ope.", dgDate, True, 1800, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("ImportoOperazione", "Importo", dgDouble, False, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    Set rsGriglia = Nothing
End If
End Sub
Private Sub Griglia_DblClick()
Dim Testo As String
    If Me.Griglia.AllColumns("Annullata") = True Then
        Testo = "La lettera d'intento risulta annullata"
        MsgBox Testo, vbCritical, "Selezione lettera d'intento"
        
        Exit Sub
    End If
    frmMain.txtIDLetteraIntento.Value = Me.Griglia.AllColumns("IDLetteraIntento").Value
    Unload Me
End Sub
