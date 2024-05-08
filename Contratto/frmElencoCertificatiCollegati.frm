VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmElencoCertificatiCollegati 
   Caption         =   "Elenco certificati collegati"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmElencoCertificatiCollegati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   18045
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   14420
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
Attribute VB_Name = "frmElencoCertificatiCollegati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    GET_GRIGLIA
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * FROM RV_POIECertificato "
sSQL = sSQL & " WHERE IDContratto=" & oDoc.IDOggetto
sSQL = sSQL & " ORDER BY DataCertificato DESC, NumeroCertificato DESC"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection

With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "IDRV_POCertificato", "IDRV_POCertificato", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAnagraficaCooperativa", "IDAnagraficaCooperativa", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceAnagraficaCooperativa", "Codice Coop", dgchar, False, 1500, dgAlignleft
    .ColumnsHeader.Add "AnagraficaCooperativa", "Cooperativa", dgchar, True, 3500, dgAlignleft
    .ColumnsHeader.Add "IDAnagraficaSocio", "IDAnagraficaSocio", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceAnagraficaSocio", "Codice socio", dgchar, False, 1500, dgAlignleft
    .ColumnsHeader.Add "AnagraficaSocio", "Socio/Produttore", dgchar, True, 3500, dgAlignleft
    .ColumnsHeader.Add "DataCertificato", "Data certificato", dgDate, True, 2000, dgAlignleft
    .ColumnsHeader.Add "NumeroCertificato", "Numero Certificato", dgchar, True, 2000, dgAlignRight
    .ColumnsHeader.Add "NumeroDocumentoSocio", "Numero doc. socio", dgchar, False, 2000, dgAlignRight
    .ColumnsHeader.Add "DataDocumentoSocio", "Data doc. socio", dgDate, False, 2000, dgAlignleft
    .ColumnsHeader.Add "Acquistato", "Acquistato", dgBoolean, False, 2000, dgAligncenter
    .ColumnsHeader.Add "IDLottoProduzione", "IDLottoProduzione", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceLotto", "Codice lotto", dgchar, False, 1500, dgAlignleft
    .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, False, 1500, dgAlignleft
    .ColumnsHeader.Add "DescrizioneArticolo", "Articolo", dgchar, False, 3500, dgAlignleft
    Set cl = .ColumnsHeader.Add("PesoNettoCalcolato", "Quantità", dgDouble, True, 1300, dgAlignRight)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("ImportoUnitario", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("TotaleRiga", "Totale", dgDouble, True, 1300, dgAlignRight)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Form_Resize()
On Error GoTo ERR_Form_Resize
Me.GrigliaCorpo.Width = Me.Width - 240
Me.GrigliaCorpo.Height = Me.Height - 600
    
Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"
End Sub
