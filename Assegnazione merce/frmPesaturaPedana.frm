VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmPesaturaPedana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesatura pedana"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesaturaPedana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin DMTDataCmb.DMTCombo DMTCombo1 
      Height          =   315
      Left            =   720
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DMTEDITNUMLib.dmtNumber txtPesatura 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtTotaleColliPedana 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtTotalePesoLordoPedana 
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtColliPesatura 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtPesoLordoLavorazione 
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtColliLavorazione 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtNumber txtPesoPedana 
      Height          =   315
      Left            =   3720
      TabIndex        =   13
      Top             =   760
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Tara pedana"
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   14
      Top             =   560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Colli lavorazione"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   12
      Top             =   2320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Peso lordo lav."
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   10
      Top             =   2320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Colli pedana"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   1720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Peso totale pedana"
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   1720
      Width           =   1695
   End
   Begin VB.Label lblPedana 
      Alignment       =   2  'Center
      Caption         =   "PEDANA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Totale peso lordo"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   1120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Totale colli pedana"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   1120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPesaturaPedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Variazione As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Variazione = True
        VARIAZIONE_DA_PESATURA = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    Me.txtPesoPedana.Value = frmMain.txtPesoPedana.Value
    Me.lblPedana.Caption = frmMain.txtCodicePedana.Text
    GET_TOTALI_PEDANA frmMain.txtIDPedana.Value
    Me.txtColliPesatura.Value = GET_COLLI_PER_IMBALLO(frmMain.txtIDPedana.Value, frmMain.CDTipoPedana.KeyFieldID, frmMain.CDCodiceImballo.KeyFieldID)
    Variazione = False
End Sub

Private Sub GET_TOTALI_PEDANA(IDPedana As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Colli) AS TotaleColli, "
sSQL = sSQL & "SUM(PesoLordo) AS TotalePesoLordo "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotaleColliPedana.Value = 0
    Me.txtTotalePesoLordoPedana.Value = 0
Else
    Me.txtTotaleColliPedana.Value = fnNotNullN(rs!TotaleColli)
    Me.txtTotalePesoLordoPedana.Value = fnNotNullN(rs!TotalePesoLordo)
End If

rs.CloseResultset
Set rs = Nothing


End Sub
Private Function GET_COLLI_PER_IMBALLO(IDPedana As Long, IDTipoPedana As Long, IDArticoloImballo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LOCAL_QUANTITA As Double
Dim LOCAL_QUANTITA_LAVORATA As Double
Dim Testo As String

sSQL = "SELECT Quantita FROM RV_POTipoPedanaImballo "
sSQL = sSQL & "WHERE IDRV_POTipoPedana=" & IDTipoPedana
sSQL = sSQL & " AND IDArticolo=" & IDArticoloImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_COLLI_PER_IMBALLO = 0
Else
    GET_COLLI_PER_IMBALLO = fnNotNullN(rs!Quantita)
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Sub Form_Unload(Cancel As Integer)
    If Variazione = True Then
        frmMain.txtPesoPedana.Value = Me.txtPesoPedana.Value
        frmMain.txtColli.Value = Me.txtColliLavorazione.Value
        frmMain.txtTara.Value = frmMain.txtColli.Value * frmMain.txtTaraUnitaria.Value
        frmMain.txtPesoLordo.Value = Me.txtPesoLordoLavorazione.Value
        frmMain.txtPesoNetto.Value = frmMain.txtPesoLordo.Value - frmMain.txtTara.Value
        CalcoloPesoNetto
    End If
End Sub
Private Sub CalcoloPesoNetto()
On Error Resume Next
Dim ArrayPesoNetto() As String
Dim PesoNetto As Double
Dim Decimal_PesoNetto As Double
Dim Tara As Double


'frmMain.txtTara.Value = frmMain.txtTaraUnitaria.Value * frmMain.txtColli.Value
    



Select Case Link_Arrontondamento

    Case 1 'Nessun arrotondamento
        frmMain.txtPesoNetto.Value = frmMain.txtPesoLordo.Value - frmMain.txtTara.Value
    Case 2 'Matematico
        ArrayPesoNetto() = Split(frmMain.txtPesoNetto.Text, ",")
        PesoNetto = FormatNumber(frmMain.txtPesoNetto.Value, 0)
        
        frmMain.txtPesoNetto.Value = PesoNetto
        frmMain.txtPesoLordo.Value = frmMain.txtTara.Value + frmMain.txtPesoNetto.Value
    Case 3 'Difetto
        ArrayPesoNetto() = Split(frmMain.txtPesoNetto.Text, ",")
        PesoNetto = Int(frmMain.txtPesoNetto.Value)
        
        frmMain.txtPesoNetto.Value = PesoNetto
        frmMain.txtPesoLordo.Value = frmMain.txtPesoNetto.Value + frmMain.txtTara.Value
        
    Case 4 'Eccesso
        ArrayPesoNetto() = Split(frmMain.txtPesoNetto.Text, ",")
        PesoNetto = Int(frmMain.txtPesoNetto.Value)
        If ArrayPesoNetto(1) > 0 Then
            frmMain.txtPesoNetto.Value = PesoNetto + 1
        Else
            frmMain.txtPesoNetto.Value = PesoNetto
        End If
        
        frmMain.txtPesoLordo.Valuef = rmMain.txtTara.Value + frmMain.txtPesoNetto.Value
        
End Select
End Sub
Private Sub txtColliPesatura_Change()
    'If (Me.txtColliPesatura.Value - Me.txtTotaleColliPedana.Value) >= 0 Then
        Me.txtColliLavorazione.Value = Abs(Me.txtColliPesatura.Value - Me.txtTotaleColliPedana.Value)
    'Else
    '    Me.txtColliLavorazione.Value = 0
    'End If
End Sub

Private Sub txtPesatura_Change()
    If (Me.txtPesatura.Value - (Me.txtTotalePesoLordoPedana.Value + Me.txtPesoPedana.Value)) >= 0 Then
        Me.txtPesoLordoLavorazione.Value = Me.txtPesatura.Value - (Me.txtTotalePesoLordoPedana.Value + Me.txtPesoPedana.Value)
    Else
        Me.txtPesoLordoLavorazione.Value = 0
    End If
End Sub

Private Sub txtPesoPedana_Change()
    txtPesatura_Change
End Sub
