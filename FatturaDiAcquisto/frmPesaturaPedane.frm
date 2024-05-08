VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmPesaturaPedane 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesatura conferimento"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesaturaPedane.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   18600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Riferimenti documento di consegna della sessione di pesatura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   18375
      Begin VB.CommandButton cmdSbloccaNuoviIns 
         Caption         =   "Sblocca nuovi inserimenti"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtNDocSessione 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin DMTDATETIMELib.dmtDate txtDataDocSessione 
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "NUMERO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraInizio 
      Caption         =   "Inserimento/Modifica pesatura del conferimento selezionato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   18375
      Begin DMTDATETIMELib.dmtDate txtDataDoc 
         Height          =   495
         Left            =   10560
         TabIndex        =   8
         Top             =   600
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   873
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
      Begin VB.TextBox txtNumeroDoc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAnnotazioni 
         Height          =   525
         Left            =   13080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdElimina 
         Height          =   495
         Left            =   17640
         Picture         =   "frmPesaturaPedane.frx":4781A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Elimina pesatura"
         Top             =   600
         Width           =   615
      End
      Begin DMTEDITNUMLib.dmtNumber txtColliConf 
         Height          =   525
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   926
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordoConf 
         Height          =   525
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   926
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroPedaneConf 
         Height          =   525
         Left            =   6360
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   926
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPezziConf 
         Height          =   525
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   926
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Annotazioni"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   13080
         TabIndex        =   30
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Data doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   10560
         TabIndex        =   29
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "N° doc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   8280
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Pezzi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   260
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Colli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   260
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Peso lordo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   21
         Top             =   260
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "N° pedane"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   6360
         TabIndex        =   20
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA OPERAZIONE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9000
      Width           =   18375
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2520
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   9551
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      RowHeight       =   33
      ColumnsHeaderHeight=   40
   End
   Begin VB.Frame fraTotali 
      Caption         =   "Totali"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   7920
      Width           =   18375
      Begin DMTEDITNUMLib.dmtNumber txtColliTot 
         Height          =   405
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   714
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordoTot 
         Height          =   405
         Left            =   1920
         TabIndex        =   14
         Top             =   480
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   706
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroPedaneTot 
         Height          =   405
         Left            =   6360
         TabIndex        =   15
         Top             =   480
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   714
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPezziTot 
         Height          =   400
         Left            =   4320
         TabIndex        =   23
         Top             =   480
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   706
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label4 
         Caption         =   "Pezzi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Numero pedane"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Peso lordo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Numero colli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPesaturaPedane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsGrigliaSomma As ADODB.Recordset
Private QUANTITA_PER_COLLO As Long
Private ID As Long
Private IDRigaPesatura As Long
Private Conferma As Boolean
Private SbloccaNuoviIns As Boolean

Private Sub cmdConferma_Click()
    
    frmMain.txtQuantitaPedana.Value = Me.txtNumeroPedaneTot.Value
    frmMain.txtColli.Value = Me.txtColliTot.Value
    frmMain.txtPesoLordo.Value = Me.txtPesoLordoTot.Value
    frmMain.txtPezzi.Value = Me.txtPezziTot.Value
    
    CONFERMA_PESATURA
    
    CONFERMA_SALVA_PESATURA = 1
    Unload Me

End Sub
Private Sub CONFERMA_PESATURA()
On Error GoTo ERR_CONFERMA_PESATURA
Dim sSQL As String

sSQL = "UPDATE RV_POTMPPesaturaConf SET "
sSQL = sSQL & "Confermato=" & fnNormBoolean(1)
sSQL = sSQL & " WHERE IDOrdinamento=" & Link_Ordinamento_riga_conf
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"

Cn.Execute sSQL
Conferma = True
Exit Sub
ERR_CONFERMA_PESATURA:
    MsgBox Err.Description, vbCritical, "CONFERMA_PESATURA"
End Sub
Private Sub GET_ARTICOLO(IDArticolo As Long)
On Error GoTo ERR_GET_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo, RV_POQuantitaPerCollo "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    QUANTITA_PER_COLLO = 0
Else
    QUANTITA_PER_COLLO = fnNotNullN(rs!RV_POQuantitaPerCollo)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_ARTICOLO"
End Sub

Private Sub GET_TOTALI()
On Error GoTo ERR_GET_TOTALI
Dim I As Long
Dim TotaleColli As Long
Dim TotalePeso As Double
Dim TotalePezzi As Long
Dim TotalePedane As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

Me.txtColliTot.Value = 0
Me.txtPesoLordoTot.Value = 0
Me.txtPezziTot.Value = 0
Me.txtNumeroPedaneTot.Value = 0

sSQL = "SELECT SUM(Colli) as TotaleColli, "
sSQL = sSQL & "SUM(PesoLordo) as TotalePeso, "
sSQL = sSQL & "SUM(Pezzi) as TotalePezzi, "
sSQL = sSQL & "SUM(NumeroPedane) as TotalePedana "
sSQL = sSQL & "FROM RV_POTMPPesaturaConf "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(0)
sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtColliTot.Value = fnNotNullN(rs!TotaleColli)
    Me.txtPesoLordoTot.Value = fnNotNullN(rs!TotalePeso)
    Me.txtPezziTot.Value = fnNotNullN(rs!TotalePezzi)
    Me.txtNumeroPedaneTot.Value = fnNotNullN(rs!TotalePedana)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub

ERR_GET_TOTALI:
    MsgBox Err.Description, vbCritical, "GET_TOTALI"
End Sub

Private Sub cmdElimina_Click()
    EliminaRecord
    
End Sub

Private Sub cmdSbloccaNuoviIns_Click()
    If Len(Trim(Me.txtNDocSessione.Text)) = 0 Then
        MsgBox "Inserire il numero documento di consegna merce", vbCritical, "Validazione dati"
        Exit Sub
    End If
    If Me.txtDataDocSessione.Value = 0 Then
        MsgBox "Inserire la data di consegna merce", vbCritical, "Validazione dati"
        Exit Sub
    End If
    SbloccaNuoviIns = True
    NuovoRecord
End Sub

Private Sub Form_Activate()
    CONFERMA_SALVA_PESATURA = 0
    Me.txtDataDocSessione.Value = Date
    SbloccaNuoviIns = ATTIVA_OBBLIGO_N_DOC_SOCIO = 0
    
    NuovoRecord

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        SalvaRecord
    End If
    If (KeyCode = vbKeyEscape) Then
        NuovoRecord
    End If
    If ((KeyCode = vbKeyCancel) Or (KeyCode = vbKeyDelete)) Then
        EliminaRecord
    End If
    
End Sub

Private Sub Form_Load()
CONFERMA_SALVA_PESATURA = 0
GET_ARTICOLO frmMain.CDArticolo.KeyFieldID

GET_GRIGLIA

Conferma = False

End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
sSQL = "SELECT * FROM RV_POTMPPesaturaConf "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(False)
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
sSQL = sSQL & " ORDER BY DataUltimaModifica DESC"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
        Set cl = .ColumnsHeader.Add("Colli", "Colli", dgInteger, True, 2500, dgAlignRight)
            'cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, True, 2500, dgAlignRight)
            'cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgInteger, True, 2500, dgAlignRight)
            'cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        Set cl = .ColumnsHeader.Add("NumeroPedane", "Numero pedane", dgInteger, True, 2500, dgAlignRight)
            'cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 0
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("NumeroDocumentoConsegna", "Numero doc. cons.", dgchar, True, 1500, dgAlignleft)
        Set cl = .ColumnsHeader.Add("DataDocumentoConsegna", "Data doc. cons.", dgchar, True, 1500, dgAlignleft)
        Set cl = .ColumnsHeader.Add("Annotazioni", "Annotazioni", dgchar, True, 2500, dgAlignleft)
        Set cl = .ColumnsHeader.Add("DataUltimaModifica", "Data Ult. mod.", dgchar, True, 2500, dgAlignleft)
        Set cl = .ColumnsHeader.Add("DataInserimento", "Data inserimento", dgchar, True, 2500, dgAlignleft)
        .EnableRowColors = True
        .RowColors.Clear
        .RowColors.Add "NuovaPesatura", "IDRV_POCaricoMerceRighePes=0", vbYellow
        
        
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor


GET_TOTALI

Exit Sub

ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "ERR_GET_GRIGLIA"
End Sub
Private Sub SalvaRecord()
On Error GoTo ERR_SalvaRecord
Dim sSQL As String
Dim rsnew As ADODB.Recordset

If SbloccaNuoviIns = False Then Exit Sub

If ATTIVA_OBBLIGO_N_DOC_SOCIO = 1 Then
    If Len(Trim(Me.txtNumeroDoc.Text)) = 0 Then
        MsgBox "Inserire il numero documento di consegna merce", vbCritical, "Validazione dati"
        Exit Sub
    End If
    If Me.txtDataDoc.Value = 0 Then
        MsgBox "Inserire la data di consegna merce", vbCritical, "Validazione dati"
        Exit Sub
    End If
End If

sSQL = "SELECT * FROM RV_POTMPPesaturaConf "
sSQL = sSQL & "WHERE ID=" & ID

Set rsnew = New ADODB.Recordset
rsnew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsnew.EOF Then
    If SbloccaNuoviIns = False Then
        rsnew.Close
        Set rsnew = Nothing
        Exit Sub
    End If
    rsnew.AddNew
    rsnew!DataInserimento = Date
    rsnew!IDUtente = TheApp.IDUser
    rsnew!IDOrdinamento = Link_Ordinamento_riga_conf
    rsnew!IDRV_POCaricoMerceRighePes = 0
    rsnew!IDTipoDocumentoCoop = 2
End If

rsnew!Colli = txtColliConf.Value
rsnew!PesoLordo = txtPesoLordoConf.Value
rsnew!NumeroPedane = txtNumeroPedaneConf.Value
rsnew!Pezzi = txtPezziConf.Value
rsnew!DataUltimaModifica = Now
rsnew!Eliminato = False
rsnew!Modificato = True
rsnew!Annotazioni = Me.txtAnnotazioni.Text
rsnew!NumeroDocumentoConsegna = Me.txtNumeroDoc.Text
If Me.txtDataDoc.Value = 0 Then
    rsnew!DataDocumentoConsegna = Null
Else
    rsnew!DataDocumentoConsegna = Me.txtDataDoc.Value
End If
rsnew.Update

rsnew.Close
Set rsnew = Nothing

GET_GRIGLIA

NuovoRecord
Exit Sub
ERR_SalvaRecord:
    MsgBox Err.Description, vbCritical, "SalvaRecord"
End Sub

Private Sub NuovoRecord()
On Error GoTo ERR_NuovoRecord


If SbloccaNuoviIns = False Then Exit Sub
ID = 0
txtColliConf.Value = 0
txtPesoLordoConf.Value = 0
txtPezziConf.Value = 0
txtNumeroPedaneConf.Value = 1
Me.txtAnnotazioni.Text = ""
txtColliConf.SetFocus
txtNumeroDoc.Text = Me.txtNDocSessione.Text
txtDataDoc.Value = Me.txtDataDocSessione.Value

Exit Sub
ERR_NuovoRecord:
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sSQL As String

    If (Conferma = False) Then
        sSQL = "UPDATE RV_POTMPPesaturaConf SET "
        sSQL = sSQL & "Eliminato=" & fnNormBoolean(0)
        sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
        sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
        Cn.Execute sSQL
    End If
End Sub

Private Sub Griglia_Reposition(ByVal AllColumns As dmtgridctl.dgColumns)
On Error GoTo ERR_Griglia_Reposition
    ID = fnNotNullN(Me.Griglia.AllColumns("ID").Value)
    IDRigaPesatura = fnNotNullN(Me.Griglia.AllColumns("IDRV_POCaricoMerceRighePes").Value)
    txtColliConf.Value = fnNotNullN(Me.Griglia.AllColumns("Colli").Value)
    txtPesoLordoConf.Value = fnNotNullN(Me.Griglia.AllColumns("PesoLordo").Value)
    txtPezziConf.Value = fnNotNullN(Me.Griglia.AllColumns("Pezzi").Value)
    txtNumeroPedaneConf.Value = fnNotNullN(Me.Griglia.AllColumns("NumeroPedane").Value)
    txtAnnotazioni.Text = fnNotNull(Me.Griglia.AllColumns("Annotazioni").Value)
    txtNumeroDoc.Text = fnNotNull(Me.Griglia.AllColumns("NumeroDocumentoConsegna").Value)
    txtDataDoc.Value = fnNotNullN(Me.Griglia.AllColumns("DataDocumentoConsegna").Value)
Exit Sub
ERR_Griglia_Reposition:
    MsgBox Err.Description, vbCritical, "Griglia_Reposition"
End Sub
Private Sub EliminaRecord()
On Error GoTo ERR_EliminaRecord
Dim sSQL As String
Dim Testo  As String

If (ID = 0) Then Exit Sub

Testo = "ATTENZIONE!!!" & vbCrLf
Testo = Testo & "Sei sicura di eliminare la riga?"

If (MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione") = vbNo) Then Exit Sub

sSQL = "UPDATE RV_POTMPPesaturaConf SET "
sSQL = sSQL & "Eliminato=" & fnNormBoolean(1)
sSQL = sSQL & " WHERE ID=" & ID
Cn.Execute sSQL

GET_GRIGLIA

NuovoRecord

Exit Sub
ERR_EliminaRecord:
    MsgBox Err.Description, vbCritical, "EliminaRecord"
End Sub

Private Sub txtAnnotazioni_LostFocus()
On Error GoTo ERR_txtNumeroPedaneConf_LostFocus
Dim Testo As String
If (ID > 0) Then Exit Sub
If ((Me.txtColliConf.Value = 0) And (Me.txtPesoLordoConf.Value = 0)) Then Exit Sub

    Testo = "Vuoi salvare la pesatura?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio pesatura") = vbNo Then Exit Sub
    
    SalvaRecord

Exit Sub
ERR_txtNumeroPedaneConf_LostFocus:
    MsgBox Err.Description, vbCritical, "txtAnnotazioni_LostFocus"
End Sub

Private Sub txtColliConf_Change()
On Error GoTo ERR_txtColliConf_Change
    If QUANTITA_PER_COLLO > 0 Then
        Me.txtPezziConf.Value = txtColliConf.Value * QUANTITA_PER_COLLO
    End If
    If ATTIVA_CALCOLO_PESO_LORDO = 1 Then
        Me.txtPesoLordoConf.Value = Me.txtColliConf.Value * PESO_LORDO_ARTICOLO
        If TIPO_PESO_ARTICOLO = 2 Then
            Me.txtPesoLordoConf.Value = Me.txtPesoLordoConf.Value + (Me.txtColliConf.Value * frmMain.txtTaraUnitaria.Value)
        End If
    End If
Exit Sub
ERR_txtColliConf_Change:
    MsgBox Err.Description, vbCritical, "txtColliConf_Change"
End Sub

