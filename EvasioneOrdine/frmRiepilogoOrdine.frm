VERSION 5.00
Begin VB.Form frmRiepilogoOrdine 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RIEPILOGO ORDINE"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNumeroColli 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox txtNumeroPedana 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NUMERO DI COLLI"
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
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NUMERO DI PEDANE"
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmRiepilogoOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

    Me.txtNumeroPedana.Text = GET_NUMERO_PEDANE_ORDINE(FrmMain.txtIDOrdine.Value)
    Me.txtNumeroColli.Text = GET_NUMERO_COLLI_ORDINE(FrmMain.txtIDOrdine.Value)
    
    FrmMain.SetFocus
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = Screen.Width - Me.Width - 100
    
End Sub


Public Function GET_NUMERO_PEDANE_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_PEDANE_ORDINE = 0

sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_PEDANE_ORDINE = GET_NUMERO_PEDANE_ORDINE + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Public Function GET_NUMERO_COLLI_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_COLLI_ORDINE = 0

sSQL = "SELECT Sum(Colli) AS NumeroColli "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_COLLI_ORDINE = 0
Else
    GET_NUMERO_COLLI_ORDINE = fnNotNullN(rs!NumeroColli)
    
End If
rs.CloseResultset
Set rs = Nothing
End Function

