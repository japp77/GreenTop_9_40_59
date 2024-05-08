VERSION 5.00
Begin VB.Form frmOraArrivoMerce 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ora arrivo merce"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSecondi 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtMinuti 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtOra 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   240
      MaxLength       =   2
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   4440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SECONDI"
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
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MINUTI"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ORA"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmOraArrivoMerce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(Me.txtOra.Text) = 1 Then
    Ora = "0" & Me.txtOra.Text
Else
    Ora = Me.txtOra.Text
End If
If Len(Me.txtMinuti.Text) = 1 Then
    Minuti = "0" & Me.txtMinuti.Text
Else
    Minuti = Me.txtMinuti.Text
End If
If Len(Me.txtSecondi.Text) = 1 Then
    Secondi = "0" & Me.txtSecondi.Text
Else
    Secondi = Me.txtSecondi.Text
End If
frmMain.txtOraArrivo.Text = Ora & "." & Minuti & "." & Secondi

Unload Me


End Sub

Private Sub Form_Load()
    GET_FORMATTA_ORARIO
End Sub

Private Sub GET_FORMATTA_ORARIO()
On Error Resume Next
Dim ArrayOrario() As String

ArrayOrario = Split(frmMain.txtOraArrivo.Text, ".")

Me.txtOra.Text = ArrayOrario(0)
Me.txtMinuti.Text = ArrayOrario(1)
Me.txtSecondi.Text = ArrayOrario(2)
End Sub
Private Sub txtMinuti_GotFocus()
    Me.txtMinuti.SelLength = Len(Me.txtMinuti.Text)
End Sub

Private Sub txtOra_GotFocus()
    Me.txtOra.SelLength = Len(Me.txtOra.Text)
End Sub

Private Sub txtSecondi_GotFocus()
    Me.txtSecondi.SelLength = Len(Me.txtSecondi.Text)
End Sub
