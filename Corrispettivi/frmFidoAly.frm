VERSION 5.00
Begin VB.Form frmFidoAly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIDO ALYANTE"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFidoAly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEsposizioneDDT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtEsposizioneAly 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   7800
      Width           =   1170
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Esposizione"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtEsposizioneNC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4680
         Width           =   2415
      End
      Begin VB.TextBox txtEsposizioneND 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox txtTotaleDocumentoPrec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   5880
         Width           =   2415
      End
      Begin VB.TextBox txtResiduoFinale 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   7080
         Width           =   2415
      End
      Begin VB.TextBox txtEsposizioneFinale 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6480
         Width           =   2415
      End
      Begin VB.TextBox txtDocumentoCorrente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox txtEsposizioneFD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtEsposizioneFA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox txtResiduoAly 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtFidoCliente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione per N.C.:"
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
         Index           =   11
         Left            =   120
         TabIndex        =   26
         Top             =   4760
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione per N.D.:"
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
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   4160
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Totale documento prec.:"
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
         Index           =   9
         Left            =   120
         TabIndex        =   22
         Top             =   5955
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Residuo complessivo:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   7155
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione complessiva:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   6555
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Totale documento corrente:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   5355
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione per F.D.:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   3560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione per F.A.:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2960
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione per D.D.T.:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Residuo in Alyante:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1760
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Esposizione in Alyante:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1160
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Fido cliente:"
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
         TabIndex        =   3
         Top             =   560
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmFidoAly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ALY_CONFERMA_SALVA_DOC = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    ALY_CONFERMA_SALVA_DOC = False
    
    Me.txtFidoCliente.Text = FormatNumber(ALY_FIDO_CLIENTE, 2)
    Me.txtEsposizioneAly.Text = FormatNumber(ALY_FIDO_CALCOLATO_ALY, 2)
    Me.txtResiduoAly.Text = FormatNumber(ALY_FIDO_RESIDUO_ALY, 2)
    Me.txtEsposizioneDDT.Text = FormatNumber(ALY_FIDO_TOT_DDT, 2)
    Me.txtEsposizioneFA.Text = FormatNumber(ALY_FIDO_TOT_FA, 2)
    Me.txtEsposizioneFD.Text = FormatNumber(ALY_FIDO_TOT_FD, 2)
    Me.txtEsposizioneND.Text = FormatNumber(ALY_FIDO_TOT_ND, 2)
    Me.txtEsposizioneNC.Text = FormatNumber(-ALY_FIDO_TOT_NC, 2)
    Me.txtDocumentoCorrente.Text = FormatNumber(ALY_TOTALE_DOC, 2)
    Me.txtTotaleDocumentoPrec.Text = FormatNumber(-ALY_TOTALE_DOC_PREC, 2)
    Me.txtEsposizioneFinale.Text = FormatNumber(ALY_FIDO_CALCOLATO, 2)
    Me.txtResiduoFinale.Text = FormatNumber(ALY_FIDO_RESIDUO, 2)
    
    If (ALY_FIDO_RESIDUO < 0) Then
        Me.txtResiduoFinale.BackColor = vbRed
        Me.txtResiduoFinale.ForeColor = vbWhite
    Else
        Me.txtResiduoFinale.BackColor = vbGreen
        Me.txtResiduoFinale.ForeColor = vbWhite
    End If
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
