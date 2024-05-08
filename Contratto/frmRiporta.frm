VERSION 5.00
Begin VB.Form frmRiporta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIPORTA ALTRE IMPOSTAZIONI"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiporta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Riporta"
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
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox Check16 
         Caption         =   "Altra destinazione"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Luogo presa merce"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Causale di trasporto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Tipo ordine"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Annotazioni interne"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   5040
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Annotazioni finali del corpo del documento di evasione"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   5760
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Annotazioni di fatturazione"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   5400
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Aspetto esteriore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Vettore successivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Istruzioni per il mittente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Targa automezzo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Vettore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Tipo trasporto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Porto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Modalità di pagamento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Value           =   1  'Checked
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmRiporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()

    RIP_AGENTE_CONTR = Abs(Check1.Value)
    RIP_ALTRA_DEST_CONTR = Abs(Check16.Value)
    RIP_ASPETTO_EST_CONTR = Abs(Check9.Value)
    RIP_CAUS_CONTR = Abs(Check14.Value)
    RIP_ISTRUZIONI_CONTR = Abs(Check7.Value)
    RIP_LUOGO_MERCE_CONTR = Abs(Check15.Value)
    RIP_NOTE_FATT_CONTR = Abs(Check10.Value)
    RIP_NOTE_FINALI_CONTR = Abs(Check11.Value)
    RIP_NOTE_INTERNE_CONTR = Abs(Check12.Value)
    RIP_PAGAMENTO_CONTR = Abs(Check2.Value)
    RIP_PORTO_CONTR = Abs(Check3.Value)
    RIP_TARGA_CONTR = Abs(Check6.Value)
    RIP_TIPO_ORD_CONTR = Abs(Check13.Value)
    RIP_TRASPORTO_CONTR = Abs(Check4.Value)
    RIP_VETTORE_CONTR = Abs(Check5.Value)
    RIP_VETTORE_SUCC_CONTR = Abs(Check8.Value)
    
    CONF_RIPORTA_CONTRATTO = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    CONF_RIPORTA_CONTRATTO = False
End Sub
