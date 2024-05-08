VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Begin VB.Form frmAltriDati 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALTRI DATI DEL DOCUMENTO"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAltriDati.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
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
      Left            =   8280
      TabIndex        =   9
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Annotazioni 3"
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
      Height          =   2415
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   6135
      Begin VB.CommandButton cmdNota3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5640
         TabIndex        =   12
         Top             =   0
         Width           =   420
      End
      Begin VB.TextBox txtAnnotazioni03 
         Height          =   1935
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Annotazioni 2"
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
      Height          =   2415
      Left            =   4320
      TabIndex        =   3
      Top             =   2640
      Width           =   6135
      Begin VB.CommandButton cmdNota2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5640
         TabIndex        =   11
         Top             =   0
         Width           =   420
      End
      Begin VB.TextBox txtAnnotazioni02 
         Height          =   1935
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Annotazioni 1"
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
      Height          =   2415
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdNote1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5640
         TabIndex        =   10
         Top             =   0
         Width           =   420
      End
      Begin VB.TextBox txtAnnotazioni01 
         Height          =   1935
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documento"
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
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin DMTDATETIMELib.dmtDate txtDataCompetenzaLiq 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   450
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data competenza liq."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmAltriDati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
    
    Conferma_dati
    
    
    CONFERMA_ALTRI_DATI = 1
    
    Unload Me
End Sub

Private Sub cmdNota2_Click()
    LINK_TIPO_NOTA_SEL = 2
    frmElencoNote.Show vbModal
    
    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazioni02.Text = RETURN_RIGA_NOTA
    End If
    
    
End Sub

Private Sub cmdNota3_Click()
    LINK_TIPO_NOTA_SEL = 3
    frmElencoNote.Show vbModal

    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazioni03.Text = RETURN_RIGA_NOTA
    End If

End Sub

Private Sub cmdNote1_Click()
    LINK_TIPO_NOTA_SEL = 1
    frmElencoNote.Show vbModal
    
    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazioni01.Text = RETURN_RIGA_NOTA
    End If
    
End Sub

Private Sub Form_Load()
    CONFERMA_ALTRI_DATI = 0
    
    Recupera_dati
End Sub
Private Sub Recupera_dati()
    
    Me.txtDataCompetenzaLiq.Text = DATA_COMPETENZA_LIQ
    Me.txtAnnotazioni01.Text = ANNOTAZIONE_01
    Me.txtAnnotazioni02.Text = ANNOTAZIONE_02
    Me.txtAnnotazioni03.Text = ANNOTAZIONE_03
    
End Sub
Private Sub Conferma_dati()
    DATA_COMPETENZA_LIQ = Me.txtDataCompetenzaLiq.Text
    ANNOTAZIONE_01 = Me.txtAnnotazioni01.Text
    ANNOTAZIONE_02 = Me.txtAnnotazioni02.Text
    ANNOTAZIONE_03 = Me.txtAnnotazioni03.Text
    
End Sub
