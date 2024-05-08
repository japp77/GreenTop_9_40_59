VERSION 5.00
Begin VB.Form FrmLicenza 
   Caption         =   "Controllo licenza"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLicenza 
      Height          =   4815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "FrmLicenza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.txtLicenza.Text = StringaLicenza
   
End Sub

