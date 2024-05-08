VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Begin VB.Form frmSalvaComeNuovo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALVA COME NUOVO"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalvaComeNuovo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "RIEPILOGO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
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
         Left            =   1680
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CheckBox chkRiportaSerre 
         Caption         =   "Riporta serre/appezzamenti"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CheckBox chkRiportaProdotti 
         Caption         =   "Riporta prodotti"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   3255
      End
      Begin DMTDataCmb.DMTCombo cboPeriodoCampagna 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboStatoLotto 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDATETIMELib.dmtDate txtDataSemina 
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Data semina pres."
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Stato del lotto"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo di campagna"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmSalvaComeNuovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
    SALVA_COME_NUOVO = True
    
    RIPORTA_LINK_PERIODO_CAMPAGNA = Me.cboPeriodoCampagna.CurrentID
    RIPORTA_LINK_STATO_LOTTO = Me.cboStatoLotto.CurrentID
    DATA_SEMINA_SCN = Me.txtDataSemina.Text
    RIPORTA_PRODOTTI = Me.chkRiportaProdotti.Value
    RIPORTA_SERRE_APPEZZAMENTI = Me.chkRiportaSerre.Value
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Init_Controlli
Me.cboPeriodoCampagna.WriteOn frmMain.cboPeriodoCampagna.CurrentID
Me.cboStatoLotto.WriteOn frmMain.cboStatoLotto.CurrentID
Me.txtDataSemina.Value = frmMain.txtDataSemina.Value

End Sub
Private Sub Init_Controlli()
    'Periodo di campagna
    With Me.cboPeriodoCampagna
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_PO01_PeriodoCampagna"
        .DisplayField = "PeriodoCampagna"
        .SQL = "SELECT * FROM RV_PO01_PeriodoCampagna "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND IDFiliale=" & TheApp.Branch
        .SQL = .SQL & " ORDER BY IDRV_PO01_PeriodoCampagna"
        .Fill
    End With

    'Stato del lotto
    With Me.cboStatoLotto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_PO01_StatoLotto"
        .DisplayField = "StatoLotto"
        .SQL = "SELECT * FROM RV_PO01_StatoLotto ORDER BY StatoLotto"
        .Fill
    End With
End Sub

