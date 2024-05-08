VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmScontoOrdVivaio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sconto applicato"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   2490
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOrdineRigaOri 
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin DMTEDITNUMLib.dmtCurrency txtImportoListinoArticolo 
         Height          =   435
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   767
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   12648447
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
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtScontoImpListino 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   767
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   12648447
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
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtImportoApplicato 
         Height          =   435
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   767
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   12648447
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
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Importo listino"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Importo applicato"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Sconto applicato"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmScontoOrdVivaio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    If (IMPORTO_DA_LISTINO > 0) Then
        
        Me.txtScontoImpListino.Value = (1 - (frmMain.txtImponibileUnitario.Value / IMPORTO_DA_LISTINO)) * 100
    End If
    
    Me.txtImportoListinoArticolo.Value = IMPORTO_DA_LISTINO
    Me.txtImportoApplicato.Value = frmMain.txtImponibileUnitario.Value

End Sub
