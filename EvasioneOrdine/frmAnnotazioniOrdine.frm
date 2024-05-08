VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.0#0"; "DMTDataCmb.OCX"
Begin VB.Form frmAnnotazioniOrdine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Altri dati ordine"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10080
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
   ScaleHeight     =   3210
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalva 
      Caption         =   "AGGIORNA"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtDescrizioneRigaDoc 
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtNumeroOrdine 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtDataOrdine 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Numero ordine"
         Height          =   255
         Index           =   2
         Left            =   7920
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Data ordine"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   5775
      End
   End
   Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
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
   Begin DMTDataCmb.DMTCombo cboVettoreSuccessivo 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.Label Label4 
      Caption         =   "Luogo di presa merce"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Vettore successivo"
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Annotazioni finali del corpo del documento di evasione"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   9855
   End
End
Attribute VB_Name = "frmAnnotazioniOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSalva_Click()


    AGGIORNA_GRIGLIA_ORDINI_PER_EVASIONE = True

    LINK_LUOGO_MERCE_PER_EVASIONE = Me.cboLuogoPresaMerce.CurrentID
    LINK_VETTORE_SUCCESSIVO_PER_EVASIONE = Me.cboVettoreSuccessivo.CurrentID
    DESCRIZIONE_CORPO_PER_EVASIONE = Me.txtDescrizioneRigaDoc.Text

    Unload Me
End Sub

Private Sub Form_Load()
    With Me.cboVettoreSuccessivo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .Sql = "SELECT * FROM Vettore ORDER BY Vettore"
        .Fill
    End With

    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica  "
        .Sql = .Sql & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
    Me.txtCliente.Text = fnNotNull(FrmFine.GrigliaOrdini("Cliente").Value)
    Me.txtDataOrdine.Text = fnNotNull(FrmFine.GrigliaOrdini("DataOrdine").Value)
    Me.txtNumeroOrdine.Text = fnNotNullN(FrmFine.GrigliaOrdini("NumeroOrdine").Value)
    
    Me.cboLuogoPresaMerce.WriteOn fnNotNullN(FrmFine.GrigliaOrdini("IDLuogoPresaMerce").Value)
    Me.cboVettoreSuccessivo.WriteOn fnNotNullN(FrmFine.GrigliaOrdini("IDVettoreSuccessivo").Value)
    Me.txtDescrizioneRigaDoc.Text = fnNotNull(FrmFine.GrigliaOrdini("DescrizioneCorpoDocEv").Value)
    

End Sub
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
End If


rs.CloseResultset
Set rs = Nothing
End Function
