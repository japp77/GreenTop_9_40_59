VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRiepilogo 
   Caption         =   "Riepilogo riga di conferimento"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13500
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
   ScaleHeight     =   7935
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   3120
      Width           =   480
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   600
      Left            =   8280
      TabIndex        =   1
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Frame FraAggMovLav 
         Caption         =   "Aggiornamento movimenti di lavorazione"
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   12615
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label2 
            Caption         =   "lblInfo"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   12375
         End
      End
      Begin VB.CommandButton cmdAggiornaMovLavorazione 
         Caption         =   "AGGIORNA"
         Height          =   375
         Left            =   10080
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDifferenza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtDifferenzaLavorazione 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtTotaleQuantitaQuadrata 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTotaleQuantitaLavorata 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTotaleQuantitaVendita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTotaleConferimento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtQuantitaVendita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtQuantitaLavorata 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtQuantitaQuadrata 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtQuantitaConferita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCodiceArticolo 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtDescrizioneDocumento 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtCausale 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblTotali 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTALI"
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
         Left            =   5400
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblInfo 
         Caption         =   "CARICAMENTO IN CORSO.............."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   2760
         Width           =   12495
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   11280
         X2              =   11280
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   9960
         X2              =   9960
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   8640
         X2              =   8640
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label lblDifferenzaLavorazione 
         Alignment       =   1  'Right Justify
         Caption         =   "DIFFERENZE"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lblDifferenza 
         Alignment       =   1  'Right Justify
         Caption         =   "DIFFERENZA TRA VENDITA E CONFERIMENTO"
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
         Left            =   5760
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Line LineFine 
         X1              =   120
         X2              =   12600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   7320
         X2              =   7320
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5400
         X2              =   5400
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   12600
         Y1              =   280
         Y2              =   280
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   840
         X2              =   840
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Causale"
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
         Left            =   0
         TabIndex        =   3
         Top             =   40
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
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
         Left            =   960
         TabIndex        =   4
         Top             =   40
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Codice articolo"
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
         Left            =   5520
         TabIndex        =   5
         Top             =   40
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Q.tà lav."
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
         Left            =   10080
         TabIndex        =   6
         Top             =   40
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Q.tà vend."
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
         Left            =   11400
         TabIndex        =   12
         Top             =   40
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Q.tà Quad."
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
         Left            =   8760
         TabIndex        =   14
         Top             =   40
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Q.tà Conf."
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
         Left            =   7440
         TabIndex        =   13
         Top             =   40
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRiepilogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TotaleTop As Long
Private Const VarTop As Long = 360
Private TotaleConferimento As Double
Private TotaleQuadratura As Double
Private TotaleVendita As Double
Private TotaleDebito As Double
Private TotaleNotaCredito As Double
Private TotaleAssegnazione As Double
Private TotaleProcesso As Double
Private IControl As Integer



Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
    Me.lblInfo.Visible = True
    Configurazione
    Me.lblInfo.Visible = False
    Form_Resize
Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    Me.Height = Me.Pic1.Height
    Me.Width = Me.Pic1.Width
    
    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
    
        'Me.Pic1.Height = Me.Height
        'Me.Pic1.Width = Me.Width
        
        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            If (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height) <= 32767 Then
                .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            Else
                .Max = 32727
            End If
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If

End Sub

Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub Configurazione()
    IControl = 0
    TotaleConferimento = 0
    TotaleQuadratura = 0
    TotaleVendita = 0
    TotaleDebito = 0
    TotaleNotaCredito = 0
    TotaleAssegnazione = 0
    TotaleProcesso = 0
    TotaleTop = 0
    
    
    CaricaConferimento
    
    DoEvents
    
    CaricaLavorazione
    CaricaAssegnazione
    CaricaProcesso
    CaricaVenditaDDT
    If ATTIVAZIONE_NUOVO_CALCOLO = False Then
        CaricaVenditaFA
        CaricaVenditaSNF
    End If
    
    Me.Line1(0).Y2 = TotaleTop + VarTop
    Me.Line1(1).Y2 = TotaleTop + VarTop
    Me.Line1(2).Y2 = TotaleTop + VarTop
    Me.Line1(3).Y2 = TotaleTop + VarTop
    Me.Line1(4).Y2 = TotaleTop + VarTop
    Me.Line1(5).Y2 = TotaleTop + VarTop
    Me.LineFine.Y1 = TotaleTop + VarTop
    Me.LineFine.Y2 = TotaleTop + VarTop

    TotaleTop = TotaleTop + (VarTop + (VarTop / 2))
    
    Me.txtTotaleConferimento.Top = TotaleTop
    Me.txtTotaleQuantitaQuadrata.Top = TotaleTop
    Me.txtTotaleQuantitaLavorata.Top = TotaleTop
    Me.txtTotaleQuantitaVendita.Top = TotaleTop
    Me.lblTotali.Top = TotaleTop
    Me.txtTotaleConferimento.Text = FormatNumber(TotaleConferimento, 2)
    Me.txtTotaleQuantitaLavorata.Text = FormatNumber((TotaleAssegnazione + TotaleProcesso), 2)
    Me.txtTotaleQuantitaQuadrata.Text = FormatNumber(TotaleQuadratura, 2)
    Me.txtTotaleQuantitaVendita.Text = FormatNumber(TotaleVendita, 2)
    
    
    'TotaleTop = TotaleTop + (VarTop + (VarTop / 2))
    
    TotaleTop = TotaleTop + VarTop
    
    
    ''''''DIFFERENZA TRA LAVORAZIONE E CONFERIMENTO'''''''''''''''''''''''''''''''''''''''''''
    
    Me.txtDifferenzaLavorazione.Top = TotaleTop
    
    Me.txtDifferenzaLavorazione.Text = FormatNumber((TotaleConferimento - (TotaleQuadratura + TotaleAssegnazione + TotaleProcesso)), 2)
    
    If Me.txtDifferenzaLavorazione.Text <> 0 Then
        Me.txtDifferenzaLavorazione.BackColor = vbRed
        Me.txtDifferenzaLavorazione.ForeColor = vbBlack
        
    Else
        Me.txtDifferenzaLavorazione.BackColor = vbGreen
        Me.txtDifferenzaLavorazione.ForeColor = vbBlack
    End If
    
    Me.lblDifferenzaLavorazione.Top = Me.txtDifferenzaLavorazione.Top
    Me.cmdAggiornaMovLavorazione.Top = Me.lblDifferenzaLavorazione.Top + 480
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''DIFFERENZA TRA VENDITA E CONFERIMENTO'''''''''''''''''''''''''''''''''''''''''''''
    'TotaleTop = TotaleTop + VarTop
    Me.txtDifferenza.Text = FormatNumber((TotaleConferimento - (TotaleQuadratura + TotaleVendita)), 2)
    If Me.txtDifferenza.Text <> 0 Then
        Me.txtDifferenza.BackColor = vbRed
        Me.txtDifferenza.ForeColor = vbBlack
        
    Else
        Me.txtDifferenza.BackColor = vbGreen
        Me.txtDifferenza.ForeColor = vbBlack
    End If
    Me.txtDifferenza.Top = TotaleTop
    Me.lblDifferenza.Top = Me.txtDifferenza.Top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TotaleTop = TotaleTop + (VarTop + (VarTop / 2))
    
    CaricaNotaDiCredito
    
    CaricaNotaDiDebito
    
    Me.Pic1.Height = TotaleTop + 480
    
    
    
End Sub
Private Sub CaricaConferimento()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceTesta.DataDocumento, RV_POCaricoMerceTesta.Anagrafica, "
sSQL = sSQL & "RV_POCaricoMerceTesta.Nome, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceRighe.Qta_UM "
sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not (rs.EOF) Then
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop
        
        Load Me.txtCausale(IControl)
        With Me.txtCausale(IControl)
            .Left = 120
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = "C"
            .BackColor = vbRed
            .ForeColor = vbYellow
        End With
        
        Load Me.txtDescrizioneDocumento(IControl)
        With Me.txtDescrizioneDocumento(IControl)
            .Left = 960
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbRed
            .ForeColor = vbYellow

            .Text = "N° " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " del socio " & fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
        End With
        
        Load Me.txtCodiceArticolo(IControl)
        With Me.txtCodiceArticolo(IControl)
            .Left = 5520
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbRed
            .ForeColor = vbYellow
            .Text = fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo)
        End With
        
        Load Me.txtQuantitaConferita(IControl)
        With Me.txtQuantitaConferita(IControl)
            .Left = 7440
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = FormatNumber(fnNotNullN(rs!Qta_UM), 2)
            .BackColor = vbRed
            .ForeColor = vbYellow
            TotaleConferimento = TotaleConferimento + fnNotNullN(rs!Qta_UM)
        End With

End If
rs.CloseResultset
Set rs = Nothing


End Sub

Private Sub CaricaLavorazione()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    sSQL = "SELECT * FROM RV_POLavorazione WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        Select Case Link_UnitaDiMisura_Coop_Conferimento
            Case 1
                If fnNotNull(rs!SegnoMovimento) = "+" Then
                    QuantitaMovimentata = fnNotNullN(-rs!Colli)
                Else
                    QuantitaMovimentata = fnNotNullN(rs!Colli)
                End If
            Case 2
                If fnNotNull(rs!SegnoMovimento) = "+" Then
                    QuantitaMovimentata = fnNotNullN(-rs!PesoLordo)
                Else
                    QuantitaMovimentata = fnNotNullN(rs!PesoLordo)
                End If
            Case 3
                If fnNotNull(rs!SegnoMovimento) = "+" Then
                    QuantitaMovimentata = fnNotNullN(-rs!PesoNetto)
                Else
                    QuantitaMovimentata = fnNotNullN(rs!PesoNetto)
                End If
            Case 4
                If fnNotNull(rs!SegnoMovimento) = "+" Then
                    QuantitaMovimentata = fnNotNullN(-rs!Tara)
                Else
                    QuantitaMovimentata = fnNotNullN(rs!Tara)
                End If
            Case 5
                If fnNotNull(rs!SegnoMovimento) = "+" Then
                    QuantitaMovimentata = fnNotNullN(-rs!Pezzi)
                Else
                    QuantitaMovimentata = fnNotNullN(rs!Pezzi)
                End If
        End Select
    
    
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "S"
                .BackColor = vbBlue
                .ForeColor = vbYellow
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = "Quadratura del " & fnNotNull(rs!DataDocumento)
                
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo)
                
            End With
            
            Load Me.txtQuantitaQuadrata(IControl)
            With Me.txtQuantitaQuadrata(IControl)
                .Left = 8760
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(QuantitaMovimentata, 2)
                .BackColor = vbBlue
                .ForeColor = vbYellow
                TotaleQuadratura = TotaleQuadratura + QuantitaMovimentata
                
            End With
            DoEvents
            
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
    Exit Sub
End If


''''''''''''''''NUOVO CALCOLO
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POLavorazioneL")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


        Load Me.txtCausale(IControl)
        With Me.txtCausale(IControl)
            .Left = 120
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = "S"
            .BackColor = vbBlue
            .ForeColor = vbYellow
        End With
        
        Load Me.txtDescrizioneDocumento(IControl)
        With Me.txtDescrizioneDocumento(IControl)
            .Left = 960
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbBlue
            .ForeColor = vbYellow
            .Text = fnNotNull(rs!Oggetto)
        End With
        
        Load Me.txtCodiceArticolo(IControl)
        With Me.txtCodiceArticolo(IControl)
            .Left = 5520
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbBlue
            .ForeColor = vbYellow
            .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
        End With
        
        Load Me.txtQuantitaQuadrata(IControl)
        With Me.txtQuantitaQuadrata(IControl)
            Select Case GET_TIPO_PRODOTTO(rs!IDArticolo)
            
                Case Link_TipoCaloPeso
                    QuantitaMovimentata = fnNotNullN(rs!RV_POQuantitaMovimentata)
                Case Link_TipoScarto
                    QuantitaMovimentata = fnNotNullN(rs!RV_POQuantitaMovimentata)
                Case Link_TipoAumentoPeso
                    QuantitaMovimentata = fnNotNullN(-rs!RV_POQuantitaMovimentata)
                Case Else
                    QuantitaMovimentata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            End Select
            
            'If RecuperaSegnoPerDisponibilita(fnNotNullN(rs!IDFunzione)) = "+" Then
            '    QuantitaMovimentata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            'Else
            '    QuantitaMovimentata = fnNotNullN(-rs!RV_POQuantitaMovimentata)
            'End If
        
            .Left = 8760
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = FormatNumber(QuantitaMovimentata, 2)
            .BackColor = vbBlue
            .ForeColor = vbYellow
            TotaleQuadratura = TotaleQuadratura + QuantitaMovimentata
        End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub CaricaAssegnazione()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    sSQL = "SELECT * FROM RV_POAssegnazioneMerce WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        Select Case Link_UnitaDiMisura_Coop_Conferimento
            Case 1
                QuantitaMovimentata = fnNotNullN(rs!Colli)
            Case 2
                QuantitaMovimentata = fnNotNullN(rs!PesoLordo)
            Case 3
                QuantitaMovimentata = fnNotNullN(rs!PesoNetto)
            Case 4
                QuantitaMovimentata = fnNotNullN(rs!Tara)
            Case 5
                QuantitaMovimentata = fnNotNullN(rs!Pezzi)
        End Select
    
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "L"
                .BackColor = vbYellow
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
    
                .Text = "Lavorazione del " & fnNotNull(rs!DataDocumento)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo)
            End With
            
            Load Me.txtQuantitaLavorata(IControl)
            With Me.txtQuantitaLavorata(IControl)
                .Left = 10080
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(QuantitaMovimentata, 2)
                .BackColor = vbYellow
                .ForeColor = vbBlack
                TotaleAssegnazione = TotaleAssegnazione + QuantitaMovimentata
            End With
            DoEvents
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
    Exit Sub
End If

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "L"
                .BackColor = vbYellow
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!Oggetto)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
            End With
            
            Load Me.txtQuantitaLavorata(IControl)
            With Me.txtQuantitaLavorata(IControl)
                .Left = 10080
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!RV_POQuantitaMovimentata), 2)
                .BackColor = vbYellow
                .ForeColor = vbBlack
                TotaleAssegnazione = TotaleAssegnazione + fnNotNullN(rs!RV_POQuantitaMovimentata)
            End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing

End Sub

Private Sub CaricaProcesso()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double

If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    sSQL = "SELECT RV_POProcessoIVGammaRighe.*, RV_POProcessoIVGamma.AnnoProcesso, RV_POProcessoIVGamma.NumeroProcesso "
    sSQL = sSQL & "FROM RV_POProcessoIVGamma INNER JOIN "
    sSQL = sSQL & "RV_POProcessoIVGammaRighe ON "
    sSQL = sSQL & "RV_POProcessoIVGamma.IDRV_POProcessoIVGamma = RV_POProcessoIVGammaRighe.IDRV_POProcessoIVGamma "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento
    
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
            
            
    
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "P"
                .BackColor = vbYellow
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
    
                .Text = "Processo numero " & fnNotNull(rs!AnnoProcesso) & "-" & fnNotNull(rs!NumeroProcesso)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo)
            End With
            
            Load Me.txtQuantitaLavorata(IControl)
            With Me.txtQuantitaLavorata(IControl)
                .Left = 10080
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!Quantita), 2)
                .BackColor = vbYellow
                .ForeColor = vbBlack
                TotaleProcesso = TotaleProcesso + fnNotNullN(rs!Quantita)
            End With
            DoEvents
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
    Exit Sub
End If

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POIVGamma")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento

Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "L"
                .BackColor = vbYellow
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!Oggetto)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbYellow
                .ForeColor = vbBlack
                .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
            End With
            
            Load Me.txtQuantitaLavorata(IControl)
            With Me.txtQuantitaLavorata(IControl)
                .Left = 10080
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!RV_POQuantitaMovimentata), 2)
                .BackColor = vbYellow
                .ForeColor = vbBlack
                TotaleProcesso = TotaleProcesso + fnNotNullN(rs!RV_POQuantitaMovimentata)
            End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing

End Sub







Private Sub CaricaVenditaDDT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimetata As Double


If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    'DOCUMENTO DI TRASPORTO
    sSQL = "SELECT ValoriOggettoPerTipo0002.IDOggetto, ValoriOggettoPerTipo0002.IDTipoOggetto, ValoriOggettoPerTipo0002.Doc_data, "
    sSQL = sSQL & "ValoriOggettoPerTipo0002.Doc_numero, ValoriOggettoPerTipo0002.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo0002.Nom_nome,"
    sSQL = sSQL & "ValoriOggettoDettaglio0004.Art_codice , ValoriOggettoDettaglio0004.Art_descrizione, ValoriOggettoDettaglio0004.Art_quantita_totale, "
    sSQL = sSQL & "Art_numero_colli, Art_peso, Art_tara, Art_quantita_pezzi "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & Link_RigaConferimento
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        Select Case Link_UnitaDiMisura_Coop_Conferimento
            Case 1
                QuantitaMovimentata = fnNotNullN(rs!Art_numero_colli)
            Case 2
                QuantitaMovimentata = fnNotNullN(rs!Art_peso)
            Case 3
                QuantitaMovimentata = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
            Case 4
                QuantitaMovimentata = fnNotNullN(rs!Art_tara)
            Case 5
                QuantitaMovimentata = fnNotNullN(rs!Art_quantita_pezzi)
        End Select
        
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "V"
                .BackColor = vbGreen
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbGreen
                .ForeColor = vbBlack
                .Text = "D.D.T. N° " & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_Data) & " al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbGreen
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!Art_codice) & " - " & fnNotNull(rs!Art_descrizione)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(QuantitaMovimentata, 2)
                .BackColor = vbGreen
                .ForeColor = vbBlack
                TotaleVendita = TotaleVendita + QuantitaMovimentata
            End With
            DoEvents
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
Exit Sub
End If

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "V"
                .BackColor = vbGreen
                .ForeColor = vbBlack
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbGreen
                .ForeColor = vbBlack
                .Text = fnNotNull(rs!Oggetto) & " " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbGreen
                .ForeColor = vbBlack
                .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!RV_POQuantitaMovimentata), 2)
                .BackColor = vbGreen
                .ForeColor = vbBlack
                TotaleVendita = TotaleVendita + fnNotNullN(rs!RV_POQuantitaMovimentata)
            End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CaricaVenditaFA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT ValoriOggettoPerTipo0072.IDOggetto, ValoriOggettoPerTipo0072.IDTipoOggetto, ValoriOggettoPerTipo0072.Doc_data, "
sSQL = sSQL & "ValoriOggettoPerTipo0072.Doc_numero, ValoriOggettoPerTipo0072.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo0072.Nom_nome,"
sSQL = sSQL & "ValoriOggettoDettaglio0001.Art_codice , ValoriOggettoDettaglio0001.Art_descrizione, ValoriOggettoDettaglio0001.Art_quantita_totale, "
sSQL = sSQL & "Art_numero_colli, Art_peso, Art_tara, Art_quantita_pezzi "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
        Select Case Link_UnitaDiMisura_Coop_Conferimento
            Case 1
                QuantitaMovimentata = fnNotNullN(rs!Art_numero_colli)
            Case 2
                QuantitaMovimentata = fnNotNullN(rs!Art_peso)
            Case 3
                QuantitaMovimentata = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
            Case 4
                QuantitaMovimentata = fnNotNullN(rs!Art_tara)
            Case 5
                QuantitaMovimentata = fnNotNullN(rs!Art_quantita_pezzi)
        End Select

        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


        Load Me.txtCausale(IControl)
        With Me.txtCausale(IControl)
            .Left = 120
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = "V"
            .BackColor = vbGreen
            .ForeColor = vbBlack
        End With
        
        Load Me.txtDescrizioneDocumento(IControl)
        With Me.txtDescrizioneDocumento(IControl)
            .Left = 960
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbGreen
            .ForeColor = vbBlack
            .Text = "Fattura accompagnatoria N° " & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_Data) & " al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome)
        End With
        
        Load Me.txtCodiceArticolo(IControl)
        With Me.txtCodiceArticolo(IControl)
            .Left = 5520
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbGreen
            .ForeColor = vbBlack
            .Text = fnNotNull(rs!Art_codice) & " - " & fnNotNull(rs!Art_descrizione)
        End With
        
        
        Load Me.txtQuantitaVendita(IControl)
        With Me.txtQuantitaVendita(IControl)
            .Left = 11400
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = FormatNumber(QuantitaMovimentata, 2)
            .BackColor = vbGreen
            .ForeColor = vbBlack
            TotaleVendita = TotaleVendita + QuantitaMovimentata
        End With
        DoEvents
    rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CaricaVenditaSNF()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'SCONTRINO NON FISCALE
sSQL = "SELECT ValoriOggettoPerTipo0008.IDOggetto, ValoriOggettoPerTipo0008.IDTipoOggetto, ValoriOggettoPerTipo0008.Doc_data, "
sSQL = sSQL & "ValoriOggettoPerTipo0008.Doc_numero, ValoriOggettoPerTipo0008.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo0008.Nom_nome,"
sSQL = sSQL & "ValoriOggettoDettaglio0034.Art_codice , ValoriOggettoDettaglio0034.Art_descrizione, ValoriOggettoDettaglio0034.Art_quantita_totale, "
sSQL = sSQL & "Art_numero_colli, Art_peso, Art_tara, Art_quantita_pezzi "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
        Select Case Link_UnitaDiMisura_Coop_Conferimento
            Case 1
                QuantitaMovimentata = fnNotNullN(rs!Art_numero_colli)
            Case 2
                QuantitaMovimentata = fnNotNullN(rs!Art_peso)
            Case 3
                QuantitaMovimentata = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
            Case 4
                QuantitaMovimentata = fnNotNullN(rs!Art_tara)
            Case 5
                QuantitaMovimentata = fnNotNullN(rs!Art_quantita_pezzi)
        End Select

        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop


        Load Me.txtCausale(IControl)
        With Me.txtCausale(IControl)
            .Left = 120
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = "V"
            .BackColor = vbGreen
            .ForeColor = vbBlack
        End With
        
        Load Me.txtDescrizioneDocumento(IControl)
        With Me.txtDescrizioneDocumento(IControl)
            .Left = 960
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbGreen
            .ForeColor = vbBlack
            .Text = "S.N.F. N° " & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_Data) & " al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome)
        End With
        
        Load Me.txtCodiceArticolo(IControl)
        With Me.txtCodiceArticolo(IControl)
            .Left = 5520
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .BackColor = vbGreen
            .ForeColor = vbBlack
            .Text = fnNotNull(rs!Art_codice) & " - " & fnNotNull(rs!Art_descrizione)
        End With
        
        Load Me.txtQuantitaVendita(IControl)
        With Me.txtQuantitaVendita(IControl)
            .Left = 11400
            .Top = TotaleTop
            .Visible = True
            .ZOrder 0
            .Text = FormatNumber(QuantitaMovimentata, 2)
            .BackColor = vbGreen
            .ForeColor = vbBlack
            TotaleVendita = TotaleVendita + QuantitaMovimentata
        End With
        DoEvents
    rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub



Private Sub CaricaFatturaImmediata()

End Sub

Private Sub CaricaNotaDiCredito()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    'Nota di credito
    sSQL = "SELECT ValoriOggettoPerTipo000B.IDOggetto, ValoriOggettoPerTipo000B.IDTipoOggetto, ValoriOggettoPerTipo000B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoPerTipo000B.Doc_numero, ValoriOggettoPerTipo000B.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo000B.Nom_nome,"
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Art_codice , ValoriOggettoDettaglio0016.Art_descrizione, ValoriOggettoDettaglio0016.Art_quantita_totale, "
    sSQL = sSQL & "ValoriOggettoDettaglio0016.Art_peso "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & Link_RigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "NC"
                .BackColor = vbBlue
                .ForeColor = vbYellow
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = "Nota credito N° " & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_Data) & " al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = fnNotNull(rs!Art_codice) & " - " & fnNotNull(rs!Art_descrizione)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!art_quantita_totale), 2)
                .BackColor = vbBlue
                .ForeColor = vbYellow
                TotaleNotaCredito = TotaleNotaCredito + fnNotNullN(rs!Art_peso)
            End With
            DoEvents
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
Exit Sub
End If

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=11 "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop
            
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "N.C."
                .BackColor = vbBlue
                .ForeColor = vbYellow
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = fnNotNull(rs!Oggetto) & " " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!RV_POQuantitaMovimentata), 2)
                .BackColor = vbBlue
                .ForeColor = vbYellow
                TotaleNotaCredito = TotaleNotaCredito + fnNotNullN(rs!RV_POQuantitaMovimentata)
            End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CaricaNotaDiDebito()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



If ATTIVAZIONE_NUOVO_CALCOLO = False Then

    'Nota di credito
    sSQL = "SELECT ValoriOggettoPerTipo006B.IDOggetto, ValoriOggettoPerTipo006B.IDTipoOggetto, ValoriOggettoPerTipo006B.Doc_data, "
    sSQL = sSQL & "ValoriOggettoPerTipo006B.Doc_numero, ValoriOggettoPerTipo006B.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo006B.Nom_nome,"
    sSQL = sSQL & "ValoriOggettoDettaglio0007.Art_codice , ValoriOggettoDettaglio0007.Art_descrizione, ValoriOggettoDettaglio0007.Art_quantita_totale "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0007.RV_POIDConferimentoRighe=" & Link_RigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
            IControl = IControl + 1
            TotaleTop = TotaleTop + VarTop
    
    
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "ND"
                .BackColor = vbBlue
                .ForeColor = vbYellow
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = "Nota debito N° " & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_Data) & " al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome)
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = fnNotNull(rs!Art_codice) & " - " & fnNotNull(rs!Art_descrizione)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!art_quantita_totale), 2)
                .BackColor = vbBlue
                .ForeColor = vbYellow
                TotaleDebito = TotaleDebito + fnNotNullN(rs!art_quantita_totale)
            End With
            DoEvents
        rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
Exit Sub
End If

''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=107 "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        IControl = IControl + 1
        TotaleTop = TotaleTop + VarTop
            
            Load Me.txtCausale(IControl)
            With Me.txtCausale(IControl)
                .Left = 120
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = "N.C."
                .BackColor = vbBlue
                .ForeColor = vbYellow
            End With
            
            Load Me.txtDescrizioneDocumento(IControl)
            With Me.txtDescrizioneDocumento(IControl)
                .Left = 960
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = fnNotNull(rs!Oggetto) & " " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            End With
            
            Load Me.txtCodiceArticolo(IControl)
            With Me.txtCodiceArticolo(IControl)
                .Left = 5520
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .BackColor = vbBlue
                .ForeColor = vbYellow
                .Text = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo)) & " - " & fnNotNull(rs!DescrizioneArticolo)
            End With
            
            Load Me.txtQuantitaVendita(IControl)
            With Me.txtQuantitaVendita(IControl)
                .Left = 11400
                .Top = TotaleTop
                .Visible = True
                .ZOrder 0
                .Text = FormatNumber(fnNotNullN(rs!RV_POQuantitaMovimentata), 2)
                .BackColor = vbBlue
                .ForeColor = vbYellow
                TotaleDebito = TotaleCredito + fnNotNullN(rs!RV_POQuantitaMovimentata)
            End With
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function RecuperaSegnoPerDisponibilita(IDFunzione) As String

End Function

Private Function GET_CODICE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceArticolo FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_ARTICOLO = ""
Else
    GET_CODICE_ARTICOLO = fnNotNull(rs!CodiceArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_PRODOTTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PRODOTTO = 0
Else
    GET_TIPO_PRODOTTO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function Get_Anagrafica(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Anagrafica, Nome "
sSQL = sSQL & "FROM IERepCliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Get_Anagrafica = ""
Else
    Get_Anagrafica = fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing
End Function

