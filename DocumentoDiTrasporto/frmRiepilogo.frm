VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRiepilogo 
   Caption         =   "Riepilogo riga di conferimento"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiepilogo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   20310
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   11175
      Left            =   0
      ScaleHeight     =   11175
      ScaleWidth      =   20295
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      Begin VB.Frame FraMargine 
         Caption         =   "Margine"
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
         Height          =   2775
         Left            =   0
         TabIndex        =   7
         Top             =   7320
         Width           =   4455
         Begin VB.TextBox txtTotaleConferimentoMarg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtMargineImporto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtPercentualeMargine 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleVenduto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label11 
            Caption         =   "Totale conferito"
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
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblTotaleVenduto 
            Caption         =   "Totale venduto"
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
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblMargine 
            Caption         =   "Margine"
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
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblPercMargine 
            Caption         =   "% su margine"
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
            Left            =   120
            TabIndex        =   25
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame FraParametri 
         Caption         =   "Parametri"
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
         Height          =   1815
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4455
         Begin VB.CheckBox chkVisNotaDebito 
            Caption         =   "Visualizza note di debito"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkVisNotaCredito 
            Caption         =   "Visualizza note di credito"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1140
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkVisVend 
            Caption         =   "Visualizza vendita"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   860
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkVisLav 
            Caption         =   "Visualizza lavorazioni"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   540
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.CheckBox chkVisQuad 
            Caption         =   "Visualizza quadratura"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Checked
            Width           =   4215
         End
      End
      Begin VB.Frame FraRiepConfVend 
         Caption         =   "Riepilogo Conferimento/Vendita"
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
         Height          =   2775
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox txtDifferenza 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleQuantitaVendita 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleQuantitaQuadrata 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleConferimento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label6 
            Caption         =   "DIFFERENZA"
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
            Left            =   120
            TabIndex        =   15
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "VENDITA"
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
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "QUADRATURA"
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
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "CONFERIMENTO"
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
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame FraRiepConfLav 
         Caption         =   "Riepilogo Conferimento/Lavorazione"
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
         Height          =   2775
         Left            =   0
         TabIndex        =   4
         Top             =   4560
         Width           =   4455
         Begin VB.TextBox txtDifferenzaLavorazione 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleQuantitaLavorata 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleQuantitaQuadrataLav 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtTotaleConferimentoLav 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label10 
            Caption         =   "DIFFERENZA"
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
            Left            =   120
            TabIndex        =   21
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "LAVORAZIONE"
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
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "QUADRATURA"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "CONFERIMENTO"
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
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
      End
      Begin DmtGridCtl.DmtGrid Griglia 
         Height          =   10935
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   19288
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMove      =   0   'False
         ColumnsHeaderHeight=   20
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
Private Link_UnitaDiMisura_Coop_Conferimento As Long


Private TotaleImportoVendita As Double
Private TotaleImportoConf As Double

Private rsGriglia As ADODB.Recordset

Private Sub CREA_RECORDSET()
    
    Set rsGriglia = New ADODB.Recordset
    
    rsGriglia.CursorLocation = adUseClient
    
    rsGriglia.Fields.Append "IDConferimento", adInteger, , adFldIsNullable
    rsGriglia.Fields.Append "IDConferimentoRiga", adInteger, , adFldIsNullable
    rsGriglia.Fields.Append "IDLavorazione", adInteger, , adFldIsNullable
    rsGriglia.Fields.Append "Causale", adVarChar, 50, adFldIsNullable
    rsGriglia.Fields.Append "Documento", adVarChar, 250, adFldIsNullable
    rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
    rsGriglia.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
    rsGriglia.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
    rsGriglia.Fields.Append "QuantitaConferita", adDouble, , adFldIsNullable
    rsGriglia.Fields.Append "QuantitaQuadrata", adDouble, , adFldIsNullable
    rsGriglia.Fields.Append "QuantitaLavorata", adDouble, , adFldIsNullable
    rsGriglia.Fields.Append "QuantitaVenduta", adDouble, , adFldIsNullable
    rsGriglia.Fields.Append "CollegamentoLavVend", adVarChar, 250, adFldIsNullable
    rsGriglia.Fields.Append "Importo", adDouble, , adFldIsNullable
    
    
    rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
    

End Sub

Private Sub chkVisLav_Click()
    GET_GRIGLIA
End Sub

Private Sub chkVisNotaCredito_Click()
    GET_GRIGLIA
End Sub

Private Sub chkVisNotaDebito_Click()
    GET_GRIGLIA
End Sub

Private Sub chkVisQuad_Click()
    GET_GRIGLIA
End Sub

Private Sub chkVisVend_Click()
    GET_GRIGLIA
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
    
    
    CREA_RECORDSET
    
    Configurazione
    
    
    
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
On Error GoTo ERR_Form_Resize
    If Me.WindowState <> 1 Then
    
        'Me.Pic1.Height = Me.Height
        'Me.Pic1.Width = Me.Width
        
        If Me.Width >= 800 Then
            Me.Pic1.Width = Me.Width - 120
             
            Me.Griglia.Width = Me.Pic1.Width - Me.FraMargine.Width - 240
        End If
        
        If Me.Height >= 600 Then
            Me.Pic1.Height = Me.Height - 480
            
            Me.Griglia.Height = Me.Pic1.Height - 240
            
        End If
        
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

Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"
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
On Error GoTo ERR_Configurazione
    IControl = 0
    TotaleConferimento = 0
    TotaleQuadratura = 0
    TotaleVendita = 0
    TotaleDebito = 0
    TotaleNotaCredito = 0
    TotaleAssegnazione = 0
    TotaleProcesso = 0
    TotaleTop = 0
    TotaleImportoVendita = 0
    TotaleImportoConf = 0
    
    If (VISUALIZZA_IMPORTO_F4 = 1) Then
        Me.FraMargine.Visible = True
    Else
        Me.FraMargine.Visible = False
    End If
    
    CaricaConferimento
    
    DoEvents
    
    CaricaLavorazione
    CaricaAssegnazione
    CaricaProcesso
    CaricaVenditaDDT
    
    
    CaricaNotaDiCredito
    
    CaricaNotaDiDebito
    
    
    GET_GRIGLIA
    
    If ATTIVAZIONE_NUOVO_CALCOLO = False Then
        CaricaVenditaFA
        CaricaVenditaSNF
    End If
    

    Me.txtTotaleConferimento.Text = FormatNumber(TotaleConferimento, 2)
    Me.txtTotaleConferimentoLav.Text = FormatNumber(TotaleConferimento, 2)
    
    Me.txtTotaleQuantitaLavorata.Text = FormatNumber((TotaleAssegnazione + TotaleProcesso), 2)
    Me.txtTotaleQuantitaQuadrata.Text = FormatNumber(TotaleQuadratura, 2)
    Me.txtTotaleQuantitaQuadrataLav.Text = FormatNumber(TotaleQuadratura, 2)
    
    Me.txtTotaleQuantitaVendita.Text = FormatNumber(TotaleVendita, 2)
    
    Me.txtMargineImporto.Visible = False
    Me.txtPercentualeMargine.Visible = False
    Me.txtTotaleVenduto.Visible = False
    Me.lblTotaleVenduto.Visible = False
    Me.lblMargine.Visible = False
    Me.lblPercMargine.Visible = False
    
    
    
    If (VISUALIZZA_IMPORTO_F4 = 1) Then
        Me.txtMargineImporto.Visible = True
        Me.txtPercentualeMargine.Visible = True
        Me.txtTotaleVenduto.Visible = True
        Me.lblTotaleVenduto.Visible = True
        Me.lblMargine.Visible = True
        Me.lblPercMargine.Visible = True
        
        Me.txtTotaleConferimentoMarg.Text = FormatNumber(TotaleImportoConf, 2)
        Me.txtTotaleVenduto.Text = FormatNumber(TotaleImportoVendita, 2)
        Me.txtMargineImporto.Text = FormatNumber((TotaleImportoVendita - TotaleImportoConf), 2)
        Me.txtPercentualeMargine.Text = ""
        If TotaleImportoConf > 0 Then
            Me.txtPercentualeMargine.Text = FormatNumber((((TotaleImportoVendita / TotaleImportoConf) * 100) - 100), 2)
        End If
        
    End If

    Me.txtDifferenzaLavorazione.Text = FormatNumber((TotaleConferimento - (TotaleQuadratura + TotaleAssegnazione + TotaleProcesso)), 2)
    
    If Me.txtDifferenzaLavorazione.Text <> 0 Then
        Me.txtDifferenzaLavorazione.BackColor = vbRed
        Me.txtDifferenzaLavorazione.ForeColor = vbBlack
        
    Else
        Me.txtDifferenzaLavorazione.BackColor = vbGreen
        Me.txtDifferenzaLavorazione.ForeColor = vbBlack
    End If
    

    Me.txtDifferenza.Text = FormatNumber((TotaleConferimento - (TotaleQuadratura + TotaleVendita)), 2)
    If Me.txtDifferenza.Text <> 0 Then
        Me.txtDifferenza.BackColor = vbRed
        Me.txtDifferenza.ForeColor = vbBlack
        
    Else
        Me.txtDifferenza.BackColor = vbGreen
        Me.txtDifferenza.ForeColor = vbBlack
    End If
    

    
Exit Sub
ERR_Configurazione:
    MsgBox Err.Description, vbCritical, "Configurazione"
End Sub
Private Sub CaricaConferimento()
On Error GoTo ERR_CaricaConferimento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POCaricoMerceTesta.NumeroDocumento, RV_POCaricoMerceTesta.DataDocumento, RV_POCaricoMerceTesta.Anagrafica, RV_POCaricoMerceRighe.IDUnitaDiMisura, "
sSQL = sSQL & "RV_POCaricoMerceTesta.Nome, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceRighe.Qta_UM, RV_POCaricoMerceRighe.TotaleImponibileRiga, RV_POCaricoMerceRighe.IDArticolo "
sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not (rs.EOF) Then

    Link_UnitaDiMisura_Coop_Conferimento = fnNotNullN(rs!IDUnitaDiMisura)
            
    rsGriglia.AddNew
        rsGriglia!IDConferimentoRiga = Link_RigaConferimento
        rsGriglia!Causale = "C"
        rsGriglia!Documento = "N° " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " del socio " & fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
        rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsGriglia!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
        rsGriglia!Articolo = fnNotNull(rs!Articolo)
        rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
        'rsGriglia!QuantitaQuadrata = 0
        'rsGriglia!QuantitaLavorata = 0
        'rsGriglia!QuantitaVenduta = 0
        'rsGriglia!CollegamentoLavVend = ""
        rsGriglia!Importo = fnNotNullN(rs!TotaleImponibileRiga)
    
    rsGriglia.Update
                    
                    
    TotaleConferimento = TotaleConferimento + fnNotNullN(rs!Qta_UM)
    TotaleImportoConf = TotaleImportoConf + fnNotNullN(rs!TotaleImponibileRiga)
    
End If
rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_CaricaConferimento:
    MsgBox Err.Description, vbCritical, "CaricaConferimento"

End Sub

Private Sub CaricaLavorazione()
On Error GoTo ERR_CaricaLavorazione
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double


sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POLavorazioneL")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    rsGriglia.AddNew
        rsGriglia!IDConferimentoRiga = Link_RigaConferimento
        rsGriglia!Causale = "S"
        rsGriglia!Documento = fnNotNull(rs!Oggetto)
        rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
        rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
        'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
        
        'rsGriglia!QuantitaLavorata = 0
        'rsGriglia!QuantitaVenduta = 0
        'rsGriglia!CollegamentoLavVend = ""
        'rsGriglia!Importo = FormatNumber(fnNotNullN(rs!TotaleImponibileRiga), 2)
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
        
        rsGriglia!QuantitaQuadrata = QuantitaMovimentata
        
        
    rsGriglia.Update
    
    
    TotaleQuadratura = TotaleQuadratura + QuantitaMovimentata
    
    DoEvents
rs.MoveNext
Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CaricaLavorazione:
    MsgBox Err.Description, vbCritical, "CaricaLavorazione"
    
End Sub
Private Sub CaricaAssegnazione()
On Error GoTo ERR_CaricaAssegnazione
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double
Dim IDUMCoop As Long
Dim MoltiplicatoreArticolo As Double


''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        rsGriglia.AddNew
            rsGriglia!IDConferimentoRiga = Link_RigaConferimento
            rsGriglia!Causale = "L"
            rsGriglia!Documento = fnNotNull(rs!Oggetto)
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
            rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
            'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
            
            'rsGriglia!QuantitaLavorata = 0
            'rsGriglia!QuantitaVenduta = 0
            rsGriglia!CollegamentoLavVend = GET_VENDITA_LAVORAZIONE(fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma))
            
            'rsGriglia!Importo = FormatNumber(fnNotNullN(rs!TotaleImponibileRiga), 2)
            rsGriglia!QuantitaLavorata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            IDUMCoop = GET_UM_COOP_CONFERIMENTO(Link_RigaConferimento)
'            If ((IDUMCoop = 1) Or (IDUMCoop = 5)) Then
'                MoltiplicatoreArticolo = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!IDArticolo))
'                rsGriglia!QuantitaLavorata = rsGriglia!QuantitaLavorata * MoltiplicatoreArticolo
'            End If
            
            TotaleAssegnazione = TotaleAssegnazione + fnNotNullN(rsGriglia!QuantitaLavorata)
            
        rsGriglia.Update
        
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CaricaAssegnazione:
    MsgBox Err.Description, vbCritical, "CaricaAssegnazione"
    
End Sub

Private Sub CaricaProcesso()
On Error GoTo ERR_CaricaProcesso
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimentata As Double


''''''NUOVO CALCOLO DAI MOVIMENTI
sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto("RV_POIVGamma")
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento

Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        rsGriglia.AddNew
            rsGriglia!IDConferimentoRiga = Link_RigaConferimento
            rsGriglia!Causale = "S"
            rsGriglia!Documento = fnNotNull(rs!Oggetto)
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
            rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
            'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
            
            'rsGriglia!QuantitaLavorata = 0
            'rsGriglia!QuantitaVenduta = 0
            'rsGriglia!CollegamentoLavVend = ""
            'rsGriglia!Importo = FormatNumber(fnNotNullN(rs!TotaleImponibileRiga), 2)
            rsGriglia!CollegamentoLavVend = GET_VENDITA_LAVORAZIONE(fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma))
            rsGriglia!QuantitaLavorata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            
        rsGriglia.Update

        TotaleProcesso = TotaleProcesso + fnNotNullN(rs!RV_POQuantitaMovimentata)
        
        DoEvents
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CaricaProcesso:
    MsgBox Err.Description, vbCritical, "CaricaProcesso"
    
End Sub
Private Sub CaricaVenditaDDT()
On Error GoTo ERR_CaricaVenditaDDT
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaMovimetata As Double
Dim IDUMCoop As Long
Dim MoltiplicatoreArticolo As Double
Dim QuantitaVenduta As Double
Dim ImportoVenduto As Double

sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)

    While Not rs.EOF
        
        rsGriglia.AddNew
            rsGriglia!IDConferimentoRiga = Link_RigaConferimento
            rsGriglia!Causale = "V"
            rsGriglia!Documento = fnNotNull(rs!Oggetto) & " " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
            rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
            'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
            
            'rsGriglia!QuantitaLavorata = 0
            'rsGriglia!QuantitaVenduta = 0
            'rsGriglia!CollegamentoLavVend = ""
            rsGriglia!Importo = fnNotNullN(rs!Importo)
            rsGriglia!QuantitaVenduta = fnNotNullN(rs!RV_POQuantitaMovimentata)
            IDUMCoop = GET_UM_COOP_CONFERIMENTO(Link_RigaConferimento)
'            If ((IDUMCoop = 1) Or (IDUMCoop = 5)) Then
'                MoltiplicatoreArticolo = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!IDArticolo))
'                rsGriglia!QuantitaVenduta = rsGriglia!QuantitaVenduta * MoltiplicatoreArticolo
'            End If
            QuantitaVenduta = rsGriglia!QuantitaVenduta
        rsGriglia.Update
        
        TotaleVendita = TotaleVendita + QuantitaVenduta 'fnNotNullN(rs!RV_POQuantitaMovimentata)
        TotaleImportoVendita = TotaleImportoVendita + fnNotNullN(rs!Importo)
        
        
        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_CaricaVenditaDDT:
    MsgBox Err.Description, vbCritical, "CaricaVenditaDDT"
    
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
        DoEvents
    rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub



Private Sub CaricaFatturaImmediata()

End Sub

Private Sub CaricaNotaDiCredito()
On Error GoTo ERR_CaricaNotaDiCredito
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=11 "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        
        rsGriglia.AddNew
            rsGriglia!IDConferimentoRiga = Link_RigaConferimento
            rsGriglia!Causale = "N.C."
            rsGriglia!Documento = fnNotNull(rs!Oggetto) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
            rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
            'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
            
            'rsGriglia!QuantitaLavorata = 0
            'rsGriglia!QuantitaVenduta = 0
            'rsGriglia!Importo = FormatNumber(fnNotNullN(rs!TotaleImponibileRiga), 2)
            rsGriglia!QuantitaLavorata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            
            TotaleNotaCredito = TotaleNotaCredito + fnNotNullN(rs!RV_POQuantitaMovimentata)
        rsGriglia.Update

        DoEvents
        
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CaricaNotaDiCredito:
    MsgBox Err.Description, vbCritical, "CaricaNotaDiCredito"
    
End Sub
Private Sub CaricaNotaDiDebito()
On Error GoTo ERR_CaricaNotaDiDebito
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=107 "
sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " ORDER BY DataDocumento"
Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
    
        rsGriglia.AddNew
            rsGriglia!IDConferimentoRiga = Link_RigaConferimento
            rsGriglia!Causale = "N.D."
            rsGriglia!Documento = fnNotNull(rs!Oggetto) & " al cliente " & Get_Anagrafica(fnNotNullN(rs!IDAnagrafica))
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))
            rsGriglia!Articolo = fnNotNull(rs!DescrizioneArticolo)
            'rsGriglia!QuantitaConferita = fnNotNullN(rs!Qta_UM)
            
            'rsGriglia!QuantitaLavorata = 0
            'rsGriglia!QuantitaVenduta = 0
            'rsGriglia!Importo = FormatNumber(fnNotNullN(rs!TotaleImponibileRiga), 2)
            rsGriglia!QuantitaLavorata = fnNotNullN(rs!RV_POQuantitaMovimentata)
            
            TotaleDebito = TotaleDebito + fnNotNullN(rs!RV_POQuantitaMovimentata)
        rsGriglia.Update

    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_CaricaNotaDiDebito:
    MsgBox Err.Description, vbCritical, "CaricaNotaDiDebito"
    
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
Private Function GET_VENDITA_LAVORAZIONE(IDAssegnazioneMerce As Long, IDProcessoIVGamma As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_VENDITA_LAVORAZIONE = ""

sSQL = "SELECT IDRV_POAssegnazioneMerce, IDOggettoOrdine, NumeroOrdine, DataOrdine, IDCliente, Nom_nome, Nom_ragione_sociale_o_cognome "
sSQL = sSQL & "FROM RV_PO_IEAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_VENDITA_LAVORAZIONE = "(Ordine n° " & fnNotNull(rs!NumeroOrdine) & " del " & fnNotNull(rs!DataOrdine) & " di " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & " " & fnNotNull(rs!Nom_nome) & ") "
End If

sSQL = "SELECT IDMovimento, IDTipoOggetto, Oggetto, NumeroDocumento, DataDocumento FROM Movimento "
sSQL = sSQL & "WHERE (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce=" & IDAssegnazioneMerce
sSQL = sSQL & " AND RV_POIDProcessoIVGamma=" & IDProcessoIVGamma
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_VENDITA_LAVORAZIONE = GET_VENDITA_LAVORAZIONE & " - " & fnNotNull(rs!Oggetto) & " N° " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento)
End If


End Function
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    sSQL = "Causale=" & fnNormString("C")
    
    If Me.chkVisQuad.Value = vbChecked Then
        sSQL = sSQL & " OR Causale=" & fnNormString("S")
    End If
    If Me.chkVisLav.Value = vbChecked Then
        sSQL = sSQL & " OR Causale=" & fnNormString("L")
    End If
    If Me.chkVisVend.Value = vbChecked Then
        sSQL = sSQL & " OR Causale=" & fnNormString("V")
    End If
    If Me.chkVisNotaDebito.Value = vbChecked Then
        sSQL = sSQL & " OR Causale=" & fnNormString("N.D.")
    End If
    If Me.chkVisNotaCredito.Value = vbChecked Then
        sSQL = sSQL & " OR Causale=" & fnNormString("N.C.")
    End If
    
    
    rsGriglia.Filter = sSQL
    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDConferimentoRiga", "IDConferimentoRiga", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "Causale", "Causale", dgchar, True, 500, dgAligncenter
                .ColumnsHeader.Add "Documento", "Documento", dgchar, True, 3500, dgAlignleft
                
                
                .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Articolo", "Articolo", dgchar, False, 2000, dgAlignleft
                
                Set cl = .ColumnsHeader.Add("QuantitaConferita", "Q.tà conf.", dgDouble, True, 1800, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("QuantitaQuadrata", "Q.tà quad.", dgDouble, True, 1800, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("QuantitaLavorata", "Q.tà lav.", dgDouble, True, 1800, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("QuantitaVenduta", "Q.tà vend.", dgDouble, True, 1800, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "CollegamentoLavVend", "Collegamento lavorazione alla vendita", dgchar, False, 3500, dgAlignleft
                
                If (VISUALIZZA_IMPORTO_F4 = 1) Then
                    Set cl = .ColumnsHeader.Add("Importo", "Importo", dgDouble, True, 1800, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                End If
                
            .EnableRowColors = True
            .RowColors.Clear
            .RowColors.Add "Conferimento", "Causale='C'", vbRed, vbWhite
            .RowColors.Add "Quadratura", "Causale='S'", vbBlue, vbWhite
            .RowColors.Add "Lavorazione", "Causale='L'", vbYellow, vbBlack
            .RowColors.Add "Vendita", "Causale='V'", vbGreen, vbBlack
            .RowColors.Add "NotaCredito", "Causale='N.C.'", vbCyan, vbBlack
            .RowColors.Add "NotaDebito", "Causale='N.D.'", vbCyan, vbBlack
            '.RowColors.Add "AltriInterventi", "NumeroFase>1", &HFFFFCC
                
        Set .Recordset = rsGriglia
        .LoadUserSettings
        .Refresh
        
    End With
    
    'Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Function GET_UM_COOP_CONFERIMENTO(IDRigaConferimento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_UM_COOP_CONFERIMENTO = 0

sSQL = "SELECT IDRV_POCaricoMerceRighe, IDUnitaDiMisura "
sSQL = sSQL & " FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_UM_COOP_CONFERIMENTO = fnNotNullN(rs!IDUnitaDiMisura)
End If
rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POMoltiplicatore FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_MOLTIPLICATORE_ARTICOLO = 1
Else
    If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
        GET_MOLTIPLICATORE_ARTICOLO = 1
    Else
        GET_MOLTIPLICATORE_ARTICOLO = fnNotNullN(rs!RV_POMoltiplicatore)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
