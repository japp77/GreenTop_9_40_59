VERSION 5.00
Object = "{3C64B308-0C04-11D2-B957-002018813989}#13.1#0"; "DMTCliFu.OCX"
Begin VB.Form frmFido 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calcolo del Fido Cliente"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   6120
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Immetti password"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   4935
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
   End
   Begin DMTCliFu.DmtFidoCliente DmtFido 
      Height          =   4545
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   8017
   End
   Begin VB.Label Label1 
      Caption         =   "Attenzione il Fido assegnato al Cliente è stato superato. Scegliere OK per continuare o Annulla per tornare indietro."
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   135
      Width           =   4965
   End
End
Attribute VB_Name = "frmFido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private mButtonType As DmtFido_ButtonType

Private Result As Integer


Public Function CheckFido() As Boolean
    
    Set DmtFido.Connection = Cn

    DmtFido.IDFirm = TheApp.IDFirm
    DmtFido.IDAnagrafica = frmMain.cdAnagrafica.KeyFieldID
    DmtFido.IDTipoAnagrafica = 2
    DmtFido.TotDocumento = frmMain.curNettoAPagare.Value
    DmtFido.IDPagamento = frmMain.cboPagamento.CurrentID
    DmtFido.ButtonType = 0 'mButtonType
    DmtFido.IDDocumento = oDoc.IDOggetto
    DmtFido.DocTableName = sTabellaTestata
    DmtFido.CheckFido
    
End Function


Private Sub cmdAnnulla_Click()
    AVVIA_FIDO_DOPO_CONTROLLO = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Select Case LINK_TIPO_FIDO_CLIENTE
    
        Case 1
            If Len(DATA_SBLOCCO_FIDO_CLIENTE) > 0 Then
                If DateDiff("d", Date, DATA_SBLOCCO_FIDO_CLIENTE) < 0 Then
                    MsgBox "Le impostazioni del cliente non permettono di salvare il documento", vbInformation, "Salvataggio documento"
                    Exit Sub
                Else
                    AVVIA_FIDO_DOPO_CONTROLLO = True
                    Unload Me
                End If
            Else
                AVVIA_FIDO_DOPO_CONTROLLO = False
                MsgBox "Le impostazioni del cliente non permettono di salvare il documento", vbInformation, "Salvataggio documento"
            End If
        Case 2
            If Me.txtPassword.Text = PASSWORD_FIDO_CLIENTE Then
                If Len(DATA_SBLOCCO_FIDO_CLIENTE) > 0 Then
                    If DateDiff("d", Date, DATA_SBLOCCO_FIDO_CLIENTE) < 0 Then
                        MsgBox "Le impostazioni del cliente non permettono di salvare il documento", vbInformation, "Salvataggio documento"
                        Exit Sub
                    Else
                        AVVIA_FIDO_DOPO_CONTROLLO = True
                        Unload Me
                    End If
                Else
                    AVVIA_FIDO_DOPO_CONTROLLO = True
                    Unload Me
                End If
            Else
                MsgBox "Password errata", vbCritical, "Salvataggio documento"
                Me.txtPassword.SetFocus
            End If
        Case 3
            AVVIA_FIDO_DOPO_CONTROLLO = True
            Unload Me
        Case 4
            Select Case LINK_TIPO_FIDO_AZIENDA
                Case 1
                    AVVIA_FIDO_DOPO_CONTROLLO = False
                    MsgBox "Le impostazioni dell'azienda non permettono di salvare il documento per questo cliente", vbInformation, "Salvataggio documento"
                    Unload Me
                Case 2
                    If Me.txtPassword.Text = PASSWORD_FIDO_AZIENDA Then
                        AVVIA_FIDO_DOPO_CONTROLLO = True
                        Unload Me
                    Else
                        MsgBox "Password errata", vbCritical, "Salvataggio documento"
                        Me.txtPassword.SetFocus
                    End If
                Case 3
                    AVVIA_FIDO_DOPO_CONTROLLO = True
                    Unload Me
                Case Else
                    AVVIA_FIDO_DOPO_CONTROLLO = True
                    Unload Me
            End Select
        Case Else
            AVVIA_FIDO_DOPO_CONTROLLO = True
            Unload Me
            
    End Select
    
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
CheckFido
If LINK_TIPO_FIDO_CLIENTE = 2 Then
    Me.Frame1.Visible = True
    Me.cmdOK.Top = 6120
    Me.cmdAnnulla.Top = 6120
    Me.Height = 7035
Else
    Me.Frame1.Visible = False
    Me.cmdOK.Top = 5520
    Me.cmdAnnulla.Top = 5520
    Me.Height = 6585
End If
If LINK_TIPO_FIDO_CLIENTE = 4 Then
    If LINK_TIPO_FIDO_AZIENDA = 2 Then
        Me.Frame1.Visible = True
        Me.cmdOK.Top = 6120
        Me.cmdAnnulla.Top = 6120
        Me.Height = 7035
    Else
        Me.Frame1.Visible = False
        Me.cmdOK.Top = 5520
        Me.cmdAnnulla.Top = 5520
        Me.Height = 6585
    End If
End If


End Sub
