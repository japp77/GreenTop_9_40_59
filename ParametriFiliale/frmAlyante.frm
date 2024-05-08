VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmAlyante 
   Caption         =   "PARAMETRI ALYANTE"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlyante.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Registrazione fatture"
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
      Height          =   1935
      Left            =   4680
      TabIndex        =   26
      Top             =   4200
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   4215
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Preleva mappature da Registrazione fatture"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Nome database"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   4215
      End
   End
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
      Height          =   495
      Left            =   6600
      TabIndex        =   21
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parametri azienda Alyante"
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
      Height          =   4095
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Disattiva calcolo N.D. non esportate "
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   4215
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Disattiva calcolo N.C. non esportate "
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   4215
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Disattiva calcolo F.D. non esportate "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   4215
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Disattiva calcolo F.A. non esportate "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Disattiva calcolo D.D.T. non fatturati"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Attiva gestione rischio da Alyante"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   4215
      End
      Begin DMTEDITNUMLib.dmtNumber txtCompanyCode 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   600
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Company code"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   660
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di connessione"
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
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox txtPwdUtenteAly 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   4920
         Width           =   4215
      End
      Begin VB.TextBox txtUtenteAly 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   4215
      End
      Begin VB.TextBox txtGruppoAly 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox txtUtenteIstanza 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox txtNomeDB 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtNomeServer 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Password utente Alyante"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   4680
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Utente Alyante"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Gruppo Alyante"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Utente proprietario istanza"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Nome database"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Nome server"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAlyante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()

    If (VerificaDati = False) Then Exit Sub
    
    SET_DATI
    
    CONFERMA_PARAMETRI_ALY = True
    
    Unload Me
    
End Sub

Private Sub cmdTest_Click()
    GET_CONNESSIONE_ALYANTE True
End Sub

Private Sub Form_Load()
    CONFERMA_PARAMETRI_ALY = False
    RECUPERA_DATI
End Sub
Private Sub RECUPERA_DATI()
    
    Me.txtNomeServer.Text = NOME_SERVER_ALY
    Me.txtNomeDB.Text = NOME_DB_ALY
    Me.txtUtenteIstanza.Text = USER_PROP_SERVER
    Me.txtPassword.Text = PWD_USER_PROP
    Me.txtGruppoAly.Text = GRUPPO_ALY
    Me.txtUtenteAly.Text = USER_ALY
    Me.txtPwdUtenteAly.Text = PWD_USER_ALY
    Me.txtCompanyCode.Value = COMPANY_CODE
    Me.Check1.Value = ATTIVA_FIDO_ALY
    Me.Check2.Value = DISATTIVA_DDT_FIDO_ALY
    Me.Check3.Value = DISATTIVA_FA_FIDO_ALY
    Me.Check4.Value = DISATTIVA_FD_FIDO_ALY
    Me.Check5.Value = DISATTIVA_NC_FIDO_ALY
    Me.Check6.Value = DISATTIVA_ND_FIDO_ALY
    Me.Text1.Text = DBNameRegFatture
    Me.Check7.Value = AttivaMappaturaDaRegFatture
End Sub
Private Sub SET_DATI()
    
    NOME_SERVER_ALY = txtNomeServer.Text
    NOME_DB_ALY = Me.txtNomeDB.Text
    USER_PROP_SERVER = Me.txtUtenteIstanza.Text
    PWD_USER_PROP = Me.txtPassword.Text
    GRUPPO_ALY = Me.txtGruppoAly.Text
    USER_ALY = Me.txtUtenteAly.Text
    PWD_USER_ALY = Me.txtPwdUtenteAly.Text
    COMPANY_CODE = Me.txtCompanyCode.Value
    ATTIVA_FIDO_ALY = Me.Check1.Value
    DISATTIVA_DDT_FIDO_ALY = Me.Check2.Value
    DISATTIVA_FA_FIDO_ALY = Me.Check3.Value
    DISATTIVA_FD_FIDO_ALY = Me.Check4.Value
    DISATTIVA_NC_FIDO_ALY = Me.Check5.Value
    DISATTIVA_ND_FIDO_ALY = Me.Check6.Value
    DBNameRegFatture = Me.Text1.Text
    AttivaMappaturaDaRegFatture = Me.Check7.Value
    
End Sub
Private Function VerificaDati() As Boolean
VerificaDati = True

If txtCompanyCode.Value <= 0 Then
    MsgBox "Inserire il 'Company code'", vbInformation, "Verifica dati"
    VerificaDati = False
    Exit Function
End If
If GET_CONNESSIONE_ALYANTE(False) = False Then
    MsgBox "Errore di inizializzazione della connessione al DB di Alyante", vbInformation, "Verifica dati"
    VerificaDati = False
    Exit Function
End If

End Function
Private Function GET_CONNESSIONE_ALYANTE(ConMessaggio As Boolean) As Boolean
On Error GoTo ERR_cmdTest_Click

GET_CONNESSIONE_ALYANTE = True

Dim i_objNucleo  As IEBO_NUCLEO.CLSIE_NUCLEO
Set i_objNucleo = New IEBO_NUCLEO.CLSIE_NUCLEO

i_objNucleo.NomeServer = txtNomeServer.Text
i_objNucleo.NomeDB = Me.txtNomeDB.Text
i_objNucleo.UserIdDB = Me.txtUtenteIstanza.Text
i_objNucleo.PasswordDB = Me.txtPassword.Text
i_objNucleo.GruppoUtenti = Me.txtGruppoAly.Text
i_objNucleo.UtenteCorrente = Me.txtUtenteAly.Text
i_objNucleo.PwdUtente = Me.txtPwdUtenteAly.Text
i_objNucleo.CodiceDitta = Me.txtCompanyCode.Value

If Not i_objNucleo.Inizializza() Then
    If ConMessaggio = True Then
        MsgBox "Errore di inizializzazione!", vbCritical, "Verifica connessione"
    End If
    GET_CONNESSIONE_ALYANTE = False
Else
    If ConMessaggio = True Then
        MsgBox "Connessione avvenuta con successo!", vbInformation, "Verifica connessione"
    End If
End If

If Not i_objNucleo Is Nothing Then
   i_objNucleo.Terminate
   Set i_objNucleo.adoConnection = Nothing
End If
Set i_objNucleo = Nothing

Exit Function
ERR_cmdTest_Click:
    GET_CONNESSIONE_ALYANTE = False
    MsgBox Err.Description, vbCritical, "Verifica test di connessione"
End Function

