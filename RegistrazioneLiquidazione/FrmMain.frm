VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creazione liquidazione (1 di 4)"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo di liquidazione"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      Begin VB.OptionButton OptScelta 
         Caption         =   "Stampe di controllo"
         Enabled         =   0   'False
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
         TabIndex        =   6
         Top             =   2640
         Width           =   6135
      End
      Begin VB.OptionButton OptScelta 
         Caption         =   "Nuovo periodo di  liquidazione"
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
         TabIndex        =   5
         Top             =   1800
         Width           =   6135
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   0
      Picture         =   "FrmMain.frx":4781A
      ScaleHeight     =   4755
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
'Variabili recordset per le visualizzazioni delle griglie
Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    Set CnDMT = m_App.Database.Connection
    
    CnDMT.CursorLocation = adUseClient
    
    PrelevaAzienda
    
    VarPassword = m_App.Password
    VarUtente = m_App.User
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    GET_PARAMETRI_AZIENDA
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
    
End Sub
Private Sub GET_PARAMETRI_AZIENDA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    RISCONTRO_PESO_SOCIO_VAL = fnNotNullN(rs!RiscontroPesoAValorePerSocio)
    RISCONTRO_PESO_FORN_VAL = fnNotNullN(rs!RiscontroPesoAValorePerFornitore)
    LINK_TIPO_CATEGORIA_SOCIO = fnNotNullN(rs!IDCategoriaAnagrafica)
Else
    RISCONTRO_PESO_SOCIO_VAL = 0
    RISCONTRO_PESO_FORN_VAL = 0
    LINK_TIPO_CATEGORIA_SOCIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property
Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property
Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della liquidazione?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub cmdAvanti_Click()
    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Else
        cmdAvanti.Value = True
    End If
End Sub

Private Sub Form_Load()
    Me.OptScelta(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Me.cmdAnnulla.Value = True Then
            
    End If

    If Me.cmdAvanti.Value = True Then
        
        Select Case TipoScelta
            
            Case 1
                Nuova_Liquidazione = 1
            Case 2
                Nuova_Liquidazione = 0
        End Select
        
        FrmNuovoPeriodo.Show

        Exit Sub
        
    End If
End Sub

Private Sub OptScelta_Click(Index As Integer)
    Select Case Index
        
        Case 0
            TipoScelta = 1
        Case 1
            TipoScelta = 2
    End Select
End Sub
Private Sub GET_MODULO_ATTIVATO(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
Else
    MODULO_ATTIVATO = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
End Sub
