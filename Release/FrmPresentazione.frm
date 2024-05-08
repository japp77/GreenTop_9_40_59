VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmPresentazione 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmPresentazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "CONTROLLO UTILIZZO DATI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   12
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Esegui script"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Controllo versioni client"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Elimina coda"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
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
   Begin VB.CommandButton cmdDocumentazione 
      Caption         =   "Documentazione"
      Height          =   495
      Left            =   6240
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.CommandButton cmdConfigMenuFiliale 
      Caption         =   "Configurazione menù filiale"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton cmdConfigParFiliale 
      Caption         =   "Configurazione parametri filiale"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdAvviaPoint 
      Caption         =   "AVVIA POINT UPDATE"
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
      Left            =   480
      Picture         =   "FrmPresentazione.frx":4781A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Avvia procedura di aggiornamento automatico"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":47DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":480BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":48658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":48BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":48D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4919E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":494B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":497D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":49C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":49F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4A390
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4A7E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4AAFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4AC56
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4BAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4BEFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4DC04
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4DF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4E238
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4E68A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4EADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4EDF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4EEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4F34E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":4F668
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5023A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5068C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":50ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":50F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":51382
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":517D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":57A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":57D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":580A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":583BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":586D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":58B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":590C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5965C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":59BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5A190
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5A72A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5ACC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":5B25E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Funzionalità abilitate"
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
      TabIndex        =   8
      Top             =   4560
      Width           =   8775
   End
   Begin VB.Label lblVersione 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data ultimo aggiornamento:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblDataInstallazione 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7680
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   2040
      Left            =   0
      Picture         =   "FrmPresentazione.frx":5B7F8
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "FrmPresentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsGriglia As ADODB.Recordset

Private Sub cmdAvviaPoint_Click()
On Error GoTo ERR_cmdAvviaPoint_Click

    DontShowAtStartup "Point"
    
    ShowAtStartup "Point"
    
    If CONTROLLO_PROCESSO_ATTIVO("Point.exe") = False Then
        Shell MenuOptions.ProgramsPath & "\Point.exe"
    End If

    MsgBox "PROCEDURA AVVIATA CON SUCCESSO", vbInformation, "GREENTOP"
    
Exit Sub
ERR_cmdAvviaPoint_Click:
    MsgBox Err.Description, vbCritical, "cmdAvviaPoint_Click"
    
End Sub

Private Sub cmdConfigMenuFiliale_Click()
    frmConfigMenu.Show vbModal
End Sub

Private Sub cmdConfigParFiliale_Click()
    frmConfigFiliale.Show vbModal
End Sub

Private Sub cmdDocumentazione_Click()
On Error GoTo ERR_cmdPubblica_Click
Dim RetVal
Dim NomeCartella As String
NomeCartella = MenuOptions.ProgramsPath & "\Help_" & GET_NOME_PROGRAMMA(IdentificativoProgramma)
RetVal = Shell("explorer" & " " & NomeCartella, 1)
Exit Sub
ERR_cmdPubblica_Click:
    MsgBox Err.Description, vbCritical, "cmdDocumentazione_Click"

End Sub



Private Sub cmdTrovaAzienda_Click()
    frmAzienda.Show vbModal
End Sub



Private Sub Command1_Click()
Dim sSQL As String


sSQL = "DELETE FROM RV_POTMP"
CnDMT.Execute sSQL

MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO"

Exit Sub

End Sub

Private Sub Command2_Click()
    frmControlloVersioni.Show vbModal
    
End Sub

Private Sub Command3_Click()
frmEseguiScript.Show vbModal

End Sub

Private Sub Command4_Click()
    frmControlloUtilizzo.Show vbModal
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If ConnessioneADODBLib = True Then
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    PrelevaAzienda
    sSQL = "SELECT * FROM RV_PORelease WHERE IDRV_POProgramma=" & IdentificativoProgramma
    sSQL = sSQL & " ORDER BY IDRelease DESC"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    Me.Caption = GET_CAPTION_FORM(IdentificativoProgramma, Trim(rs!Release))

    Me.lblVersione.Caption = Trim(rs!Release)
    Me.lblDataInstallazione.Caption = fnNotNull(rs!DataInstallazione)
    rs.CloseResultset
    Set rs = Nothing
    
    GET_GRIGLIA 80
    
    'GET_RECUPERA_DATI_LICENZA
    'CONTROLLO_LICENZA
    
    Me.Command1.Visible = True
End If

End Sub
Private Function GET_CAPTION_FORM(IDProgramma As Long, ReleaseModulo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & IDProgramma

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CAPTION_FORM = ""
Else
    GET_CAPTION_FORM = "Modulo " & UCase(rs!programma) & " [Release " & ReleaseModulo & "]"
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICE_DIAMANTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM ComponenteSwAbilitata "
sSQL = sSQL & "WHERE NomeCompSW=" & fnNormString("*IDSW___")

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_DIAMANTE = ""
Else
    GET_CODICE_DIAMANTE = Trim(fnNotNull(rs!descrizione))
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_NOME_PROGRAMMA(IDProgramma As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & IDProgramma

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NOME_PROGRAMMA = ""
Else
    GET_NOME_PROGRAMMA = fnNotNull(rs!programma)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_GRIGLIA(IdentificativoProgramma As Long)
On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_POProgrammaModulo "
    sSQL = sSQL & "WHERE IdentificazioneProgramma=" & IdentificativoProgramma
    
    
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection
        
        With Me.Griglia
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_POProgrammaModulo", "IDRV_POProgrammaModulo", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "DescrizioneModulo", "Modulo", dgchar, True, 5000, dgAlignleft
                .ColumnsHeader.Add "Attivato", "Attivato", dgBoolean, True, 1500, dgAligncenter
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox "Errore recupero licenza", vbCritical, "Funzionalità abilitate"
End Sub

