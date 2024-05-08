VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTracciabilita 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurazione tracciabilità web dell'azienda"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14220
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWebService 
      Height          =   285
      Left            =   9840
      TabIndex        =   19
      Top             =   5400
      Width           =   4335
   End
   Begin VB.CommandButton cmdCartellaImgSocio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      Picture         =   "frmTracciabilita.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Percorso"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Width           =   4815
   End
   Begin VB.CommandButton cmdCartellaImgArt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Picture         =   "frmTracciabilita.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Percorso"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtNomeUtente 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5400
      Width           =   4575
   End
   Begin VB.TextBox txtCodiceTracciabilita 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox RtxtPresentazione 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmTracciabilita.frx":0B14
   End
   Begin VB.CommandButton cmdSelezionaImmagine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11040
      Picture         =   "frmTracciabilita.frx":0B90
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Percorso"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdEliminaRifImg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11400
      Picture         =   "frmTracciabilita.frx":111A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Elimina riferimento immagine"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdTracciabilita 
      Caption         =   "SALVA IMPOSTAZIONI"
      Height          =   1005
      Left            =   12000
      Picture         =   "frmTracciabilita.frx":16A4
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtGoogleMaps 
      Height          =   1005
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Frame FraNonVisualizzare 
      Caption         =   "Non visualizzare"
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
      Height          =   3855
      Left            =   12000
      TabIndex        =   22
      Top             =   0
      Width           =   2175
      Begin VB.CheckBox chkIndirizzo 
         Caption         =   "Indirizzo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkTelefono 
         Caption         =   "Telefono"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkSitoInternet 
         Caption         =   "Sito internet"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkLogo 
         Caption         =   "Immagine"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkPresentazione 
         Caption         =   "Presentazione"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox chkGoogleMaps 
         Caption         =   "Google maps"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CheckBox chkPartitaIva 
         Caption         =   "Partita I.V.A."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtPercorsoImmagine 
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Indirizzo web service"
      Height          =   255
      Left            =   9840
      TabIndex        =   29
      Top             =   5160
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   5160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Nome utente"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   14160
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      Caption         =   "CODICE TRACCIABILITA' AZIENDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   11880
      X2              =   11880
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Label Label11 
      Caption         =   "Presentazione"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label11 
      Caption         =   "Indirizzo Google Maps"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Label Label11 
      Caption         =   "Immagine di presentazione"
      Height          =   255
      Index           =   2
      Left            =   7080
      TabIndex        =   23
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmTracciabilita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long



Private Const SW_SHOWNORMAL = 1
Private Const SE_ERR_FNF = 2&
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const MAX_PATH = 260


'Private Sub cmdCartellaImgArt_Click()
'Dim strCondivisione As String
'If SfogliaCondivisione(Me.hwnd, strCondivisione, 1, "Condivisione non valida!") Then
'    Me.txtImmagineArticolo.Text = strCondivisione
'End If
'End Sub
'
'Private Sub cmdCartellaImgSocio_Click()
'Dim strCondivisione As String
'If SfogliaCondivisione(Me.hwnd, strCondivisione, 1, "Condivisione non valida!") Then
'    Me.txtCartellaImgSocio.Text = strCondivisione
'End If
'End Sub

Private Sub cmdEliminaRifImg_Click()
On Error GoTo ERR_cmdEliminaRifImg_Click
Dim Testo As String
Dim sSQL As String

Testo = "Sei sicuro di voler eliminare l'immagine?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione immagine web") = vbNo Then Exit Sub

sSQL = "UPDATE RV_POSchemaCoop SET "
sSQL = sSQL & "LogoPresentazione=NULL"
sSQL = sSQL & " WHERE IDRV_POSchemaCoop=" & LINK_PARAMETRI_FILIALE

Cn.Execute sSQL

GET_DATI_CONFIGURAZIONE

Exit Sub
ERR_cmdEliminaRifImg_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifImg_Click"
End Sub

Private Sub cmdSelezionaImmagine_Click()
On Error GoTo ERR_cmdPubblica_Click
Dim RetVal

RetVal = Shell("explorer", 1)

Exit Sub

ERR_cmdPubblica_Click:
    MsgBox Err.Description, vbCritical, "cmdPubblica_Click"
End Sub

Private Sub cmdTracciabilita_Click()
On Error GoTo ERR_cmdSalvaDatiAggiuntivi_Click
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim sts As ADODB.Stream
Dim Testo As String


''''CONTROLLO INSERIMENTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(Trim(Me.txtCodiceTracciabilita.Text)) = 0 Then
    Testo = "Inserire il codice di tracciabilità web dell'azienda"
    
    MsgBox Testo, vbCritical, "Configurazione tracciabilità web"
    
    Exit Sub
End If
'If Len(Trim(Me.txtImmagineArticolo.Text)) = 0 Then
'    Testo = "Inserire il percorso di rete della cartella immagini degli articoli"
'
'    MsgBox Testo, vbCritical, "Configurazione tracciabilità web"
'
'    Exit Sub
'End If
'
'If Len(Trim(Me.txtCartellaImgSocio.Text)) = 0 Then
'    Testo = "Inserire il percorso di rete della cartella dei produttori"
'
'    MsgBox Testo, vbCritical, "Configurazione tracciabilità web"
'
'    Exit Sub
'End If

If Len(Trim(Me.txtWebService.Text)) = 0 Then
    Testo = "Inserire l'indirizzo web della tracciabilità web"
    
    MsgBox Testo, vbCritical, "Configurazione tracciabilità web"
    
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & LINK_PARAMETRI_FILIALE


Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
End If

rs!TestoPresentazione = rtf2html3(Me.RtxtPresentazione.TextRTF)
rs!TestoPresentazioneRTF = Me.RtxtPresentazione.TextRTF
rs!TestoGoogleMaps = Me.txtGoogleMaps.Text
rs!NonVisualizzareIndirizzo = Me.chkIndirizzo.Value
rs!NonVisualizzareTelefono = Me.chkTelefono.Value
rs!NonVisualizzareFax = Me.chkFax.Value
rs!NonVisualizzareEmail = Me.chkEmail.Value
rs!NonVisualizzareSitoWeb = Me.chkSitoInternet.Value
rs!NonVisualizzareLogo = Me.chkLogo.Value
rs!NonVisualizzarePartitaIva = Me.chkPartitaIva.Value
rs!NonVisualizzarePresentazione = Me.chkPresentazione.Value
rs!NonVisualizzareGoogleMaps = Me.chkGoogleMaps.Value
rs!WebService = Me.txtWebService.Text
rs!CodiceTraccAzienda = Me.txtCodiceTracciabilita.Text
rs!NomeUtenteTracciabilitaWeb = Me.txtNomeUtente.Text
rs!PasswordUtenteTracciabilitaWeb = Me.txtPassword.Text
If Len(Trim(Me.txtPercorsoImmagine.Text)) > 0 Then
    Set sts = New ADODB.Stream
    sts.Type = ADODB.adTypeBinary
    sts.Open
    sts.LoadFromFile Me.txtPercorsoImmagine.Text
    rs!LogoPresentazione = sts.Read
    
    sts.Close
    Set sts = Nothing
End If

rs.Update

rs.Close
Set rs = Nothing

MsgBox "Configurazione avvenuta con successo", vbInformation, "Configurazione tracciabilità web"

CONFERMA_TRACC_ONLINE = 1

Unload Me
Exit Sub
ERR_cmdSalvaDatiAggiuntivi_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdSalvaDatiAggiuntivi_Click"
End Sub

Private Sub Form_Load()

    CONFERMA_TRACC_ONLINE = 0

    GET_DATI_CONFIGURAZIONE
End Sub

Private Sub GET_DATI_CONFIGURAZIONE()
On Error GoTo ERR_GET_DATI_CONFIGURAZIONE
Dim sSQL As String
Dim rs As ADODB.Recordset

Me.txtPercorsoImmagine.Text = ""


sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & LINK_PARAMETRI_FILIALE

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

If rs.EOF Then
    Me.RtxtPresentazione.TextRTF = ""
    Me.txtGoogleMaps.Text = ""
    Me.chkIndirizzo.Value = 0
    Me.chkTelefono.Value = 0
    Me.chkFax.Value = 0
    Me.chkEmail.Value = 0
    Me.chkSitoInternet.Value = 0
    Me.chkLogo.Value = 0
    Me.chkPresentazione.Value = 0
    Me.chkGoogleMaps.Value = 0
    Me.chkPartitaIva.Value = 0
    Me.Picture1.Picture = LoadPicture("")
    
    Me.txtCodiceTracciabilita.Text = ""
    
    Me.txtWebService.Text = ""
    Me.txtNomeUtente.Text = ""
    Me.txtPassword.Text = ""
Else
    Me.txtCodiceTracciabilita.Text = fnNotNull(rs!CodiceTraccAzienda)
    Me.txtWebService.Text = fnNotNull(rs!WebService)
    Me.RtxtPresentazione.TextRTF = fnNotNull(rs!TestoPresentazioneRTF)
    Me.txtGoogleMaps.Text = fnNotNull(rs!TestoGoogleMaps)
    Me.chkIndirizzo.Value = Abs(fnNotNullN(rs!NonVisualizzareIndirizzo))
    Me.chkTelefono.Value = Abs(fnNotNullN(rs!NonVisualizzareTelefono))
    Me.chkFax.Value = Abs(fnNotNullN(rs!NonVisualizzareFax))
    Me.chkEmail.Value = Abs(fnNotNullN(rs!NonVisualizzareEmail))
    Me.chkSitoInternet.Value = Abs(fnNotNullN(rs!NonVisualizzareSitoWeb))
    Me.chkLogo.Value = Abs(fnNotNullN(rs!NonVisualizzareLogo))
    Me.chkPresentazione.Value = Abs(fnNotNullN(rs!NonVisualizzarePresentazione))
    Me.chkGoogleMaps.Value = Abs(fnNotNullN(rs!NonVisualizzareGoogleMaps))
    Me.chkPartitaIva.Value = Abs(fnNotNullN(rs!NonVisualizzarePartitaIva))
    Me.txtNomeUtente.Text = fnNotNull(rs!NomeUtenteTracciabilitaWeb)
    Me.txtPassword.Text = fnNotNull(rs!PasswordUtenteTracciabilitaWeb)
    
    If IsNull(rs!LogoPresentazione) = False Then
        RIPRISTINO_IMMAGINE
        Me.Picture1.Picture = LoadPicture(AVVIA_FILE(rs, Me.txtCodiceTracciabilita.Text))
        DIMENSIONA_IMMAGINE
    Else
        Me.Picture1.Picture = LoadPicture("")
    End If
End If

rs.Close
Set rs = Nothing

Exit Sub
ERR_GET_DATI_CONFIGURAZIONE:
    MsgBox Err.Description, vbCritical, "GET_DATI_CONFIGURAZIONE"
    
End Sub
Private Sub DIMENSIONA_IMMAGINE()
Dim START_REDIM As Boolean

START_REDIM = False

If Me.Picture1.Width > 4695 Then
    Me.Picture1.Width = 4695
    START_REDIM = True
End If
If Me.Picture1.Height > 4335 Then
    Me.Picture1.Height = 4335
    If START_REDIM = False Then
        START_REDIM = True
    End If
End If

If START_REDIM = True Then
    FitPicture Me.Picture1.Picture, Picture1
End If
End Sub
Private Sub RIPRISTINO_IMMAGINE()
Me.Picture1.Width = 100
Me.Picture1.Height = 100
End Sub
'Private Sub txtCartellaImgSocio_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF6 Then
'        If Me.txtCartellaImgSocio.Locked = False Then
'            Me.txtCartellaImgSocio.Locked = True
'        Else
'            Me.txtCartellaImgSocio.Locked = False
'        End If
'    End If
'End Sub

Private Sub txtCodiceTracciabilita_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        If Me.txtCodiceTracciabilita.Locked = True Then
            Me.txtCodiceTracciabilita.Locked = False
        Else
            Me.txtCodiceTracciabilita.Locked = True
        End If
    End If
End Sub

'Private Sub txtImmagineArticolo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF6 Then
'        If Me.txtImmagineArticolo.Locked = False Then
'            Me.txtImmagineArticolo.Locked = True
'        Else
'            Me.txtImmagineArticolo.Locked = False
'        End If
'    End If
'End Sub

Private Sub txtPercorsoImmagine_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_txtPercorsoImmagine_OLEDragDrop
Dim START_REDIM As Boolean
    
    Me.txtPercorsoImmagine.Text = Data.Files.Item(1)
    
    If Len(Trim(Me.txtPercorsoImmagine.Text)) > 0 Then
        RIPRISTINO_IMMAGINE
        Me.Picture1.Picture = LoadPicture(Me.txtPercorsoImmagine.Text)
        DIMENSIONA_IMMAGINE
        
    End If
    
Exit Sub
ERR_txtPercorsoImmagine_OLEDragDrop:
    MsgBox Err.Description, vbCritical, "Selezione immagine Azienda"
End Sub
Public Sub FitPicture(srcImage As StdPicture, dstPicBox As PictureBox)
Dim pW As Long
Dim pH As Long
Dim r As Double


pW = (dstPicBox.ScaleX(srcImage.Width, vbHimetric, vbTwips))
pH = (dstPicBox.ScaleY(srcImage.Height, vbHimetric, vbTwips))

If srcImage.Width >= srcImage.Height Then
    r = pW / dstPicBox.ScaleWidth
Else
    r = pH / dstPicBox.ScaleHeight
End If

dstPicBox.AutoRedraw = True
dstPicBox.Cls
dstPicBox.Picture = LoadPicture
dstPicBox.PaintPicture srcImage, (dstPicBox.ScaleWidth - (pW / r)) / 2, (dstPicBox.ScaleHeight - (pH / r)) / 2, pW / r, pH / r

End Sub
Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
End Function

Private Function AVVIA_FILE(rs As ADODB.Recordset, CodiceSocio As String) As String
Dim sts As ADODB.Stream
Dim Scr_hDC As Long
Dim X
Dim msg As String
Dim PercorsoDocArchiviati As String
Dim F As FileSystemObject
Dim NomeCartella As String

    NomeCartella = TrovaCartella(CSIDL_COMMON_APPDATA) & "GreenTop\"
    Set F = New FileSystemObject
        If F.FolderExists(NomeCartella) = False Then
            F.CreateFolder NomeCartella
        End If
    Set F = Nothing

    Set sts = New ADODB.Stream
    sts.Type = ADODB.adTypeBinary
    sts.Open
    sts.Write rs.Fields("LogoPresentazione").Value
    sts.SaveToFile NomeCartella & Me.txtCodiceTracciabilita.Text & ".jpg", adSaveCreateOverWrite

    
    sts.Close
    Set sts = Nothing
    
    AVVIA_FILE = NomeCartella & CodiceSocio & ".jpg"

End Function

Private Sub txtWebService_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        If Me.txtWebService.Locked = False Then
            Me.txtWebService.Locked = True
        Else
            Me.txtWebService.Locked = False
        End If
    End If
End Sub
