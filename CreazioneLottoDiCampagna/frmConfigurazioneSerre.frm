VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmConfigurazioneSerre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurazione Serre/Appezzamenti"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6270
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
   ScaleHeight     =   6345
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicIntestazione 
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   6105
      TabIndex        =   10
      Top             =   720
      Width           =   6135
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "SERRE/APPEZZAMENTI"
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
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   12
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "SETTORI"
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
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
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
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   6105
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdConferma 
         Caption         =   "Conferma"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin DMTDataCmb.DMTCombo cboSchema 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.Label Label2 
         Caption         =   "Schema"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   3495
      End
   End
   Begin VB.PictureBox PicCont 
      Appearance      =   0  'Flat
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
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5265
      ScaleWidth      =   6105
      TabIndex        =   4
      Top             =   960
      Width           =   6135
      Begin VB.VScrollBar VScroll2 
         Height          =   5285
         Left            =   5840
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
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
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   5775
         TabIndex        =   5
         Top             =   0
         Width           =   5775
         Begin VB.CheckBox chkLotti 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   1
            Top             =   340
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CheckBox chkSerre 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox chkSettore 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            X1              =   2880
            X2              =   2880
            Y1              =   120
            Y2              =   5160
         End
      End
   End
   Begin VB.CommandButton cmdFiller 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmConfigurazioneSerre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AGGIORNAMENTO As Integer
Private Link_Settore As Long
Private Link_Serra As Long
Private Link_Lotto As Long
Private ValoreCampo As Integer
Private VarTop_Totale As Long
Private VarTop_TotaleSerre As Long
Private VarWidth_Pic As Long
Private VarTop_TotaleSettori As Long


Const SB_WIDTH = 300
Const SB_HEIGHT = 300

Private Sub cboSchema_Click()
Dim Testo As String

Testo = "ATTENZIONE!!!" & vbCrLf
Testo = Testo & "Cambiando lo schema verranno eliminati tutti i dati delle impostazioni delle superfici." & vbCrLf
Testo = Testo & "Vuoi continuare?"

If Me.cboSchema.CurrentID > 0 Then
    If Me.cboSchema.CurrentID <> frmMain.cboSchema.CurrentID Then
        If MsgBox(Testo, vbInformation + vbYesNo, "Configurazione superficie") = vbYes Then
            Riconfigura
        End If
    Else
        AGGIORNAMENTO = 1
        VarWidth_Pic = 5775
        Me.Pic1.Height = 5175
        
        ConfigurazioneSettori
        
        ConfigurazioneSerre
        
        If VarTop_TotaleSerre >= VarTop_TotaleSettori Then
            If VarTop_TotaleSerre >= Me.Pic1.Height Then
                Me.Pic1.Height = VarTop_TotaleSerre
                Me.Line1.Y2 = Me.Pic1.Height - 110
            End If
        Else
            If VarTop_TotaleSettori >= Me.Pic1.Height Then
                Me.Pic1.Height = VarTop_TotaleSettori
                Me.Line1.Y2 = Me.Pic1.Height - 110
            End If
        End If
        
        Me.Pic1.Width = VarWidth_Pic
        Me.VScroll2.Left = Me.Pic1.Left + Me.Pic1.Width
        Me.PicCont.Width = Me.VScroll2.Left + Me.VScroll2.Width + 120
        Me.Width = Me.PicCont.Width + 120
        Me.PicIntestazione.Width = Me.PicCont.Width
        Me.Pic2.Width = Me.PicCont.Width
        
        GET_SERRAPERLOTTO
        
        GET_SETTOREPERLOTTO
        
        GET_ATTIVAZIONE_SCROLL
        AGGIORNAMENTO = 0
    End If
    
End If

End Sub
Private Sub EliminaRiferimentiSchedaSuperficie()
Dim sSQL As String

sSQL = "DELETE FROM RV_PO01_SchedaPerSerra WHERE IDRV_PO01_SchedaTrattamenti=" & fnNotNullN(Link_Scheda)
Cn.Execute sSQL

sSQL = "DELETE FROM RV_PO01_SchedaPerLotto WHERE IDRV_PO01_SchedaTrattamenti=" & fnNotNullN(Link_Scheda)
Cn.Execute sSQL

sSQL = "DELETE FROM RV_PO01_SchedaPerSettore WHERE IDRV_PO01_SchedaTrattamenti=" & fnNotNullN(Link_Scheda)
Cn.Execute sSQL

End Sub

Private Sub chkSerre_Click(Index As Integer)
'On Error Resume Next

Dim I As Integer

Dim ArraySettori(500, 1) As Long
Dim ArrayLotti(500, 1) As Long
Dim X As Integer 'Totale serre del settore
Dim Y As Integer 'Totale serre vistate

Dim Valore As Integer
Dim IControl As Integer
Dim ctrl As Control
Dim Link_Settore As Long

If AGGIORNAMENTO = 0 Then
AGGIORNAMENTO = 1
'    Link_Serra = Me.chkSerre(Index).Tag
    Select Case Me.chkSerre(Index).Value
        Case Checked
            ValoreCampo = 1
        Case Unchecked
            ValoreCampo = 0
    End Select
    
    Link_Settore = fnNotNullN(Me.chkSerre(Index).Tag)



    For Each ctrl In Me.Controls
        Select Case ctrl.Name
            Case "chkSerre"
                If ctrl.Index > 0 Then
                    If fnNotNullN(ctrl.Tag) = Link_Settore Then
                        X = X + 1
                        If ctrl.Value = vbChecked Then
                            Y = Y + 1
                        End If
                    End If
                End If
        End Select
    Next
    
            
        If (X - Y) = 0 Then
            Valore = 1
        End If
        If (X - Y) = X Then
            Valore = 0
        End If
        If ((X - Y) > 0) And ((X - Y) <> X) Then
            Valore = 2
        End If
        'Creazione del controllo che contiene le informazioni del settore
        'On Error Resume Next
        Me.chkSettore(Link_Settore).Value = Valore

    
        
    AGGIORNAMENTO = 0
Else

End If


End Sub

Private Sub chkSettore_Click(Index As Integer)
Dim sSQL As String
Dim I As Integer
Dim rs As DmtOleDbLib.adoResultset
Dim rsSerra As DmtOleDbLib.adoResultset
Dim Valore As Integer
Dim IControl As Integer
Dim X As Integer
Dim Y As Integer
Dim ctrl As Control

If AGGIORNAMENTO = 0 Then
AGGIORNAMENTO = 1
    Link_Settore = Me.chkSettore(Index).Tag
    
    Select Case Me.chkSettore(Index).Value
        Case Checked
            ValoreCampo = 1
        Case Unchecked
            ValoreCampo = 0
    End Select
    

    'ELIMINA SERRE
    For Each ctrl In Me.Controls
        Select Case ctrl.Name
            Case "chkSerre"
                If ctrl.Index > 0 Then
                    If ctrl.Tag = Link_Settore Then
                        ctrl.Value = ValoreCampo
                    End If
                End If
        End Select
    Next
AGGIORNAMENTO = 0
End If

End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
Dim I As Integer
Dim rsCTRL As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim TotaleDimensioneSerreDisp As Double
Dim TotaleDimensioneSerreOccupata As Double
Dim DimensioneHaSerra As String
Dim Testo As String
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long


If Link_Schema > 0 Then
    If Link_Schema <> Me.cboSchema.CurrentID Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Nella configurazione del lotto di campagna potrebbero essere presenti "
        Testo = Testo & "serre\appezzamenti che non fanno parte dello schema selezionato." & vbCrLf
        Testo = Testo & "Se si continua con questo comando verranno eliminati tutti i riferimenti delle serre\appezzamenti "
        Testo = Testo & "precedentemente salvati." & vbCrLf
        Testo = Testo & "Continuare con questo comando?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Configurazione serre/Appezzamenti") = vbNo Then
            If MsgBox("Ripristinare lo schema precedentemente salvato?", vbQuestion + vbYesNo, "Ripristino configurazione") = vbYes Then
                Me.cboSchema.WriteOn Link_Schema
            End If
            Exit Sub
        Else
            sSQL = "DELETE FROM RV_PO01_SerraPerLotto "
            sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & Link_LottoCampagna
            Cn.Execute sSQL
        End If
    End If
End If


SCRIVI_CODA Link_LottoCampagna
APERTURA_FORM_CODA = False

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
        If X = -1 Then
            X = 1
            ErroreCoda = True
        End If
    Loop Until X = 1
    
    If ErroreCoda = True Then
        X = -1
    End If
    
    If X = -1 Then
        Me.Enabled = True
        Me.SetFocus
        Me.Caption = "Configurazione serre"
        Screen.MousePointer = 0
        ''''''''ELIMINAZIONE RIFERIMENTO CODA'''''''''''''''''''''''''''''''
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
        sSQL = sSQL & " AND IDTipoOggetto=" & 0
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Me.Enabled = True
    Me.SetFocus
    Me.Caption = "Configurazione serre"
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    'frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    'frmAttesa.lblInfo = Me.Caption
    DoEvents

    sSQL = "SELECT RV_PO01_SettoreSerra.IDRV_PO01_SettoreSerra, RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema, RV_PO01_SettoreSerra.IDRV_PO01_Serra, "
    sSQL = sSQL & "RV_PO01_SettoreSchema.IDRV_PO01_Settore, RV_PO01_SettoreSchema.IDRV_PO01_Schema, RV_PO01_SettoreSerra.DimensioneMQ,"
    sSQL = sSQL & "RV_PO01_SettoreSerra.DimensioneHA , RV_PO01_Serra.Codice, RV_PO01_Serra.Descrizione "
    sSQL = sSQL & "FROM RV_PO01_SettoreSerra INNER JOIN "
    sSQL = sSQL & "RV_PO01_SettoreSchema ON "
    sSQL = sSQL & "RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema = RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Serra ON RV_PO01_SettoreSerra.IDRV_PO01_Serra = RV_PO01_Serra.IDRV_PO01_Serra "
    sSQL = sSQL & "WHERE RV_PO01_SettoreSchema.IDRV_PO01_Schema = " & Me.cboSchema.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
        
    While Not rs.EOF
        If Me.chkSerre(rs!IDRV_PO01_Serra).Value = Checked Then
            sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
            sSQL = sSQL & "WHERE IDRV_PO01_Serra=" & fnNotNullN(rs!IDRV_PO01_Serra)
            sSQL = sSQL & " AND IDRV_PO01_LottoCampagna=" & Link_LottoCampagna
            
            Set rsCTRL = Cn.OpenResultset(sSQL)
            
                If rsCTRL.EOF Then
                    
                    TotaleDimensioneSerreOccupata = GET_DIMENSIONE_SERRA_DISP(fnNotNullN(rs!IDRV_PO01_Serra), Me.cboSchema.CurrentID, fnNotNullN(rs!DimensioneMq))
                    TotaleDimensioneSerreDisp = fnNotNullN(rs!DimensioneMq) - TotaleDimensioneSerreOccupata
                    If TotaleDimensioneSerreDisp > 0 Then
                        DimensioneHaSerra = GET_DIMENSIONE_SERRA_HA(TotaleDimensioneSerreDisp)
                        sSQL = "INSERT INTO RV_PO01_SerraPerLotto ("
                        sSQL = sSQL & "IDRV_PO01_SerraPerLotto, IDRV_PO01_Serra, IDRV_PO01_LottoCampagna, "
                        sSQL = sSQL & "DimensioneMQ, DimensioneHA, GuidID) "
                        sSQL = sSQL & "VALUES ("
                        sSQL = sSQL & fnGetNewKey("RV_PO01_SerraPerLotto", "IDRV_PO01_SerraPerLotto") & ", "
                        sSQL = sSQL & Me.chkSerre(fnNotNullN(rs!IDRV_PO01_Serra)).Index & ", "
                        sSQL = sSQL & Link_LottoCampagna & ", "
                        sSQL = sSQL & fnNormNumber(TotaleDimensioneSerreDisp) & ", "
                        sSQL = sSQL & fnNormString(DimensioneHaSerra) & ", "
                        sSQL = sSQL & fnNormString(GetGUID) & ")"
                        Cn.Execute sSQL
                    End If
                End If
            rsCTRL.CloseResultset
            Set rsCTRL = Nothing
        Else
            sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
            sSQL = sSQL & "WHERE IDRV_PO01_Serra=" & fnNotNullN(rs!IDRV_PO01_Serra)
            sSQL = sSQL & " AND IDRV_PO01_LottoCampagna=" & Link_LottoCampagna
            
            Set rsCTRL = Cn.OpenResultset(sSQL)
            If rsCTRL.EOF = False Then
                sSQL = "DELETE FROM RV_PO01_SerraPerLotto "
                sSQL = sSQL & "WHERE IDRV_PO01_Serra=" & fnNotNullN(rs!IDRV_PO01_Serra)
                sSQL = sSQL & " AND IDRV_PO01_LottoCampagna=" & Link_LottoCampagna
                Cn.Execute sSQL
            End If
            rsCTRL.CloseResultset
            Set rsCTRL = Nothing
        End If
    rs.MoveNext
    Wend

rs.CloseResultset
Set rs = Nothing



'Unload frmAttesa
Me.Enabled = True
Me.SetFocus
Me.Caption = "Configurazione serre"
Cn.CursorLocation = OLD_Cursor

''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto(App.EXEName)
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Link_Schema = Me.cboSchema.CurrentID
CONFERMA_SELEZIONE_SERRE = True
Unload Me


Exit Sub
ERR_cmdConferma_Click:
    'Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = "Configurazione serre/appezzamenti"
End Sub
Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    'Schema per settori
    With Me.cboSchema
        Set .Database = Cn
        .AddFieldKey "IDRV_PO01_Schema"
        .DisplayField = "SchemaSettori"
        .SQL = "SELECT * FROM RV_PO01_Schema "
        .SQL = .SQL & "WHERE IDAzienda=" & VarIDAzienda
        .SQL = .SQL & " AND IDSocio=" & Link_Socio
        .Fill
    End With
    AGGIORNAMENTO = 1
    
    If Link_Schema = 0 Then
        Me.cboSchema.WriteOn GET_SCHEMAPREDEFINITO
    Else
        Me.cboSchema.WriteOn Link_Schema
    End If
    
    Me.Pic1.Height = 5175
    
    'If Me.cboSchema.CurrentID > 0 Then
        
    '    ConfigurazioneSettori
        
    '    ConfigurazioneSerre
        
    If VarTop_TotaleSerre >= VarTop_TotaleSettori Then
        If VarTop_TotaleSerre >= Me.Pic1.Height Then
            Me.Pic1.Height = VarTop_TotaleSerre
            Me.Line1.Y2 = Me.Pic1.Height - 110
        End If
    Else
        If VarTop_TotaleSettori >= Me.Pic1.Height Then
            Me.Pic1.Height = VarTop_TotaleSettori
            Me.Line1.Y2 = Me.Pic1.Height - 110
        End If
    End If
        
    'GET_SERRAPERLOTTO
        
    'End If
    
    AGGIORNAMENTO = 0

    'GET_ATTIVAZIONE_SCROLL
    
End Sub
Private Sub GET_SERRAPERLOTTO()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCTRL As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_SerraPerLotto "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & Link_LottoCampagna

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Me.chkSerre(fnNotNullN(rs!IDRV_PO01_Serra)).Value = Checked
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_SCHEMAPREDEFINITO() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PO01_Schema "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocio.KeyFieldID
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_SCHEMAPREDEFINITO = 0
Else
    GET_SCHEMAPREDEFINITO = fnNotNullN(rs!IDRV_PO01_Schema)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ConfigurazioneSettori()
Dim sSQL As String
Dim rsSchema As DmtOleDbLib.adoResultset
Dim I As Integer
Dim VarTop As Long
Dim ctrl As Control

'ELIMINA SETTORE
For Each ctrl In Me.Controls
    Select Case ctrl.Name
        Case "chkSettore"
            If ctrl.Index > 0 Then
                Unload ctrl
            End If
    End Select
Next


    sSQL = "SELECT RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema, RV_PO01_SettoreSchema.IDRV_PO01_Settore, RV_PO01_SettoreSchema.IDRV_PO01_Schema, "
    sSQL = sSQL & "RV_PO01_SettoreSchema.DimensioneMQ , RV_PO01_SettoreSchema.DimensioneHA, RV_PO01_Settore.Descrizione, RV_PO01_Settore.Codice "
    sSQL = sSQL & "FROM RV_PO01_SettoreSchema INNER JOIN "
    sSQL = sSQL & "RV_PO01_Schema ON RV_PO01_SettoreSchema.IDRV_PO01_Schema = RV_PO01_Schema.IDRV_PO01_Schema LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Settore ON RV_PO01_SettoreSchema.IDRV_PO01_Settore = RV_PO01_Settore.IDRV_PO01_Settore "
    sSQL = sSQL & "WHERE RV_PO01_SettoreSchema.IDRV_PO01_Schema=" & Me.cboSchema.CurrentID
    sSQL = sSQL & " ORDER BY RV_PO01_Settore.Codice"
    
    Set rsSchema = Cn.OpenResultset(sSQL)
    I = 1
    VarTop = 120
    While Not rsSchema.EOF
        Load chkSettore(rsSchema!IDRV_PO01_Settore)
        With chkSettore(rsSchema!IDRV_PO01_Settore)
            '.Left = 120
            .Top = VarTop
            .Tag = rsSchema!IDRV_PO01_Settore
            .Caption = Trim(rsSchema!Codice)
            .Visible = True
            .ZOrder 0
            '.Width = 2055
            .Value = Unchecked
        End With
        VarTop = VarTop + 220
        I = I + 1
    rsSchema.MoveNext
    Wend
    rsSchema.CloseResultset
    Set rsSchema = Nothing
    
    VarTop_TotaleSettori = VarTop
End Sub

Private Sub ConfigurazioneSerre()
Dim sSQL As String
Dim rsSchema As DmtOleDbLib.adoResultset
Dim rsCTRL As DmtOleDbLib.adoResultset
Dim I As Integer
Dim VarTop As Long
Dim VarTop_TotaleSerreLocal As Long
Dim LeftLocal As Long
Dim NumeroColonne As Long
Dim ctrl As Control

'ELIMINA SERRE



    sSQL = "SELECT RV_PO01_SettoreSerra.IDRV_PO01_SettoreSerra, dbo.RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema, "
    sSQL = sSQL & "RV_PO01_SettoreSerra.IDRV_PO01_Serra, dbo.RV_PO01_SettoreSerra.DimensioneMQ, dbo.RV_PO01_SettoreSerra.DimensioneHA,"
    sSQL = sSQL & "RV_PO01_SettoreSchema.IDRV_PO01_Settore, dbo.RV_PO01_Schema.IDRV_PO01_Schema, dbo.RV_PO01_Serra.Codice,"
    sSQL = sSQL & "RV_PO01_Serra.Descrizione "
    sSQL = sSQL & "FROM RV_PO01_SettoreSerra INNER JOIN "
    sSQL = sSQL & "RV_PO01_SettoreSchema ON "
    sSQL = sSQL & "RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema = dbo.RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema INNER JOIN "
    sSQL = sSQL & "RV_PO01_Schema ON dbo.RV_PO01_SettoreSchema.IDRV_PO01_Schema = dbo.RV_PO01_Schema.IDRV_PO01_Schema LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Serra ON dbo.RV_PO01_SettoreSerra.IDRV_PO01_Serra = dbo.RV_PO01_Serra.IDRV_PO01_Serra "
    sSQL = sSQL & "WHERE RV_PO01_Schema.IDRV_PO01_Schema = " & Me.cboSchema.CurrentID
    sSQL = sSQL & " ORDER BY RV_PO01_Serra.Codice"

    Set rsSchema = Cn.OpenResultset(sSQL)
    I = 1
    VarTop = 120
    NumeroColonne = 1
    While Not rsSchema.EOF
        Load chkSerre(rsSchema!IDRV_PO01_Serra)
        With chkSerre(rsSchema!IDRV_PO01_Serra)
            Select Case NumeroColonne
                Case 1
                    LeftLocal = Me.chkSerre(0).Left
                Case 2
                    LeftLocal = Me.chkSerre(0).Left + (Me.chkSerre(0).Width * (NumeroColonne - 1))
                Case 3
                    LeftLocal = Me.chkSerre(0).Left + (Me.chkSerre(0).Width * (NumeroColonne - 1))
                Case 4
                    LeftLocal = Me.chkSerre(0).Left + (Me.chkSerre(0).Width * (NumeroColonne - 1))
                Case 5
                    LeftLocal = Me.chkSerre(0).Left + (Me.chkSerre(0).Width * (NumeroColonne - 1))
                Case 6
            End Select
            .Top = VarTop
            .Left = LeftLocal
            .Tag = rsSchema!IDRV_PO01_Settore
            .Caption = Trim(rsSchema!Codice)
            .Visible = True
            .Value = vbUnchecked
            .ZOrder 0
        End With
        
        VarTop = VarTop + 220
        
        If VarTop >= 32400 Then
            NumeroColonne = NumeroColonne + 1
            VarTop_TotaleSerre = VarTop + 60
            VarTop = 120
        End If
        
        I = I + 1
        
    rsSchema.MoveNext
    Wend
    rsSchema.CloseResultset
    Set rsSchema = Nothing
    If NumeroColonne = 1 Then
        VarTop_TotaleSerre = VarTop
    Else
        VarWidth_Pic = VarWidth_Pic + (Me.chkSerre(0).Width * (NumeroColonne - 1))
    End If
    
End Sub

Private Sub Riconfigura()
Dim I As Integer
Dim ctrl As Control

For Each ctrl In Me.Controls
    Select Case ctrl.Name
        Case "chkSettore"
            If ctrl.Index > 0 Then
                Unload ctrl
            End If
        Case "chkLotti"
            If ctrl.Index > 0 Then
                Unload ctrl
            End If
        Case "chkSerre"
            If ctrl.Index > 0 Then
                Unload ctrl
            End If
    End Select
Next

VarWidth_Pic = 5775
Me.Pic1.Height = 5175

AGGIORNAMENTO = 1
If Me.cboSchema.CurrentID > 0 Then
    
    ConfigurazioneSettori
    
    ConfigurazioneSerre
    
    If VarTop_TotaleSerre >= VarTop_TotaleSettori Then
        If VarTop_TotaleSerre >= Me.Pic1.Height Then
            Me.Pic1.Height = VarTop_TotaleSerre + 10
            Me.Line1.Y2 = Me.Pic1.Height - 135
        End If
    Else
        If VarTop_TotaleSettori >= Me.Pic1.Height Then
            Me.Pic1.Height = VarTop_TotaleSettori + 10
            Me.Line1.Y2 = Me.Pic1.Height - 135
        End If
    End If

    Me.Pic1.Width = VarWidth_Pic
    Me.VScroll2.Left = Me.Pic1.Left + Me.Pic1.Width
    Me.PicCont.Width = Me.VScroll2.Left + Me.VScroll2.Width + 120
    Me.Width = Me.PicCont.Width + 120
    Me.PicIntestazione.Width = Me.PicCont.Width
    Me.Pic2.Width = Me.PicCont.Width

    GET_ATTIVAZIONE_SCROLL
AGGIORNAMENTO = 0
End If

End Sub

Private Function GET_DIMENSIONE_SERRA_DISP(IDSerra As Long, IDSchema As Long, DimSerra As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(RV_PO01_SerraPerLotto.DimensioneMQ) AS TotaleDimensione "
sSQL = sSQL & "FROM RV_PO01_SettoreSchema INNER JOIN "
sSQL = sSQL & "RV_PO01_SettoreSerra ON RV_PO01_SettoreSchema.IDRV_PO01_SettoreSchema = RV_PO01_SettoreSerra.IDRV_PO01_SettoreSchema INNER JOIN "
sSQL = sSQL & "RV_PO01_Schema ON RV_PO01_SettoreSchema.IDRV_PO01_Schema = RV_PO01_Schema.IDRV_PO01_Schema INNER JOIN "
sSQL = sSQL & "RV_PO01_SerraPerLotto ON RV_PO01_SettoreSerra.IDRV_PO01_Serra = RV_PO01_SerraPerLotto.IDRV_PO01_Serra INNER JOIN "
sSQL = sSQL & "RV_PO01_LottoCampagna ON RV_PO01_SerraPerLotto.IDRV_PO01_LottoCampagna = RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE (RV_PO01_LottoCampagna.Chiuso =" & fnNormBoolean(0) & ") "
sSQL = sSQL & "AND (RV_PO01_Schema.IDRV_PO01_Schema = " & IDSchema & ") "
sSQL = sSQL & "AND (RV_PO01_SerraPerLotto.IDRV_PO01_Serra = " & IDSerra & ") "


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Or frmMain.chkChiuso.Value = 1 Then
    GET_DIMENSIONE_SERRA_DISP = 0
Else
    GET_DIMENSIONE_SERRA_DISP = fnNotNullN(rs!TotaleDimensione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DIMENSIONE_SERRA_HA(DimensioneMq As Double) As String
Dim I As Integer
Dim X As String 'Stringa da formattare
Dim J As String '
Dim Ris As String
Const MAX_X As Integer = 10
Dim Rimanenza As Integer
Dim Valore As Long
Valore = DimensioneMq

Rimanenza = MAX_X - Len(CStr(Valore))

For I = 1 To Rimanenza
    X = X & "0"
Next
    'Stringa da formattare
    X = X & CStr(Valore)
    J = ""
For I = 1 To Len(X)
    
    J = J & Mid(X, I, 1)
    If I < Len(X) Then
        If I Mod 2 = 0 Then
            J = J & "."
        End If
    End If
Next
GET_DIMENSIONE_SERRA_HA = J
End Function

Private Sub SCRIVI_CODA(IDOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID

Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rs = New ADODB.Recordset

rs.Open "RV_POTMP", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDSessione = fnGetNewKey("RV_POTMP", "IDSessione")
    rs!IDUtente = TheApp.IDUser
    rs!IDTipoOggetto = fnGetTipoOggetto(App.EXEName)
    rs!IDOggetto = IDOggetto
    rs!Utente = TheApp.User
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Function GET_NUMERO_DOCUMENTO(NuovoDocumento As Boolean) As Long
On Error GoTo ERR_GET_NUMERO_DOCUMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim X_FRM As Form
Dim OLD_Cursor As Long

GET_NUMERO_DOCUMENTO = 0

sSQL = "SELECT * FROM RV_POTMP "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto(App.EXEName)
sSQL = sSQL & " ORDER BY IDSessione, IDUtente"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDUtente) = TheApp.IDUser Then
        Me.Caption = "SALVATAGGIO IN CORSO.........."
        
        DoEvents
       
        'If APERTURA_FORM_CODA = True Then
        '    Unload frmCoda
        '    APERTURA_FORM_CODA = False
        'End If
        
        GET_NUMERO_DOCUMENTO = 1
        
        rs.CloseResultset
        Set rs = Nothing
    Else
        rs.CloseResultset
        Set rs = Nothing
    
        'If APERTURA_FORM_CODA = False Then
        '    APERTURA_FORM_CODA = True
        '    Me.Enabled = False
        '    frmCoda.Show
        'End If
        
        Me.Caption = "ATTENDERE......."
        DoEvents
        'GET_NUMERO_DOCUMENTO NuovoDocumento
        
    End If
End If
Exit Function

ERR_GET_NUMERO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "Errore coda"
    GET_NUMERO_DOCUMENTO = -1
    Unload frmCoda
End Function


Private Function fnGetTipoOggetto(NomeGestore) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(NomeGestore)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub GET_ATTIVAZIONE_SCROLL()
On Error Resume Next

    With Me.VScroll2
        '.Top = Me.Pic1.Top
        .Max = (Pic1.ScaleHeight)
        .LargeChange = .Max \ 10
        .SmallChange = .Max \ 10
    End With

    If Me.PicCont.ScaleHeight < Me.Pic1.ScaleHeight Then
        Me.VScroll2.Visible = True
        'Me.VScroll2.Top = Me.Pic1.Top
        'Me.VScroll2.Left = Me.PicCont.ScaleWidth - Me.VScroll2.Width
    Else
        Me.VScroll2.Visible = False
    End If
    

    With VScroll2
        .Max = (Pic1.ScaleHeight - PicCont.ScaleHeight) ' + Me.HScroll1.Height)
        If .Max > 0 Then
            .LargeChange = .Max \ 10
            .SmallChange = .Max \ 10
        End If
    End With

End Sub

Private Sub VScroll2_Change()
    Me.Pic1.Top = -VScroll2.Value
End Sub
Private Sub VScroll2_Scroll()
   Me.Pic1.Top = -VScroll2.Value
End Sub
Private Function GET_NUMERO_SERRE_SETTORE_SCHEMA(IDSchema As Long, IDSettore As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = ""

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_SETTOREPERLOTTO()
Dim I As Integer

Dim ArraySettori(500, 1) As Long
Dim ArrayLotti(500, 1) As Long
Dim X As Integer 'Totale serre del settore
Dim Y As Integer 'Totale serre vistate

Dim Valore As Integer
Dim IControl As Integer
Dim ctrl As Control
Dim Link_Settore As Long
Dim ctrlSerre As Control


    For Each ctrl In Me.Controls
        Select Case ctrl.Name
            Case "chkSettore"
                If ctrl.Index > 0 Then
                    Link_Settore = ctrl.Index
                    X = 0
                    Y = 0
                        For Each ctrlSerre In Me.Controls
                            Select Case ctrlSerre.Name
                                Case "chkSerre"
                                    If ctrlSerre.Index > 0 Then
                                        If fnNotNullN(ctrlSerre.Tag) = Link_Settore Then
                                            X = X + 1
                                            If ctrlSerre.Value = vbChecked Then
                                                Y = Y + 1
                                            End If
                                        End If
                                    End If
                            End Select
                        Next
                        
                                
                        If (X - Y) = 0 Then
                            Valore = 1
                        End If
                        If (X - Y) = X Then
                            Valore = 0
                        End If
                        If ((X - Y) > 0) And ((X - Y) <> X) Then
                            Valore = 2
                        End If
                                    
                    ctrl.Value = Valore
                End If
        End Select
    Next
End Sub
