VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmAltreOperazioni 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Altre Operazioni"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAltreOperazioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Altre operazioni"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   11055
      Begin VB.CommandButton Command4 
         Caption         =   "BLOCCO Q.TA CONFERITA LAVORATA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         TabIndex        =   13
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SEZIONALI PER C.M.R."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdAggiornaReportEtichette 
         Caption         =   "AGGIORNA REPORT ETICHETTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdUtentiSbloccoLotto 
         Caption         =   "UTENTI SBLOCCO LOTTO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DISATTIVA FORMULA QUANTITA'"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RIPRISTINA VELOCITA' LISTA PEDANE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operazioni utenti bloccate"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11055
      Begin VB.CommandButton cmdEliminaUtente 
         Caption         =   "Elimina"
         Height          =   375
         Left            =   9120
         TabIndex        =   7
         Top             =   2760
         Width           =   1815
      End
      Begin DmtGridCtl.DmtGrid Griglia 
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4260
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
   Begin VB.Frame frame1 
      Caption         =   "Ripristino tabella filtro "
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.OptionButton optFiltroTutti 
         Caption         =   "Elimina tutti i riferimenti"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin VB.OptionButton optFiltroSoloUtente 
         Caption         =   "Elimina riferimenti dell'utente loggato"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.CommandButton cmdFiltroTabella 
         Caption         =   "CONFERMA OPERAZIONE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAltreOperazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub cmdAggiornaReportEtichette_Click()
On Error GoTo ERR_cmdAggiornaReportEtichette_Click
Dim IDTipoOggettoEtichette As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset


'If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

Screen.MousePointer = 11

IDTipoOggettoEtichette = fnGetTipoOggetto("RV_POEtichetteLavorazione")

Set rsNew = New ADODB.Recordset

rsNew.Open "SELECT * FROM RV_POEtichetteDefault", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsNew.EOF
    rsNew!IDReportPerTipoOggetto = GET_LINK_REPORT(IDTipoOggettoEtichette, fnNotNull(rsNew!ReportPerTipoOggetto))
    rsNew.Update
rsNew.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

Screen.MousePointer = 0
Exit Sub
ERR_cmdAggiornaReportEtichette_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdAggiornaReportEtichette_Click"
    Screen.MousePointer = 0
End Sub
Private Function GET_LINK_REPORT(IDTipoOggetto As Long, NomeReport As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND ReportTipoOggetto=" & fnNormString(NomeReport)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REPORT = 0
Else
    GET_LINK_REPORT = fnNotNullN(rs!IDReportTipoOggetto)
End If



rs.CloseResultset
Set rs = Nothing


End Function
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
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub cmdEliminaUtente_Click()
On Error GoTo ERR_cmdEliminaUtente_Click
Dim sSQL As String
Dim Testo As String

If TheApp.IDUser <> fnNotNullN(Me.Griglia.AllColumns("IDUtente")) Then
    
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Eliminando l'utente dalla coda di salvataggio si potrebbero avere degli effetti indesiderati." & vbCrLf
    Testo = Testo & "Continuare con il comando di eliminazione?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione dati di coda") = vbYes Then
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & Me.Griglia("IDUtente").Value
        'sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto
        Cn.Execute sSQL
        
        SettaggioGriglia
        
    End If
End If

Exit Sub
ERR_cmdEliminaUtente_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaUtente_Click"
End Sub

Private Sub cmdFiltroTabella_Click()
On Error GoTo ERR_cmdFiltroTabella_Click
Dim Testo As String
Dim sSQL As String

If (Me.optFiltroSoloUtente.Value = True) Then
    Testo = "Sei sicuro di procedere con questo comando?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub

    sSQL = "DELETE FROM RV_POFiltro "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    
    Cn.Execute sSQL
End If

If (Me.optFiltroTutti.Value = True) Then
    Testo = "Sei sicuro di procedere con questo comando?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub

    sSQL = "DELETE FROM RV_POFiltro "
    
    Cn.Execute sSQL
End If

MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO", vbInformation, "Cancellazione dati"

Exit Sub
ERR_cmdFiltroTabella_Click:
    MsgBox Err.Description, vbCritical, "cmdFiltroTabella_Click"


End Sub

Private Sub cmdUtentiSbloccoLotto_Click()
On Error GoTo ERR_cmdUtentiSbloccoLotto_Click
    
    frmSbloccoLotto.Show vbModal

Exit Sub
ERR_cmdUtentiSbloccoLotto_Click:
    MsgBox Err.Description, vbCritical, "cmdUtentiSbloccoLotto_Click"
End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
Dim sSQL As String

'ELIMINAZIONE INDICE
sSQL = "DROP INDEX " & "RV_POAnno" & " ON " & "RV_POPedana"
Cn.Execute sSQL

'ELIMINAZIONE INDICE
sSQL = "DROP INDEX " & "RV_POIDAzienda" & " ON " & "RV_POPedana"
Cn.Execute sSQL

'ELIMINAZIONE INDICE
sSQL = "DROP INDEX " & "RV_POIDFiliale" & " ON " & "RV_POPedana"
Cn.Execute sSQL


'CREAZIONE INDICE
sSQL = "CREATE INDEX " & "RV_POAnno" & " ON " & "RV_POPedana"
sSQL = sSQL & " (" & "Anno" & ")"
Cn.Execute sSQL

'CREAZIONE INDICE
sSQL = "CREATE INDEX " & "RV_POIDAzienda" & " ON " & "RV_POPedana"
sSQL = sSQL & " (" & "IDAzienda" & ")"
Cn.Execute sSQL

'CREAZIONE INDICE
sSQL = "CREATE INDEX " & "RV_POIDFiliale" & " ON " & "RV_POPedana"
sSQL = sSQL & " (" & "IDFiliale" & ")"
Cn.Execute sSQL

MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO", vbInformation, "Rispristino indici lista pedane"

Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Command1_Click"
    
    MsgBox "Se il problema persiste contattare immediatamente l'assistenza", vbExclamation, "Ripristino indici lista pedane"
End Sub

Private Sub Command2_Click()
On Error GoTo ERR_DISATTIVA_FORMULA

    ELIMINA_FORMULA_QUANTITA

Exit Sub
ERR_DISATTIVA_FORMULA:
MsgBox Err.Description, vbCritical, "DISATTIVA_FORMULA"
End Sub

Private Sub Command3_Click()
    frmSezPerCMR.Show vbModal
    
End Sub

Private Sub Command4_Click()
On Error GoTo ERR_cmdUtentiSbloccoLotto_Click
    
    frmBloccoQtaConf.Show vbModal

Exit Sub
ERR_cmdUtentiSbloccoLotto_Click:
    MsgBox Err.Description, vbCritical, "cmdUtentiSbloccoLotto_Click"
End Sub

Private Sub Form_Load()
SettaggioGriglia
End Sub

Private Sub optFiltroSoloUtente_Click()
    If optFiltroSoloUtente.Value = True Then optFiltroTutti.Value = False
End Sub

Private Sub optFiltroTutti_Click()
    If optFiltroTutti.Value = True Then optFiltroSoloUtente.Value = False

End Sub
Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_POTMP "
    'sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto
    sSQL = sSQL & "ORDER BY IDSessione, IDUtente "
    
    Set rsArt = Cn.OpenResultset(sSQL)
        Set rsEvent = rsArt.Data
    
    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDSessione", "Sessione", dgNumeric, True, 1000, dgAlignleft
            .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 3000, dgAlignleft
        Set .Recordset = rsArt.Data
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia operazioni utente"
End Sub

Private Sub ELIMINA_FORMULA_QUANTITA()
On Error GoTo ERR_ELIMINA_FORMULA_QUANTITA
Dim sSQL As String

sSQL = "UPDATE CampoDiamante SET "
sSQL = sSQL & "Formula2=NULL "
sSQL = sSQL & "WHERE CampoDiamante=" & fnNormString("Art_quantita_totale")

Cn.Execute sSQL

MsgBox "Formula quantità disattivata con successo", vbInformation, "Disattivazione formula quantità"

Exit Sub
ERR_ELIMINA_FORMULA_QUANTITA:
    MsgBox Err.Description, vbCritical, "ELIMINA_FORMULA_QUANTITA"
End Sub
