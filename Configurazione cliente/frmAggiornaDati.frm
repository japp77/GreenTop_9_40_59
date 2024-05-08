VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Begin VB.Form frmAggiornaDati 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Aggiornamento dati "
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAggiornaDati.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "ELABORA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri di aggiornamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataInizio 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataFine 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   600
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Data fine"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Data inizio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAggiornaDati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    
    If Me.txtDataInizio.Value = 0 Then
        
        MsgBox "Inserire la data di inizio elaborazione", vbInformation, "Controllo dati"
        Exit Sub
    End If
    
    If Me.txtDataFine.Value = 0 Then
        
        MsgBox "Inserire la data di fine elaborazione", vbInformation, "Controllo dati"
        Exit Sub
    End If
    
    If Me.txtDataInizio.Value > Me.txtDataFine.Value Then
        MsgBox "La data di inizio elaborazione non può essere maggiore della data fine elaborazione", vbInformation, "Controllo dati"
        Exit Sub
    End If

    AGGIORNAMENTO_DATI "ValoriOggettoPerTipo0072", "ValoriOggettoDettaglio0001"
    
    AGGIORNAMENTO_DATI "ValoriOggettoPerTipo0002", "ValoriOggettoDettaglio0004"
    
    AGGIORNAMENTO_DATI "ValoriOggettoPerTipo0008", "ValoriOggettoDettaglio0034"
    
    AGGIORNAMENTO_DATI "ValoriOggettoPerTipo000B", "ValoriOggettoDettaglio0016"
    
    AGGIORNAMENTO_DATI "ValoriOggettoPerTipo006B", "ValoriOggettoDettaglio0007"
    
    lblInfo.Caption = "OPERAZIONE COMPLETA CON SUCCESSO!"
    
    Unload Me
    
End Sub
Private Sub AGGIORNAMENTO_DATI(tabellaTestata As String, tabelleRighe As String)
On Error GoTo ERR_AGGIORNAMENTO_DATI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



lblInfo.Caption = "AGGIORNAMENTO " & tabellaTestata & "..."
DoEvents

Screen.MousePointer = 11

sSQL = "SELECT " & tabelleRighe & ".IDValoriOggettoDettaglio, " & tabelleRighe & ".IDOggetto, " & tabelleRighe & ".IDTipoOggetto, " & tabelleRighe & ".Link_Art_articolo, "
sSQL = sSQL & tabelleRighe & ".Art_codice_lotto, " & tabelleRighe & ".Art_descrizione_lotto, " & tabelleRighe & ".RV_POPrezzoMedioInLiq,"
sSQL = sSQL & tabelleRighe & ".RV_POIDTipoImportoVenditaLiq "
sSQL = sSQL & "FROM " & tabelleRighe & " INNER JOIN "
sSQL = sSQL & tabellaTestata & " ON " & tabelleRighe & ".IDOggetto = " & tabellaTestata & ".IDOggetto AND " & tabelleRighe & ".IDTipoOggetto = " & tabellaTestata & ".IDTipoOggetto"
sSQL = sSQL & " WHERE " & tabellaTestata & ".Doc_data>=" & fnNormDate(Me.txtDataInizio.Text)
sSQL = sSQL & " AND " & tabellaTestata & ".Doc_data<=" & fnNormDate(Me.txtDataFine.Text)
sSQL = sSQL & " AND " & tabelleRighe & ".Link_Art_articolo=" & frmMain.CDArticoloVend.KeyFieldID
sSQL = sSQL & " AND " & tabellaTestata & ".Link_Nom_anagrafica=" & frmMain.CDCliente.KeyFieldID
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    sSQL = "UPDATE " & tabelleRighe & " SET "
    sSQL = sSQL & "RV_POIDTipoImportoVenditaLiq=" & frmMain.cboTipoImportoLiqVend.CurrentID & ", "
    sSQL = sSQL & "RV_POPrezzoMedioInLiq=" & frmMain.chkNonCalcPrezzoMedioVend.Value
    sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rs!IDValoriOggettoDettaglio)
    
    Cn.Execute sSQL
    
    sSQL = "UPDATE Movimento SET "
    sSQL = sSQL & "RV_POIDTipoImportoVenditaLiq=" & frmMain.cboTipoImportoLiqVend.CurrentID & ", "
    sSQL = sSQL & "RV_POPrezzoMedioInLiq=" & frmMain.chkNonCalcPrezzoMedioVend.Value
    sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & fnNotNullN(rs!IDValoriOggettoDettaglio)
    Cn.Execute sSQL
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Screen.MousePointer = 0
DoEvents


Exit Sub
ERR_AGGIORNAMENTO_DATI:
    MsgBox Err.Description, vbCritical, "AGGIORNAMENTO_DATI"
    Screen.MousePointer = 0
    DoEvents

End Sub
