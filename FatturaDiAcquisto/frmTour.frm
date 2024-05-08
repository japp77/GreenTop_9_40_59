VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmTour 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TOUR"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DMTEDITNUMLib.dmtNumber txtIDTourRighe 
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   1200
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   253
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin DMTEDITNUMLib.dmtNumber txtIDTour 
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   253
      BackColor       =   16777215
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEliminaRiferimento 
      Caption         =   "Elimina riferimento"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin DMTEDITNUMLib.dmtNumber txtNumero 
      Height          =   675
      Left            =   1800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   1191
      _StockProps     =   253
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin DMTEDITNUMLib.dmtNumber txtAnno 
      Height          =   675
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1191
      _StockProps     =   253
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin DMTEDITNUMLib.dmtNumber txtPosizione 
      Height          =   675
      Left            =   4920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   1191
      _StockProps     =   253
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Posizione"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Anno"
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
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Numero"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmTour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEliminaRiferimento_Click()
On Error GoTo ERR_cmdEliminaRiferimento_Click
Dim Testo As String
Dim sSQL As String
Dim IDTour As Long
Testo = "Sei sicuro di voler eliminare il riferimento ordine nel tour?"
If MsgBox(Testo, vbYesNo + vbQuestion, "Eliminazione dati") = vbNo Then Exit Sub


Testo = GET_CONTROLLO_BLOCCO_ORDINE_TOUR

If Len(Testo) > 0 Then
    MsgBox Testo, vbCritical, "Eliminazione riferimento tour"
    Exit Sub
End If
Screen.MousePointer = 11

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTourRighe "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & Link_Oggetto
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "UPDATE Oggetto SET "
sSQL = sSQL & "RV_POIDTour=0"
sSQL = sSQL & " WHERE IDOggetto=" & Link_Oggetto
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


IDTour = Me.txtIDTour.Value


GET_CAMBIO_POSIZIONE Me.txtIDTour.Value, Me.txtIDTourRighe.Value, Me.txtPosizione.Value

GET_DATI

'SCRIVI_RICERCA_RIGHE IDTour

Screen.MousePointer = 0
Exit Sub
ERR_cmdEliminaRiferimento_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRiferimento_Click"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    GET_DATI
'
'    Me.Top = frmMain.cmdTour.Top
'    Me.Left = frmMain.cmdTour.Left + frmMain.cmdTour.Width + 20
    
End Sub
Private Sub GET_DATI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT RV_POTour.IDRV_POTour, RV_POTour.Numero, RV_POTour.Anno, RV_POTourRighe.IDRV_POTourRighe, RV_POTourRighe.IDOggettoOrdine, "
sSQL = sSQL & "RV_POTourRighe.Posizione "
sSQL = sSQL & "FROM RV_POTour INNER JOIN "
sSQL = sSQL & "RV_POTourRighe ON RV_POTour.IDRV_POTour = RV_POTourRighe.IDRV_POTour "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & Link_Oggetto

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
    Me.txtAnno.Value = 0
    Me.txtNumero.Value = 0
    Me.txtPosizione.Value = 0
    Me.txtIDTour.Value = 0
    Me.txtIDTourRighe.Value = 0
Else
    Me.txtAnno.Value = fnNotNullN(rs!Anno)
    Me.txtNumero.Value = fnNotNullN(rs!Numero)
    Me.txtPosizione.Value = fnNotNullN(rs!Posizione)
    Me.txtIDTour.Value = fnNotNullN(rs!IDRV_POTour)
    Me.txtIDTourRighe.Value = fnNotNullN(rs!IDRV_POTourRighe)
End If

rs.CloseResultset
Set rs = Nothing


If Me.txtIDTour.Value = 0 Then
    Me.cmdEliminaRiferimento.Enabled = False
Else
    Me.cmdEliminaRiferimento.Enabled = True
End If

End Sub
Private Function GET_CAMBIO_POSIZIONE(IDTour As Long, IDRigaTour As Long, PosizioneOriginale As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_POTourRighe "
'sSQL = sSQL & "WHERE IDRV_POTourRighe<>" & IDRigaTour
sSQL = sSQL & " WHERE Posizione>" & PosizioneOriginale
sSQL = sSQL & " AND IDRV_POTour=" & IDTour
sSQL = sSQL & " ORDER BY Posizione"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    rs!Posizione = rs!Posizione - 1
    rs.Update
rs.MoveNext
Wend
rs.Close
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
Private Function GET_FUNZIONE(IDTipoOggetto) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione "
sSQL = sSQL & "FROM Funzione  "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = 0
Else
    GET_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_BLOCCO_ORDINE_TOUR() As String
Dim LINK_TIPO_OGGETTO_TOUR As Long
Dim LINK_FUNZIONE_TOUR As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_BLOCCO_ORDINE_TOUR = ""

LINK_TIPO_OGGETTO_TOUR = fnGetTipoOggetto("RV_POTour")
LINK_FUNZIONE_TOUR = GET_FUNZIONE(LINK_TIPO_OGGETTO_TOUR)

sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDTour.Value
sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO_TOUR
sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE_TOUR

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_BLOCCO_ORDINE_TOUR = "Il tour risulta aperto dall'utente " & GET_UTENTE(fnNotNullN(rs!IDUtente))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_UTENTE(IDUtente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Utente FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_UTENTE = ""
Else
    GET_UTENTE = fnNotNull(rs!Utente)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub SCRIVI_RICERCA_RIGHE(IDTour As Long)
'On Error GoTo ERR_SCRIVI_RICERCA_RIGHE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'''''VARIABILI DI RICERCA'''''''''''''''''''''''''''''''''''''
Dim RIC_NUMERO_ORDINE As String
Dim RIC_DATA_ORDINE As String
Dim RIC_CLIENTE_ORDINE As String
Dim RIC_DEST_ORDINE As String
Dim RIC_DATA_ARRIVO_DEST_ORDINE As String
Dim RIC_ORA_ARRIVO_DEST_ORDINE As String
Dim RIC_LUOGO_ORDINE As String
Dim RIC_DATA_ARRIVO_LUOGO_ORDINE As String
Dim RIC_ORA_ARRIVO_LUOGO_ORDINE As String
Dim RIC_ARTICOLI_ORDINE As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


RIC_NUMERO_ORDINE = ""
RIC_DATA_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_DEST_ORDINE = ""
RIC_DATA_ARRIVO_DEST_ORDINE = ""
RIC_ORA_ARRIVO_DEST_ORDINE = ""
RIC_LUOGO_ORDINE = ""
RIC_DATA_ARRIVO_LUOGO_ORDINE = ""
RIC_ORA_ARRIVO_LUOGO_ORDINE = ""

''''''''''''''''''''''RICERCA DI TESTA DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIETourRicerca "
sSQL = sSQL & "WHERE IDRV_POTour=" & IDTour

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    While Not rs.EOF
        RIC_NUMERO_ORDINE = RIC_NUMERO_ORDINE & fnNotNull(rs!Doc_Numero) & "|"
        RIC_DATA_ORDINE = RIC_DATA_ORDINE & fnNotNull(rs!Doc_data) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_nome) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_codice) & "|"
        
        RIC_DEST_ORDINE = RIC_DEST_ORDINE & fnNotNull(rs!DestinazioneDiversa) & "|"
        RIC_DATA_ARRIVO_DEST_ORDINE = RIC_DATA_ARRIVO_DEST_ORDINE & fnNotNull(rs!RV_PODataArrivoMerce) & "|"
        RIC_ORA_ARRIVO_DEST_ORDINE = RIC_ORA_ARRIVO_DEST_ORDINE & fnNotNull(rs!RV_POOraArrivoMerce) & "|"
        
        RIC_LUOGO_ORDINE = RIC_LUOGO_ORDINE & fnNotNull(rs!LuogoPresaMerce)
        RIC_DATA_ARRIVO_LUOGO_ORDINE = RIC_DATA_ARRIVO_LUOGO_ORDINE & fnNotNull(rs!RV_PODataArrivoMerceLuogo)
        RIC_ORA_ARRIVO_LUOGO_ORDINE = RIC_ORA_ARRIVO_LUOGO_ORDINE & fnNotNull(rs!RV_POOraArrivoMerceLuogo)
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''RICERCA DI ARTICOLO DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIETourRicercaArt "
sSQL = sSQL & "WHERE IDRV_POTour=" & IDTour

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    While Not rs.EOF

        RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!art_codice) & "|"
        RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!art_descrizione) & "|"
        If fnNotNullN(rs!RV_POTipoRiga) = 1 Then
            RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!RV_POCodiceTipoPedana) & "|"
            RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!RV_PODescrizioneTipoPedana) & "|"
        End If
        
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''AGGIORNAMENTO RICERCA PER TOUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "UPDATE RV_POTour SET "
sSQL = sSQL & "NumeroOrdineRic=" & fnNormString(RIC_NUMERO_ORDINE) & ", "
sSQL = sSQL & "DataOrdineRic=" & fnNormString(RIC_DATA_ORDINE) & ", "
sSQL = sSQL & "ClienteOrdineRic=" & fnNormString(RIC_CLIENTE_ORDINE) & ", "
sSQL = sSQL & "DestinazioneDiversaRic=" & fnNormString(RIC_DEST_ORDINE) & ", "
sSQL = sSQL & "DataArrivoMerceDestRic=" & fnNormString(RIC_DATA_ARRIVO_DEST_ORDINE) & ", "
sSQL = sSQL & "OraArrivoMerceDestRic=" & fnNormString(RIC_ORA_ARRIVO_DEST_ORDINE) & ", "
sSQL = sSQL & "PresaLuogoMerceRic=" & fnNormString(RIC_LUOGO_ORDINE) & ", "
sSQL = sSQL & "DataArrivoLuogoMerceRic=" & fnNormString(RIC_DATA_ARRIVO_LUOGO_ORDINE) & ", "
sSQL = sSQL & "OraArrivoLuogoMerceRic=" & fnNormString(RIC_ORA_ARRIVO_LUOGO_ORDINE) & ", "
sSQL = sSQL & "ArticoloOrdineRic=" & fnNormString(RIC_ARTICOLI_ORDINE)
sSQL = sSQL & " WHERE IDRV_POTour=" & IDTour

Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Exit Sub
ERR_SCRIVI_RICERCA_RIGHE:
    MsgBox Err.Description, vbCritical, "Funzionalità per ricerca righe"
    
End Sub


