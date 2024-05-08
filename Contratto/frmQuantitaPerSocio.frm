VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmQuantitaPerSocio 
   Caption         =   "Quantità contratto per socio"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuantitaPerSocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "TOTALI DA CONTRATTO"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   11655
      Begin DMTEDITNUMLib.dmtNumber txtQtaConferitaContratto 
         Height          =   570
         Left            =   6840
         TabIndex        =   11
         Top             =   525
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1005
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777152
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
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaContratto 
         Height          =   570
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   1005
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777152
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
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà da Contratto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà da certificato/conferito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   12
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Elimina"
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuovo"
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salva"
      Height          =   495
      Left            =   11880
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11655
      Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1085
         PropCodice      =   $"frmQuantitaPerSocio.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmQuantitaPerSocio.frx":47868
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmQuantitaPerSocio.frx":478CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantita 
         Height          =   330
         Left            =   7080
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantitaConferita 
         Height          =   330
         Left            =   9360
         TabIndex        =   8
         Top             =   360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà da certificato"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Quantità contratto"
         Height          =   255
         Index           =   6
         Left            =   7080
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
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
Attribute VB_Name = "frmQuantitaPerSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private NuovoRecord As Long

Private Sub CDSocioFatt_ChangeElement()
    Me.txtQuantitaConferita.Value = CALCOLA_TOTALE_CERTIFICATO(Me.CDSocioFatt.KeyFieldID)
End Sub

Private Sub Command1_Click()
    Salva
        
    
End Sub

Private Sub Command2_Click()
    Nuovo
    
End Sub

Private Sub Command3_Click()
    Elimina
End Sub

Private Sub dmtNumber1_Change()

End Sub

Private Sub Form_Load()
    INIT_CONTROLLI
    
    GET_GRIGLIA
    ChangeDatiQtaPerSocio = 0
    
    Me.txtQtaContratto.Value = CALCOLA_TOTALE_CONTRATTO
    Me.txtQtaConferitaContratto.Value = CALCOLA_TOTALE_CERTIFICATO_CONTRATTO
    
End Sub
Private Sub INIT_CONTROLLI()
    'Cooperativa
    With Me.CDSocioFatt
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "DenominazioneCompleta"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaCooperativaDaLibroSoci"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Denominazione Completa"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Denominazione completa"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * FROM RV_POIETMPQuantitaContrattoPerSocio "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
'sSQL = sSQL & " AND ID_Art_dettaglio_prog=" & IDArtDettaglioProgSEL
sSQL = sSQL & " ORDER BY Anagrafica"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection

With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignRight
    '.ColumnsHeader.Add "ID_Art_dettaglio_prog", "ID_Art_dettaglio_prog", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAnagraficaSocio", "IDAnagraficaSocio", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Anagrafica", "Cooperativa/Socio diretto", dgchar, True, 5500, dgAlignleft
    .ColumnsHeader.Add "Quantita", "Q.tà da contratto", dgDouble, True, 2500, dgAlignRight
    .ColumnsHeader.Add "QuantitaConferita", "Q.tà da certificato", dgDouble, True, 2500, dgAlignRight
    
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub GrigliaCorpo_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    Me.CDSocioFatt.Load fnNotNullN(Me.GrigliaCorpo.AllColumns("IDAnagraficaSocio").Value)
    Me.txtQuantita.Value = fnNotNullN(Me.GrigliaCorpo.AllColumns("Quantita").Value)
    
    NuovoRecord = 0
End Sub
Private Sub Salva()
On Error GoTo ERR_Salva
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim IDSocioSel As Long


If (Me.CDSocioFatt.KeyFieldID = 0) Then
    MsgBox "Inserire la cooperativa/socio diretto", vbCritical, "Controllo dati"
    Exit Sub
End If

IDSocioSel = Me.CDSocioFatt.KeyFieldID

sSQL = "SELECT * "
sSQL = sSQL & " FROM RV_POTMPQuantitaContrattoPerSocio "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
'sSQL = sSQL & " AND ID_Art_dettaglio_prog=" & IDArtDettaglioProgSEL
sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocioFatt.KeyFieldID

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
End If
rs!IDUtente = TheApp.IDUser
'rs!ID_Art_dettaglio_prog = IDArtDettaglioProgSEL
rs!IDOggetto = oDoc.IDOggetto
rs!IDAnagraficaSocio = Me.CDSocioFatt.KeyFieldID
rs!Quantita = Me.txtQuantita.Value
rs!QuantitaConferita = CALCOLA_TOTALE_CERTIFICATO(rs!IDAnagraficaSocio)
rs.Update

rs.Close
Set rs = Nothing

GET_GRIGLIA

SetIndexGriglia IDSocioSel

ChangeDatiQtaPerSocio = 1



Exit Sub
ERR_Salva:
    MsgBox Err.Description, vbCritical, "Salva"
End Sub
Private Sub Elimina()
On Error GoTo ERR_Elimina
Dim sSQL As String

If MsgBox("Vuoi eliminare l'elemento selezionato?", vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub

sSQL = "DELETE "
sSQL = sSQL & " FROM RV_POTMPQuantitaContrattoPerSocio "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
'sSQL = sSQL & " AND ID_Art_dettaglio_prog=" & IDArtDettaglioProgSEL
sSQL = sSQL & " AND IDAnagraficaSocio=" & Me.CDSocioFatt.KeyFieldID

Cn.Execute sSQL

GET_GRIGLIA

Nuovo

ChangeDatiQtaPerSocio = 1
Exit Sub
ERR_Elimina:
    MsgBox Err.Description, vbCritical, "Elimina"
End Sub
Private Sub Nuovo()
Me.CDSocioFatt.Load 0
Me.txtQuantita.Value = 0
Me.txtQuantitaConferita.Value = 0

NuovoRecord = 1

Me.CDSocioFatt.SetFocus
End Sub
Private Function CALCOLA_TOTALE_CERTIFICATO(IDSocio As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ID As Long

CALCOLA_TOTALE_CERTIFICATO = 0

sSQL = "SELECT SUM(PesoNettoCalcolato) AS Totale "
sSQL = sSQL & " FROM RV_POCertificato "
sSQL = sSQL & " WHERE IDContratto=" & oDoc.IDOggetto
If (IDSocio > 0) Then
    sSQL = sSQL & " AND IDAnagraficaCooperativa=" & IDSocio
End If

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    CALCOLA_TOTALE_CERTIFICATO = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function CALCOLA_TOTALE_CERTIFICATO_CONTRATTO() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ID As Long

CALCOLA_TOTALE_CERTIFICATO_CONTRATTO = 0

sSQL = "SELECT SUM(PesoNettoCalcolato) AS Totale "
sSQL = sSQL & " FROM RV_POCertificato "
sSQL = sSQL & " WHERE IDContratto=" & oDoc.IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    CALCOLA_TOTALE_CERTIFICATO_CONTRATTO = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function CALCOLA_TOTALE_CONTRATTO() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ID As Long

CALCOLA_TOTALE_CONTRATTO = 0

sSQL = "SELECT SUM(Art_quantita_totale) AS Totale "
sSQL = sSQL & " FROM ValoriOggettoDettaglio0038 "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    CALCOLA_TOTALE_CONTRATTO = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub SetIndexGriglia(IDSocio As Long)

While Not GrigliaCorpo.Recordset.EOF
    If ((Me.GrigliaCorpo.AllColumns("IDAnagraficaSocio").Value = IDSocio)) Then Exit Sub
        
GrigliaCorpo.Recordset.MoveNext
Wend

End Sub
