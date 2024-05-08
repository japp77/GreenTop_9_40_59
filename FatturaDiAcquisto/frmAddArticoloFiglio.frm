VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Begin VB.Form frmAddArticoloFiglio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aggiunta articolo per campionatura"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      PropCodice      =   $"frmAddArticoloFiglio.frx":0000
      BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PropDescrizione =   $"frmAddArticoloFiglio.frx":004F
      BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuFunctions   =   $"frmAddArticoloFiglio.frx":00A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin DmtCodDescCtl.DmtCodDesc CDArticoloQuad 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      PropCodice      =   $"frmAddArticoloFiglio.frx":0100
      BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PropDescrizione =   $"frmAddArticoloFiglio.frx":014F
      BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuFunctions   =   $"frmAddArticoloFiglio.frx":01A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmAddArticoloFiglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LINK_LISTINO_ACQ As Long

Private Sub InitControlli()
    'Articolo
    With Me.CDArticolo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = 6 'Articoli
        .CodeIsNumeric = False
    End With
    
    With Me.CDArticoloQuad
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto = " & Link_TipoScarto & ") OR (IDTipoProdotto = " & Link_TipoAumentoPeso & ") OR (IDTipoProdotto = " & Link_TipoCaloPeso & "))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
End Sub

Private Sub cmdConferma_Click()
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim Testo As String


If GET_ESISTENZA_ARTICOLO_DERIVATO(Me.CDArticolo.KeyFieldID, Link_RigaConferimento) = True Then
    Testo = "L'articolo selezionato è gia esistente" & vbCrLf
    Testo = Testo & "Vuoi inserirlo?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Articolo in campionatura") = vbNo Then Exit Sub
End If

sSQL = "SELECT * FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDRV_POCampionaturaRigheTMP = fnGetNewKey("RV_POCampionaturaRigheTMP", "IDRV_POCampionaturaRigheTMP")
    rs!IDUtente = TheApp.IDUser
    rs!IDRV_POCaricoMerceRighe = Link_RigaConferimento
    rs!IDArticolo = Me.CDArticolo.KeyFieldID
    rs!QuantitaCampionata = 0
    rs!QuantitaDefinitiva = 0
    rs!ImportoUnitario = GET_PREZZO_ARTICOLO(LINK_LISTINO_ACQ, Me.CDArticolo.KeyFieldID)
    rs!ImportoNettoRiga = 0
    rs!ImportoImpostaRiga = 0
    rs!ImportoLordoRiga = 0
    rs!Invenduto = 0
    If Me.CDArticoloQuad.KeyFieldID = 0 Then
        rs!DescrizioneAggiuntiva = ""
    Else
        rs!DescrizioneAggiuntiva = "Quadratura inserita manualmente"
    End If
    rs!IDArticoloQuadratura = Me.CDArticoloQuad.KeyFieldID
    rs!CodiceArticoloQuadratura = Me.CDArticoloQuad.Code
    rs!ArticoloQuadratura = Me.CDArticoloQuad.Description
    rs!IDCollegamento = 0
    rs!IDRV_POCampionaturaRighe = 0
    
    GET_INFO_ARTICOLO rs, Me.CDArticolo.KeyFieldID
    
    
rs.Update

rs.Close
Set rs = Nothing

Unload Me

End Sub
Private Function GET_NUMERO_ORDINAMENTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(PesoPerOrdinamento) AS Progressivo FROM RV_POArticoloFiglio "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_ORDINAMENTO = 0
Else
    GET_NUMERO_ORDINAMENTO = fnNotNullN(rs!Progressivo) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ESISTENZA_ARTICOLO_DERIVATO(IDArticolo As Long, IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ARTICOLO_DERIVATO = False
Else
    GET_ESISTENZA_ARTICOLO_DERIVATO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_Load()
    
    InitControlli
    ParametroListinoCampionatura
End Sub

Private Sub GET_INFO_CAMPIONATURA(rsTmp As ADODB.Recordset, IDArticolo As Long, IDRigaConferimento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsTmp!QuantitaCampionata = 0
    rsTmp!QuantitaDefinitiva = 0
    rsTmp!ImportoUnitario = 0 'GET_PREZZO_ARTICOLO(LINK_LISTINO_ACQ, IDArticolo)
    rsTmp!ImportoNettoRiga = 0
    rsTmp!ImportoImpostaRiga = 0
    rsTmp!ImportoLordoRiga = 0
    
Else
    rsTmp!QuantitaCampionata = fnNotNullN(rs!QuantitaCampionata)
    rsTmp!QuantitaDefinitiva = fnNotNullN(rs!QuantitaDefinitiva)
    rsTmp!ImportoUnitario = fnNotNullN(rs!ImportoUnitario)
    rsTmp!ImportoNettoRiga = fnNotNullN(rs!ImportoNettoRiga)
    rsTmp!ImportoImpostaRiga = fnNotNullN(rs!ImportoImpostaRiga)
    rsTmp!ImportoLordoRiga = fnNotNullN(rs!ImportoLordoRiga)

End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_INFO_ARTICOLO(rsTmp As ADODB.Recordset, IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM IERepArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsTmp!CodiceArticolo = ""
    rsTmp!Articolo = ""
    rsTmp!IDUnitaDiMisura = 0
    rsTmp!UnitaDiMisura = ""
    rsTmp!IDIva = 0
    rsTmp!AliquotaIva = 0
Else
    rsTmp!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
    rsTmp!Articolo = fnNotNull(rs!Articolo)
    rsTmp!IDUnitaDiMisura = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    rsTmp!UnitaDiMisura = fnNotNull(rs!UnitaDiMisuraAcquisto)
    rsTmp!IDIva = fnNotNullN(rs!IDIvaAcquisto)
    rsTmp!AliquotaIva = GET_ALIQUOTA_IVA(fnNotNullN(rs!IDIvaAcquisto))
End If

rs.CloseResultset
Set rs = Nothing

End Sub

Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
Else
    GET_ALIQUOTA_IVA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PREZZO_ARTICOLO(IDListino As Long, IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoNettoIVA "
sSQL = sSQL & "FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDListino=" & IDListino
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_ARTICOLO = 0
Else
    GET_PREZZO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub ParametroListinoCampionatura()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoCampionatura FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_LISTINO_ACQ = fnNotNullN(rs!IDListinoCampionatura)
Else
    LINK_LISTINO_ACQ = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

