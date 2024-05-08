VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmTrovaOrdine 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TROVA ORDINE"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GridOrdine 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   9255
      _ExtentX        =   16325
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
   Begin VB.Frame Frame1 
      Caption         =   "Parametri"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9255
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "RICERCA"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin DmtCodDescCtl.DmtCodDesc cdCliente 
         Height          =   615
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         PropCodice      =   $"frmTrovaOrdine.frx":0000
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmTrovaOrdine.frx":004E
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmTrovaOrdine.frx":00A0
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
      Begin DMTDATETIMELib.dmtDate txtDataOrdine 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
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
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
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
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Data ordine"
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
         Left            =   5520
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Numero ordine"
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
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ORDINE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   9255
   End
End
Attribute VB_Name = "frmTrovaOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsGriglia As DmtOleDbLib.adoResultset

Private Sub cmdRicerca_Click()
    
    fncGriglia
    Me.GridOrdine.SetFocus

End Sub

Private Sub Form_Activate()
    
    Me.cdCliente.Load frmMain.cdCliente.KeyFieldID
    'Me.txtNumeroOrdine.Value = frmMain.txtNumeroOrdine.Value
    'M 'e.txtDataOrdine.Value = frmMain.txtDataOrdine.Value
    
    fncGriglia

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
     With Me.cdCliente
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Clienti"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    
End Sub
Private Sub fncGriglia()
    Dim sSQL As String
    Dim sSQL_WHERE As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    Dim i As Integer
    
    sSQL_WHERE = ""
    sSQL = "SELECT ValoriOggettoPerTipo000F.IDOggetto, Doc_data, Doc_numero, Nom_ragione_sociale_o_cognome, Link_nom_anagrafica, "
    sSQL = sSQL & "Link_Nom_ult_sito, SitoPerAnagrafica, Doc_annotazioni_variazio, Doc_data_presso_nom, Doc_numero_presso_nom, Doc_data_prevista_evasione, "
    sSQL = sSQL & "RV_POAnnotazioniInterna, RV_PODescrizioneCorpoDocEv, RV_POIDLuogoPresaMerce, RV_POIDTrasportatoreSuccessivo "
    sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto AND "
    sSQL = sSQL & "ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto LEFT OUTER JOIN "
    sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo000F.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica "
    sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
    'sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    
    
    If Me.cdCliente.KeyFieldID > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Link_nom_anagrafica=" & Me.cdCliente.KeyFieldID
    End If
    
    If Me.txtNumeroOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_numero=" & Me.txtNumeroOrdine.Value
    End If
    
    If Me.txtDataOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_data=" & fnNormDate(Me.txtDataOrdine.Value)
    End If
    
    sSQL = sSQL & sSQL_WHERE & " ORDER BY Doc_data DESC, Doc_numero DESC"
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        Set rsGriglia = Cn.OpenResultset(sSQL)
            Set rsEvent = rsGriglia.Data
        
        With Me.GridOrdine
            
            .ColumnsHeader.Clear
            '.ColumnsHeader.Add "StatoOrdine", "Stato ordine", dgchar, True, 2000, dgAlignleft, True, , True
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Link_nom_anagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Doc_data", "Data ordine", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Doc_numero", "Numero ordine", dgNumeric, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 5000, dgAlignleft
            .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione diversa", dgchar, True, 5000, dgAlignleft
            .ColumnsHeader.Add "Doc_data_presso_nom", "Data ordine cliente", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Doc_numero_presso_nom", "Numero ordine cliente", dgchar, True, 2000, dgAlignleft
            Set .Recordset = rsGriglia.Data
            .Refresh
        End With
        
        Cn.CursorLocation = OLDCursor



End Sub



Private Sub Form_Unload(Cancel As Integer)
    rsGriglia.CloseResultset
    Set rsGriglia = Nothing
End Sub

Private Sub GridOrdine_DblClick()
Dim IDUtenteBlocco As Long
Dim Testo As String

    frmMain.txtIDOrdine.Value = fnNotNullN(Me.GridOrdine.AllColumns("IDOggetto").Value)
    
    Unload Me

End Sub
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
Private Function GET_STATO_ORDINE(IDOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_STATO_ORDINE = 0
Else
    GET_STATO_ORDINE = 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
