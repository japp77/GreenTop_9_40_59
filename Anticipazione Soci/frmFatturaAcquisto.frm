VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmFatturaAcquisto 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7470
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
   ScaleHeight     =   4155
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
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
Attribute VB_Name = "frmFatturaAcquisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Riporta As Boolean
Private Sub Form_Load()
    Riporta = False
    fnGriglia
    
End Sub
Private Sub fnGriglia()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT ValoriOggettoPerTipo012D.IDOggetto, ValoriOggettoPerTipo012D.IDTipoOggetto, "
    sSQL = sSQL & "ValoriOggettoPerTipo012D.Doc_data, ValoriOggettoPerTipo012D.Doc_Numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo012D.doc_data_presso_nom, ValoriOggettoPerTipo012D.doc_numero_presso_nom, "
    sSQL = sSQL & "ValoriOggettoPerTipo012D.Link_nom_anagrafica, ValoriOggettoPerTipo012D.Nom_codice, "
    sSQL = sSQL & "ValoriOggettoPerTipo012D.Nom_ragione_sociale_o_cognome "
    sSQL = sSQL & "FROM ValoriOggettoPerTipo012D INNER JOIN "
    sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo012D.IDOggetto = Oggetto.IDOggetto AND ValoriOggettoPerTipo012D.IDTipoOggetto = Oggetto.IDTipoOggetto "
    sSQL = sSQL & " WHERE ValoriOggettoPerTipo012D.Link_nom_anagrafica=" & frmMain.CDSocio.KeyFieldID
    sSQL = sSQL & " AND Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY ValoriOggettoPerTipo012D.Doc_data_presso_nom DESC, ValoriOggettoPerTipo012D.Doc_numero_presso_nom"
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, Cn.InternalConnection
            'Set rsEvent = rsGriglia2.Data
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Doc_data_presso_nom", "Data documento", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Doc_numero_presso_nom", "N° Documento", dgNumeric, True, 1500, dgAlignRight
            .ColumnsHeader.Add "Doc_data", "Data doc. interno", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Doc_numero", "N° Doc. interno", dgNumeric, True, 1500, dgAlignRight
            .ColumnsHeader.Add "Link_nom_anagrafica", "IDAnagraficaFornitore", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Nom_codice", "Codice socio", dgchar, False, 1500, dgAlignleft
            .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Socio", dgchar, False, 3000, dgAlignleft
    
            
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Riporta = True Then
        frmMain.txtIDOggetto.Value = Me.Griglia("IDOggetto").Value
    End If
End Sub
Private Sub Griglia_DblClick()
    Riporta = True
    Unload Me
End Sub

Private Sub Griglia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Riporta = False
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        Griglia_DblClick
    End If
End Sub
