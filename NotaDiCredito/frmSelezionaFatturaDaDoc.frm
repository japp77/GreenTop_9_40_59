VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaFatturaDaDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona fattura da collegare alla nota di credito"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelezionaFatturaDaDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid DmtGrid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12515
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
Attribute VB_Name = "frmSelezionaFatturaDaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Conferma As Boolean

Private Sub DmtGrid1_DblClick()
    Conferma = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Conferma = True
    Unload Me
End If
If KeyCode = vbKeyEscape Then
    Conferma = False
    Unload Me
End If


End Sub

Private Sub Form_Load()
    
    Conferma = False
    GET_GRIGLIA
    
End Sub
Private Sub GET_GRIGLIA()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    If frmMain.cboTipoOggettoColl.CurrentID = 114 Then
        sSQL = "SELECT ValoriOggettoPerTipo0072.IDOggetto, ValoriOggettoPerTipo0072.IDTipoOggetto, "
        sSQL = sSQL & "ValoriOggettoPerTipo0072.Doc_numero, ValoriOggettoPerTipo0072.doc_data, Sezionale.Sezionale, "
        sSQL = sSQL & "SitoPerAnagrafica.SitoPerAnagrafica,Sezionale.Prefisso "
        sSQL = sSQL & "FROM ValoriOggettoPerTipo0072 INNER JOIN "
        sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo0072.IDOggetto = Oggetto.IDOggetto AND "
        sSQL = sSQL & "ValoriOggettoPerTipo0072.IDTipoOggetto = Oggetto.IDTipoOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo0072.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica LEFT OUTER JOIN "
        sSQL = sSQL & "Sezionale ON ValoriOggettoPerTipo0072.Link_Doc_sezionale = Sezionale.IDSezionale "
    End If
    
    
    If frmMain.cboTipoOggettoColl.CurrentID = 4 Then
        sSQL = "SELECT ValoriOggettoPerTipo0004.IDOggetto, ValoriOggettoPerTipo0004.IDTipoOggetto, "
        sSQL = sSQL & "ValoriOggettoPerTipo0004.Doc_numero, ValoriOggettoPerTipo0004.doc_data, Sezionale.Sezionale, "
        sSQL = sSQL & "SitoPerAnagrafica.SitoPerAnagrafica, Sezionale.Prefisso "
        sSQL = sSQL & "FROM ValoriOggettoPerTipo0004 INNER JOIN "
        sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo0004.IDOggetto = Oggetto.IDOggetto AND "
        sSQL = sSQL & "ValoriOggettoPerTipo0004.IDTipoOggetto = Oggetto.IDTipoOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "SitoPerAnagrafica ON ValoriOggettoPerTipo0004.Link_Nom_ult_sito = SitoPerAnagrafica.IDSitoPerAnagrafica LEFT OUTER JOIN "
        sSQL = sSQL & "Sezionale ON ValoriOggettoPerTipo0004.Link_Doc_sezionale = Sezionale.IDSezionale "
    End If

    sSQL = sSQL & "WHERE Link_nom_anagrafica=" & frmMain.cdAnagrafica.KeyFieldID
    sSQL = sSQL & " AND Oggetto.IDAzienda=" & TheApp.IDFirm
    
    If frmMain.cboAltroSito.CurrentID > 0 Then
        sSQL = sSQL & " AND Link_Nom_ult_sito=" & frmMain.cboAltroSito.CurrentID
    End If
    
    sSQL = sSQL & " ORDER BY Doc_data desc, Doc_numero desc"
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
    
        With Me.DmtGrid1
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            
            .ColumnsHeader.Add "Doc_data", "Data doc.", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Doc_numero", "N° doc.", dgInteger, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Prefisso", "Prefisso", dgchar, True, 1000, dgAlignleft
            .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione", dgchar, True, 2500, dgAlignleft
            
            Set .Recordset = rsGriglia
            .LoadUserSettings
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Conferma = True Then
        frmMain.txtIDOggettoCollegato.Value = Me.DmtGrid1.AllColumns("IDOggetto").Value
    End If
End Sub

