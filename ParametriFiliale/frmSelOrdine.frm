VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelOrdine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleziona ordine"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10965
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
   ScaleHeight     =   7260
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GridOrdine 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12726
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
Attribute VB_Name = "frmSelOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsGriglia As DmtOleDbLib.adoResultset


Private Sub Form_Activate()

fncGriglia


End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

End Sub
Private Sub fncGriglia()
    Dim sSQL As String
    Dim sSQL_WHERE As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    sSQL_WHERE = ""
    
    sSQL = "SELECT * FROM RV_POIERicercaOrdine "
    sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    
    sSQL = sSQL & sSQL_WHERE & " ORDER BY RV_PODataOrdinePadre DESC, RV_PONumeroOrdinePadre DESC, RV_PONumeroListaPrelievo DESC"

    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    Set rsGriglia = Cn.OpenResultset(sSQL)
        Set rsEvent = rsGriglia.Data

    With Me.GridOrdine
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Link_nom_anagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POIDTipoOrdine", "RV_POIDTipoOrdine", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoOrdine", "Tipo ordine", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "RV_PONumeroOrdinePadre", "Numero ordine", dgNumeric, True, 2000, dgAlignRight
            .ColumnsHeader.Add "RV_PODataOrdinePadre", "Data ordine", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_PONumeroListaPrelievo", "N° lista prelievo", dgNumeric, True, 2000, dgAlignRight
            .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 5000, dgAlignleft
            .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione diversa", dgchar, True, 5000, dgAlignleft
            .ColumnsHeader.Add "Doc_data_prevista_evasione", "Data partenza", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Doc_data_presso_nom", "Data ordine cliente", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Doc_numero_presso_nom", "Numero ordine cliente", dgchar, True, 2000, dgAlignleft
                        
            .ColumnsHeader.Add "RV_POIDOrdinePadre", "IDOggettoPadre", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Doc_data", "Data ordine originale", dgDate, False, 2000, dgAlignleft
            .ColumnsHeader.Add "Doc_numero", "Numero ordine originale", dgNumeric, False, 2000, dgAlignleft
            .ColumnsHeader.Add "Link_Doc_sezionale", "IDSezionale", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, False, 2500, dgAlignleft
            
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
    ControlIDOrdine.Value = fnNotNullN(Me.GridOrdine("IDOggetto").Value)
    Unload Me
End Sub

