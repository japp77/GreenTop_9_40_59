VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form frmRiscontroPeso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Riscontro peso"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   12000
      TabIndex        =   0
      Top             =   5640
      Width           =   2295
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9763
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
   Begin DmtSearchAccount.DmtSearchACS ACSSocio 
      Height          =   585
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1032
      WidthDescription=   3800
      WidthSecondDescription=   1500
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SearchType      =   2
      HideLeaf        =   0   'False
      BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionDescription=   "Socio/Fornitore"
      CaptionCode     =   "Codice"
      OnlyAccounts    =   -1  'True
   End
End
Attribute VB_Name = "frmRiscontroPeso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private IDSocioPredefinito As Long

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim IDRiga  As Integer

Set rsGrigliaRP = New ADODB.Recordset
Set rsGrigliaRP1 = New ADODB.Recordset

rsGrigliaRP.CursorLocation = adUseClient
rsGrigliaRP1.CursorLocation = adUseClient


rsGrigliaRP.Fields.Append "IDRiga", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "IDArticoloImballo", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "CodiceArticoloImballo", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "ArticoloImballo", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "ImportoUnitario", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "IDIva", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "CodiceIva", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "AliquotaIva", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "Iva", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "Sconto1", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "Sconto2", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "ImportoUnitarioScontato", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "IDUnitaDiMisura", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "UnitaDiMisura", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "IDUnitaDiMisuraCoop", adInteger, , adFldIsNullable
rsGrigliaRP.Fields.Append "UnitaDiMisuraCoop", adVarChar, 250, adFldIsNullable
rsGrigliaRP.Fields.Append "Colli", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "QuantitaRiscontrata", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "Differenza", adDouble, , adFldIsNullable
rsGrigliaRP.Fields.Append "ImportoInclusoImballo", adBoolean, , adFldIsNullable


rsGrigliaRP1.Fields.Append "IDRiga", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "IDArticoloImballo", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "CodiceArticoloImballo", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "ArticoloImballo", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "ImportoUnitario", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "IDIva", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "CodiceIva", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "AliquotaIva", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "Iva", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "Sconto1", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "Sconto2", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "ImportoUnitarioScontato", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "IDUnitaDiMisura", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "UnitaDiMisura", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "IDUnitaDiMisuraCoop", adInteger, , adFldIsNullable
rsGrigliaRP1.Fields.Append "UnitaDiMisuraCoop", adVarChar, 250, adFldIsNullable
rsGrigliaRP1.Fields.Append "Colli", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "QuantitaRiscontrata", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "Differenza", adDouble, , adFldIsNullable
rsGrigliaRP1.Fields.Append "ImportoInclusoImballo", adBoolean, , adFldIsNullable

rsGrigliaRP.Open , , adOpenKeyset, adLockBatchOptimistic
rsGrigliaRP1.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT " & sTabellaDettaglio & ".Link_Art_articolo, " & sTabellaDettaglio & ".Art_codice, " & sTabellaDettaglio & ".Art_descrizione, " & sTabellaDettaglio & ".Link_Art_IVA, "
sSQL = sSQL & sTabellaDettaglio & ".Art_prezzo_unitario_neutro, " & sTabellaDettaglio & ".Art_sco_in_percentuale_1, " & sTabellaDettaglio & ".Art_sco_in_percentuale_2, "
sSQL = sSQL & "SUM(" & sTabellaDettaglio & ".Art_numero_colli) AS NumeroColli, " & sTabellaDettaglio & ".Link_Art_unita_di_misura, "
sSQL = sSQL & "UnitaDiMisura.DescrizioneFattura , UnitaDiMisura.RV_POIDUnitaDiMisuraCoop, RV_POUnitaDiMisuraCoop.UnitaDiMisuraCoop, Sum(" & sTabellaDettaglio & ".Art_quantita_totale) "
sSQL = sSQL & "AS QuantitaTotale, " & sTabellaDettaglio & ".RV_POIDImballo, " & sTabellaDettaglio & ".RV_POCodiceImballo, " & sTabellaDettaglio & ".RV_PODescrizioneImballo, Iva.Iva, Iva.AliquotaIva,"
sSQL = sSQL & "Iva.Codice, " & sTabellaDettaglio & ".RV_POImportoImballoInArticolo "
sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
sSQL = sSQL & sTabellaDettaglio & " ON Iva.IDIva = " & sTabellaDettaglio & ".Link_Art_IVA LEFT OUTER JOIN "
sSQL = sSQL & "RV_POUnitaDiMisuraCoop RIGHT OUTER JOIN "
sSQL = sSQL & "UnitaDiMisura ON RV_POUnitaDiMisuraCoop.IDRV_POUnitaDiMisuraCoop = UnitaDiMisura.RV_POIDUnitaDiMisuraCoop ON "
sSQL = sSQL & sTabellaDettaglio & ".Link_Art_unita_di_misura = UnitaDiMisura.IDUnitaDiMisura"
sSQL = sSQL & " WHERE " & sTabellaDettaglio & ".IDOggetto = " & oDoc.IDOggetto
sSQL = sSQL & " AND " & sTabellaDettaglio & ".RV_POTipoRiga = 1 "
sSQL = sSQL & " AND " & sTabellaDettaglio & ".IDTipoOggetto = " & oDoc.IDTipoOggetto


sSQL = sSQL & " GROUP BY " & sTabellaDettaglio & ".Link_Art_articolo, " & sTabellaDettaglio & ".Art_codice, " & sTabellaDettaglio & ".Art_descrizione, " & sTabellaDettaglio & ".Link_Art_IVA,"
sSQL = sSQL & sTabellaDettaglio & ".Art_prezzo_unitario_neutro, " & sTabellaDettaglio & ".Art_sco_in_percentuale_1, " & sTabellaDettaglio & ".Art_sco_in_percentuale_2, "
sSQL = sSQL & "ValoriOggettoDettaglio0004.Link_Art_unita_di_misura, UnitaDiMisura.DescrizioneFattura, UnitaDiMisura.RV_POIDUnitaDiMisuraCoop,"
sSQL = sSQL & "RV_POUnitaDiMisuraCoop.UnitaDiMisuraCoop, " & sTabellaDettaglio & ".RV_POIDImballo, " & sTabellaDettaglio & ".RV_POCodiceImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODescrizioneImballo, Iva.Iva, Iva.AliquotaIva, Iva.Codice, " & sTabellaDettaglio & ".RV_POImportoImballoInArticolo "


Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

IDRiga = 1
While Not rs.EOF
    rsGrigliaRP.AddNew
        rsGrigliaRP!IDRiga = IDRiga
        rsGrigliaRP!IDArticolo = fnNotNullN(rs!Link_Art_articolo)
        rsGrigliaRP!CodiceArticolo = fnNotNull(rs!art_codice)
        rsGrigliaRP!Articolo = fnNotNull(rs!art_descrizione)
        rsGrigliaRP!IDArticoloImballo = fnNotNullN(rs!RV_POIDImballo)
        rsGrigliaRP!CodiceArticoloImballo = fnNotNull(rs!RV_POCodiceImballo)
        rsGrigliaRP!ArticoloImballo = fnNotNull(rs!RV_PODescrizioneImballo)
        rsGrigliaRP!ImportoUnitario = fnNotNullN(rs!Art_prezzo_unitario_neutro)
        rsGrigliaRP!IDIva = fnNotNullN(rs!Link_Art_IVA)
        rsGrigliaRP!CodiceIva = fnNotNull(rs!Codice)
        rsGrigliaRP!AliquotaIva = fnNotNullN(rs!AliquotaIva)
        rsGrigliaRP!Iva = fnNotNull(rs!Iva)
        rsGrigliaRP!Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
        rsGrigliaRP!Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
'        rsGrigliaRP!ImportoUnitarioScontato = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
        rsGrigliaRP!IDUnitaDiMisura = fnNotNullN(rs!Link_Art_unita_di_misura)
        rsGrigliaRP!UnitaDiMisura = fnNotNull(rs!DescrizioneFattura)
        rsGrigliaRP!IDUnitaDiMisuraCoop = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
        rsGrigliaRP!UnitaDiMisuraCoop = fnNotNull(rs!UnitaDiMisuraCoop)
        rsGrigliaRP!Colli = fnNotNullN(rs!NumeroColli)
        rsGrigliaRP!Quantita = fnNotNullN(rs!QuantitaTotale)
        rsGrigliaRP!QuantitaRiscontrata = fnNotNullN(rs!QuantitaTotale)
        rsGrigliaRP!Differenza = rsGrigliaRP!QuantitaRiscontrata - rsGrigliaRP!Quantita
        rsGrigliaRP!ImportoInclusoImballo = fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1
    rsGrigliaRP.Update
    IDRiga = IDRiga + 1
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim I As Integer
    If RISCONTRO_PESO_PER_CONF = 0 Then
        If Me.ACSSocio.IDAnagrafica = 0 Then
            MsgBox "Inserire il socio di riferimento", vbCritical, "Riscontro peso"
            Exit Sub
        End If
    End If
    LINK_SOCIO_RISCONTRO_PESO = Me.ACSSocio.IDAnagrafica
    
    If Not ((rsGrigliaRP.EOF) And (rsGrigliaRP.BOF)) Then
        rsGrigliaRP.MoveFirst
        While Not rsGrigliaRP.EOF
            rsGrigliaRP1.AddNew
                For I = 0 To rsGrigliaRP.Fields.Count - 1
                    rsGrigliaRP1.Fields(rsGrigliaRP.Fields(I).Name).Value = rsGrigliaRP.Fields(I).Value
                Next
            rsGrigliaRP1.Update
        rsGrigliaRP.MoveNext
        Wend
        
        RiscontroPeso = True
    End If
    
    Unload Me
End Sub
Private Sub INIT_CONTROLLI()
    Set Me.ACSSocio.Connection = TheApp.Database.Connection
    ACSSocio.ApplicationName = App.Title
    ACSSocio.Client = App.EXEName
    ACSSocio.IDFirm = TheApp.IDFirm
    ACSSocio.IDUser = TheApp.IDUser
    ACSSocio.UserName = TheApp.User
    ACSSocio.SearchType = DmtSearchSuppliers
    ACSSocio.HwndContainer = Me.hwnd

    
    Me.ACSSocio.sbLoadCFByIDAnagrafica 7, IDSocioPredefinito
    
End Sub
Private Sub Form_Load()
    ParametroSocioPred
    
    INIT_CONTROLLI
    
    CREA_RECORDSET
        
    GET_GRIGLIA
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3



With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
    .LoadUserSettings
            .ColumnsHeader.Add "IDRiga", "IDRiga", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1100, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisura", "U.M.", dgchar, True, 1100, dgAlignleft
             Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 0
                cl.FormatOptions.FormatNumericThousandSep = "."

             Set cl = .ColumnsHeader.Add("Quantita", "Quantità", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
                
             Set cl = .ColumnsHeader.Add("QuantitaRiscontrata", "Q.tà riscontrata", dgDouble, True, 1300, dgAlignRight)
                cl.Editable = True
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."

             Set cl = .ColumnsHeader.Add("Differenza", "Differenza", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."

            .ColumnsHeader.Add "IDArticoloImballo", "IDOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticoloImballo", "Codice imballo", dgchar, True, 1100, dgAlignleft
            .ColumnsHeader.Add "ArticoloImballo", "Articolo imballo", dgchar, True, 3000, dgAlignleft
            
            Set cl = .ColumnsHeader.Add("ImportoUnitario", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."

            .ColumnsHeader.Add "IDIva", "IDIva", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceIva", "Codice Iva", dgchar, True, 1100, dgAlignleft
            .ColumnsHeader.Add "Iva", "Iva", dgchar, True, 3000, dgAlignleft
            Set cl = .ColumnsHeader.Add("AliquotaIva", "Aliquota", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."


            Set cl = .ColumnsHeader.Add("Sconto1", "% Sc1", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

             Set cl = .ColumnsHeader.Add("Sconto2", "% Sc2", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."


'            Set cl = .ColumnsHeader.Add("ImportoUnitarioScontato", "Importo scontato", dgDouble, True, 1300, dgAlignRight)
'                cl.FormatOptions.FormatNumericRegionalSettings = False
'                cl.FormatOptions.UseFormatControlSettings = False
'                cl.FormatOptions.FormatNumericCurSymbol = ""
'                cl.FormatOptions.FormatNumericDecSep = ","
'                cl.FormatOptions.FormatNumericDecimals = 5
'                cl.FormatOptions.FormatNumericThousandSep = "."



            .ColumnsHeader.Add "IDUnitaDiMisuraCoop", "IDUnitaDiMisuraCoop", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisuraCoop", "U.M. Coop", dgchar, True, 1100, dgAlignleft

    
    Set .Recordset = rsGrigliaRP
    .LoadUserSettings
    .Refresh
    
    
End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"

End Sub

Private Sub GrigliaCorpo_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
    rsGrigliaRP("Differenza").Value = fnNotNullN(rsGrigliaRP("QuantitaRiscontrata").Value) - fnNotNullN(rsGrigliaRP("Quantita").Value)
    rsGrigliaRP.Update
    Me.GrigliaCorpo.Refresh
    
End Sub
Private Sub ParametroSocioPred()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSocioPerRiscontroPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IDSocioPredefinito = fnNotNullN(rs!IDSocioPerRiscontroPeso)
Else
    IDSocioPredefinito = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
