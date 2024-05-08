VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmContrattoDettaglio 
   Caption         =   "SELEZIONA DETTAGLIO CONTRATTO"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrattoDettaglio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   21030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   2535
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20895
      _ExtentX        =   36856
      _ExtentY        =   13150
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
Attribute VB_Name = "frmContrattoDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private ColonnaSelezionata As String

Private Sub cmdConferma_Click()
    CONFERMA_OPERAZIONE
End Sub

Private Sub Form_Activate()
    ColonnaSelezionata = ""
    CREA_RECORDSET
End Sub

Private Sub CREA_RECORDSET()
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim I As Long

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    Set rsGriglia = Nothing
End If
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

If Not (rsContrattoDettaglioSel Is Nothing) Then
    If rsContrattoDettaglioSel.State > 0 Then
        rsContrattoDettaglioSel.Close
    End If
    Set rsContrattoDettaglioSel = Nothing
End If
Set rsContrattoDettaglioSel = New ADODB.Recordset
rsContrattoDettaglioSel.CursorLocation = adUseClient



sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
        Case adInteger
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    End Select
Next

rsGriglia.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable
rsContrattoDettaglioSel.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable

rs.Close
Set rs = Nothing

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
rsContrattoDettaglioSel.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & " WHERE IDOggetto=" & LINK_CONTRATTO
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsGriglia.AddNew
    For I = 0 To rs.Fields.Count - 1
        rsGriglia.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
    Next
    rsGriglia!Selezionato = 0
    rsGriglia.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

GET_GRIGLIA

End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_CURSOR As Long

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
    .LoadUserSettings
        Set cl = .ColumnsHeader.Add("Selezionato", "Sel.", dgBoolean, True, 1300, dgAligncenter)
        cl.Editable = True
        .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Codice pedana", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione pedana", dgchar, False, 3000, dgAlignleft
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, False, 3000, dgAlignleft
        .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, False, 3000, dgAlignleft
        .ColumnsHeader.Add "RV_POIDImballoPrimario", "RV_POIDImballoPrimario", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceImballoPrimario", "Codice imballo primario", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballoPrimario", "Descrizione imballo primario", dgchar, False, 3000, dgAlignleft

        Set cl = .ColumnsHeader.Add("RV_POColliPerPedana", "Colli per pedana", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_numero_colli", "Totale colli", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POColliSfusi", "Colli sfusi", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedana", "N° pedane", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POTaraImballoPrimario", "Tara imballo primario", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_PONumeroConfezioniPerImballo", "N° confezioni", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPerCollo", "Q.tà per collo", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POPesoPerCollo", "Peso per collo", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POMoltiplicatorePerCollo", "Moltiplicatore", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set cl = .ColumnsHeader.Add("Art_peso", "Peso lordo", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_Tara", "Tara", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "Link_Art_unita_di_misura", "Link_Art_unita_di_misura", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisura", "U.M.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "RV_POIDUnitaDiMisuraCoop", "RV_POIDUnitaDiMisuraCoop", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisuraCoop", "U.M. coop", dgchar, False, 1500, dgAlignleft
            
        Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà U.M.", dgDouble, True, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_1", "% Sc1", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
         Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_2", "% Sc2", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_pre_uni_net_sco_net_IVA", "Importo netto IVA", dgDouble, True, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POImportoImballoInArticolo", "Incluso Imballo", dgBoolean, True, 1300, dgAligncenter)
            cl.Editable = True
            cl.BackColor = vbYellow
        Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioImballo", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POTaraUnitariaImballo", "Tara imballo", dgDouble, False, 1300, dgAlignRight)
            'cl.Editable = True
            cl.BackColor = vbGreen
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
    
End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub CONFERMA_OPERAZIONE()
On Error Resume Next
Dim I As Long

rsGriglia.Filter = "Selezionato=1"

While Not rsGriglia.EOF
    rsContrattoDettaglioSel.AddNew
        For I = 0 To rsGriglia.Fields.Count - 1
            rsContrattoDettaglioSel.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
        Next
    rsContrattoDettaglioSel.Update
rsGriglia.MoveNext
Wend

rsGriglia.Close
Set rsGriglia = Nothing

Unload Me

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    Me.cmdConferma.Top = Me.Height - Me.cmdConferma.Height - 600
    Me.Griglia.Width = Me.Width - 360
    Me.Griglia.Height = Me.Height - Me.cmdConferma.Height - 720
    
End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
On Error GoTo ERR_Griglia_AfterChangeFieldValue
Dim IDTipoPesoArticolo As Long
Dim ImportoUnitarioScontato As Double

IDTipoPesoArticolo = fnNotNullN(rsGriglia!RV_POIDTipoPesoArticolo) 'DA ARTICOLO
If IDTipoPesoArticolo = 0 Then IDTipoPesoArticolo = fnNotNullN(rsGriglia!IDRV_POTipoPesoArticoloAzienda) 'DA PARAMETRI AZIENDA

If Column.FieldName = "Art_numero_colli" Then
    If (fnNotNullN(rsGriglia!RV_POColliPerPedana) = 0) Then 'rsGriglia!RV_POColliPerPedana = 1
        rsGriglia!RV_POQuantitaPedana = 0
    Else
        rsGriglia!RV_POQuantitaPedana = rsGriglia!Art_numero_colli / rsGriglia!RV_POColliPerPedana
        If ((rsGriglia!RV_POQuantitaPedana) < 1) Then
            rsGriglia!RV_POQuantitaPedana = 0
        Else
            rsGriglia!RV_POQuantitaPedana = fnRoundDown(rsGriglia!RV_POQuantitaPedana)
        End If
    End If

    rsGriglia!RV_POColliSfusi = rsGriglia!Art_numero_colli - (rsGriglia!RV_POQuantitaPedana * rsGriglia!RV_POColliPerPedana)
    
    
'    rsGriglia!RV_POColliPerPedana = rsGriglia!Art_numero_colli
'    If (fnNotNullN(rsGriglia!RV_POQuantitaPedana) > 0) Then
'    rsGriglia!RV_POColliPerPedana = rsGriglia!Art_numero_colli / fnNotNullN(rsGriglia!RV_POQuantitaPedana)
End If
If Column.FieldName = "RV_POColliPerPedana" Then
    If fnNotNullN(rsGriglia!Art_numero_colli) <= 1 Then
        rsGriglia!Art_numero_colli = (rsGriglia!RV_POQuantitaPedana * rsGriglia!RV_POColliPerPedana) + rsGriglia!RV_POColliSfusi
    Else
        If (fnNotNullN(rsGriglia!RV_POColliPerPedana) = 0) Then 'rsGriglia!RV_POColliPerPedana = 1
            rsGriglia!RV_POQuantitaPedana = 0
        Else
            rsGriglia!RV_POQuantitaPedana = rsGriglia!Art_numero_colli / rsGriglia!RV_POColliPerPedana
            If ((rsGriglia!RV_POQuantitaPedana) < 1) Then
                rsGriglia!RV_POQuantitaPedana = 0
            Else
                rsGriglia!RV_POQuantitaPedana = fnRoundDown(rsGriglia!RV_POQuantitaPedana)
            End If
        End If
        rsGriglia!RV_POColliSfusi = rsGriglia!Art_numero_colli - (rsGriglia!RV_POQuantitaPedana * rsGriglia!RV_POColliPerPedana)
    End If
End If
If Column.FieldName = "RV_POColliSfusi" Then
    rsGriglia!Art_numero_colli = (rsGriglia!RV_POQuantitaPedana * rsGriglia!RV_POColliPerPedana) + rsGriglia!RV_POColliSfusi
End If
If Column.FieldName = "RV_POQuantitaPedana" Then
    rsGriglia!Art_numero_colli = (rsGriglia!RV_POQuantitaPedana * rsGriglia!RV_POColliPerPedana) + rsGriglia!RV_POColliSfusi
End If

rsGriglia!Art_tara = (fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_POTaraUnitariaImballo)) + (fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_PONumeroConfezioniPerImballo) * fnNotNullN(rsGriglia!RV_POTaraImballoPrimario))

If fnNotNullN(rsGriglia!RV_POPesoPerCollo) > 0 Then
    If IDTipoPesoArticolo <= 1 Then
        rsGriglia!Art_peso = fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_POPesoPerCollo)
        rsGriglia!PesoNetto = rsGriglia!Art_peso - rsGriglia!Art_tara
    Else
        rsGriglia!PesoNetto = fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_POPesoPerCollo)
        rsGriglia!Art_peso = rsGriglia!PesoNetto + rsGriglia!Art_tara
    End If
Else
    If IDTipoPesoArticolo <= 1 Then
        rsGriglia!PesoNetto = rsGriglia!Art_peso - rsGriglia!Art_tara
    Else
        rsGriglia!Art_peso = rsGriglia!PesoNetto + rsGriglia!Art_tara
    End If
End If
If fnNotNullN(rsGriglia!RV_PONumeroConfezioniPerImballo) > 0 Then
    If fnNotNullN(rsGriglia!RV_POQuantitaPerCollo) > 0 Then
        rsGriglia!Art_quantita_pezzi = fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_POQuantitaPerCollo)
    Else
        rsGriglia!Art_quantita_pezzi = fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_PONumeroConfezioniPerImballo)
    End If
Else
    If fnNotNullN(rsGriglia!RV_POQuantitaPerCollo) > 0 Then
        rsGriglia!Art_quantita_pezzi = fnNotNullN(rsGriglia!Art_numero_colli) * fnNotNullN(rsGriglia!RV_POQuantitaPerCollo)
    End If
End If
Select Case fnNotNullN(rsGriglia!RV_POIDUnitaDiMisuraCoop)
    Case 1
        rsGriglia!Art_quantita_totale = rsGriglia!Art_numero_colli
    Case 2
        rsGriglia!Art_quantita_totale = rsGriglia!Art_peso
    Case 3
        rsGriglia!Art_quantita_totale = rsGriglia!PesoNetto
    Case 4
        rsGriglia!Art_quantita_totale = rsGriglia!Art_tara
    Case 5
        rsGriglia!Art_quantita_totale = rsGriglia!Art_quantita_pezzi
    Case Else
        rsGriglia!Art_quantita_totale = rsGriglia!PesoNetto
End Select

ImportoUnitarioScontato = fnNotNullN(rsGriglia!Art_prezzo_unitario_neutro) - ((fnNotNullN(rsGriglia!Art_prezzo_unitario_neutro) / 100)) * (fnNotNullN(rsGriglia!Art_sco_in_percentuale_1))
ImportoUnitarioScontato = ImportoUnitarioScontato - ((ImportoUnitarioScontato / 100)) * (fnNotNullN(rsGriglia!Art_sco_in_percentuale_2))

rsGriglia!Art_pre_uni_net_sco_net_IVA = ImportoUnitarioScontato

rsGriglia.UpdateBatch

Me.Griglia.Refresh

Exit Sub
ERR_Griglia_AfterChangeFieldValue:
    MsgBox Err.Description, vbCritical, "Griglia_AfterChangeFieldValue"
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
    Select Case ColonnaSelezionata
        Case "RV_POImportoImballoInArticolo"
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("RV_POImportoImballoInArticolo").Value)), "RV_POImportoImballoInArticolo"
        Case "Selezionato"
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("Selezionato").Value)), "Selezionato"
    End Select
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, nomecampo As String)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
        rsGriglia.Fields(nomecampo).Value = Abs(CLng(Selected))
        rsGriglia.UpdateBatch
        Me.Griglia.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub

Private Sub Griglia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Select Case ColonnaSelezionata
'        Case "RV_POImportoImballoInArticolo"
'            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("RV_POImportoImballoInArticolo").Value)), "RV_POImportoImballoInArticolo"
'        Case "Selezionato"
'            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("Selezionato").Value)), "Selezionato"
'    End Select
End Sub

Private Sub Griglia_MouseMoveOnField(Button As Integer, Shift As Integer, X As Single, Y As Single, ByVal FieldName As String, ByVal Value As Variant, ByVal Row As Integer)
    Select Case FieldName
        Case "RV_POImportoImballoInArticolo"
            ColonnaSelezionata = "RV_POImportoImballoInArticolo"
        Case "Selezionato"
            ColonnaSelezionata = "Selezionato"
        Case Else
            ColonnaSelezionata = ""
    End Select
End Sub

Private Sub Griglia_MouseUpOnField(Button As Integer, Shift As Integer, X As Single, Y As Single, ByVal FieldName As String, ByVal Value As Variant, ByVal Row As Integer)
'    Select Case FieldName
'        Case "RV_POImportoImballoInArticolo"
'            Text1.Text = "RV_POImportoImballoInArticolo"
'        Case "Selezionato"
'            Text1.Text = "Selezionato"
'        Case Else
'            Text1.Text = ""
'    End Select
End Sub
Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If Griglia.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If Griglia.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("Selezionato").Value)), "Selezionato"
            End If
        End If
    End If
    
    
End Sub
