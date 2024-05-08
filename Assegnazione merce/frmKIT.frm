VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmKIT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Componenti KIT in lavorazione"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKIT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   17760
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
      Height          =   495
      Left            =   15000
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpoOrdine 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   9128
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
Attribute VB_Name = "frmKIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

For I = 0 To rsKIT.Fields.Count - 1
    Select Case rsKIT.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsGriglia.Fields.Append rsKIT.Fields(I).Name, rsKIT.Fields(I).Type, rsKIT.Fields(I).DefinedSize, rsKIT.Fields(I).Attributes
        Case adInteger
            rsGriglia.Fields.Append rsKIT.Fields(I).Name, rsKIT.Fields(I).Type, , rsKIT.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rsKIT.Fields(I).Name, rsKIT.Fields(I).Type, , rsKIT.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rsKIT.Fields(I).Name, adBoolean, , rsKIT.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rsKIT.Fields(I).Name, adDouble, , rsKIT.Fields(I).Attributes
    End Select
Next

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

If Not ((rsKIT.EOF) And (rsKIT.BOF)) Then
    rsKIT.MoveFirst
    
    While Not rsKIT.EOF
        rsGriglia.AddNew
            For I = 0 To rsKIT.Fields.Count - 1
                rsGriglia.Fields(rsKIT.Fields(I).Name).Value = rsKIT.Fields(I).Value
            Next
        rsGriglia.Update
    rsKIT.MoveNext
    Wend

End If

GET_GRIGLIA

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long


OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

With Me.GrigliaCorpoOrdine
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        
        Set cl = .ColumnsHeader.Add("Selezionato", "Selezionato", dgBoolean, True, 1500, dgAligncenter)
            cl.Editable = True
        
        .ColumnsHeader.Add "IDRV_PODistintaBaseRighe", "IDRV_PODistintaBaseRighe", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDRV_PODistintaBase", "IDRV_PODistintaBase", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "Articolo", "Descrizione articolo", dgchar, True, 4500, dgAlignleft
        
        .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "UnitaDiMisura", "Unità di misura", dgchar, True, 1100, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("Quantita", "Q.tà unitaria", dgDouble, True, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        
        .ColumnsHeader.Add "IDUnitaDiMisuraCoop", "IDUnitaDiMisuraCoop", dgNumeric, False, 500, dgAlignRight
        .ColumnsHeader.Add "UnitaDiMisuraCoop", "Unità di misura coop.", dgchar, True, 1100, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("QuantitaTotale", "Q.tà totale", dgDouble, True, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        
        
        .ColumnsHeader.Add "IDArticoloMerce", "IDArticoloMerce", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDArticoloImballo", "IDArticoloImballo", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDArticoloConfezione", "IDArticoloConfezione", dgInteger, False, 500, dgAlignRight
        .ColumnsHeader.Add "IDRV_PODistintaBaseRigheConf", "IDRV_PODistintaBaseRigheConf", dgNumeric, False, 500, dgAlignRight

        Set cl = .ColumnsHeader.Add("Costo", "Costo unitario", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."

        
        Set cl = .ColumnsHeader.Add("CostoTotale", "Costo totale", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "Annotazioni", "Note", dgchar, False, 2500, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("GiacenzaLottoImballo", "Giacenza", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("DisponibilitaLottoImballo", "Disponibilità", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        
        
        
        
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub cmdConferma_Click()
    
    CONFERMA
    
    Unload Me
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CREA_RECORDSET
End Sub

Private Sub CONFERMA()
On Error GoTo ERR_CONFERMA
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        
        rsKIT.MoveFirst
        
        While Not rsKIT.EOF
            rsKIT.Delete
        rsKIT.MoveNext
        Wend

       rsGriglia.MoveFirst
       
       While Not rsGriglia.EOF
           rsKIT.AddNew
                For I = 0 To rsGriglia.Fields.Count - 1
                    rsKIT.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
                Next
                
           rsKIT.Update
           
       rsGriglia.MoveNext
       Wend
       
       rsGriglia.MoveFirst
    
    End If
    
Exit Sub
ERR_CONFERMA:
    MsgBox Err.Description, vbCritical, "CONFERMA"

End Sub

Private Sub GrigliaCorpoOrdine_DblClick()
If ((rsGriglia.EOF) And (rsGriglia.BOF)) Then Exit Sub

If fnNotNullN(rsGriglia!RV_POTracciabilitaImballo) = 0 Then Exit Sub

QUANTITA_KIT_SEL = rsGriglia!QuantitaTotale
LINK_KIT_SEL = rsGriglia!IDRV_PODistintaBaseRighe
ARTICOLO_KIT_SEL = rsGriglia!Articolo

frmLottoImballoKIT.Show vbModal
    
End Sub

Private Sub GrigliaCorpoOrdine_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaCorpoOrdine_KeyPress
    If KeyAscii = vbKeySpace Then
        If Me.GrigliaCorpoOrdine.GuiMode = dgNormal Then
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("Selezionato").Value))
        End If
    End If

Exit Sub
ERR_GrigliaCorpoOrdine_KeyPress:
    MsgBox Err.Description, vbCritical, "GrigliaCorpoOrdine_KeyPress"
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean)
On Error GoTo ERR_sbSelectSelectedRow

    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
                
        rsGriglia.Fields("Selezionato").Value = Abs(CLng(Selected))
    
        
        rsGriglia.UpdateBatch
                
        Me.GrigliaCorpoOrdine.Refresh

    End If
Exit Sub
ERR_sbSelectSelectedRow:
MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
