VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmCommissioni 
   Caption         =   "COMMISSIONI"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommissioni.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   16680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   6015
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   16455
      Begin VB.CommandButton cmdEliminaCommissione 
         Caption         =   "Elimina"
         Height          =   375
         Left            =   15000
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalvaCommissione 
         Caption         =   "Salva"
         Height          =   375
         Left            =   15000
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuovaCommissione 
         Caption         =   "Nuovo"
         Height          =   375
         Left            =   15000
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox chkCommissionePerPedana 
         Caption         =   "Commissione per pedana"
         Height          =   315
         Left            =   9600
         TabIndex        =   6
         Top             =   600
         Width           =   2535
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantitaCommissione 
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtImportoTotaleCommissioni 
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   6480
         Visible         =   0   'False
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
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
      Begin DMTEDITNUMLib.dmtCurrency txtImportoRigaComm 
         Height          =   315
         Left            =   8160
         TabIndex        =   5
         Top             =   600
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DmtGridCtl.DmtGrid GrigliaCommissioni 
         Height          =   3135
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
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
      Begin DMTEDITNUMLib.dmtCurrency txtImportoCommissioni 
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtNumber txtPercCommissioni 
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTDataCmb.DMTCombo cboCommissioni 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DmtCodDescCtl.DmtCodDesc CDTipoPedanaCommissione 
         Height          =   615
         Left            =   12240
         TabIndex        =   7
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1085
         PropCodice      =   $"frmCommissioni.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmCommissioni.frx":47869
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmCommissioni.frx":478BF
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
      Begin DMTEDITNUMLib.dmtNumber txtPercCommissioniOri 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Importo totale commissioni"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   22
         Top             =   6240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Importo"
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Valore"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Perc.%"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo di commissione"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblTotaleMercePerComm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   9615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità"
         Height          =   255
         Index           =   4
         Left            =   7080
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Perc. ori."
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTotaleMercePerComm 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   4920
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmCommissioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As DmtOleDbLib.adoResultset

Private FLAG_COMMISSIONE_PER_PEDANA As Long
Private LINK_TIPO_VALORE_COMMISSIONE As Long
Private NuovoRecordComm As Long


Private Sub fnGrigliaCommissioni()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader


sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = Cn.OpenResultset(sSQL)
            'Set rsEvent = rsGriglia.Data
    
        
    
        With Me.GrigliaCommissioni.ColumnsHeader
            .Clear
                .Add "IDRV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc", dgInteger, False, 500, 0, True, True, False
                .Add "IDRV_POTipoCommissione", "IDRV_POTipoCommissione", dgInteger, False, 500, 0, True, True, False
                .Add "TipoCommissione", "Tipo commissione", dgchar, True, 4500, 0, True, True, False
                Set cl = .Add("Percentuale", "%", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .Add("ImportoTotale", "Valore", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .Add("Quantita", "Quantita", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .Add("ImportoRiga", "Importo %", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .Add "IDRV_POTipoPedana", "IDRV_POTipoPedana", dgInteger, False, 500, 0, True, True, False
                .Add "CodiceTipoPedana", "Codice tipo pedana", dgchar, True, 2000, 0, True, True, False
                .Add "TipoPedana", "Descrizione tipo pedana", dgchar, True, 3000, 0, True, True, False
                .Add "APercentuale", "Comm. per pedana", dgBoolean, True, 1500, dgAligncenter, True, True, False
            
            Set Me.GrigliaCommissioni.Recordset = rsGriglia.Data
            Me.GrigliaCommissioni.Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
    If (rsGriglia.BOF And rsGriglia.EOF) Then
        cmdNuovaCommissione_Click
    End If
    
End Sub

Private Sub chkCommissionePerPedana_Click()
    If FLAG_COMMISSIONE_PER_PEDANA = 0 Then Exit Sub
    
    If Me.chkCommissionePerPedana.Value = vbUnchecked Then
        Me.CDTipoPedanaCommissione.Load 0
        ABILITA_CONTROLLI_COMMISSIONE 3
    End If
    If Me.chkCommissionePerPedana.Value = vbChecked Then
        
        ABILITA_CONTROLLI_COMMISSIONE 2
    End If
    
End Sub


Private Sub cmdEliminaCommissione_Click()
On Error GoTo ERR_cmdEliminaCommissione_Click
Dim sSQL As String
Dim Testo As String

If CONTROLLO_COLLEGAMENTO_DOC_COMPLETO(oDoc.IDOggetto, oDoc.IDTipoOggetto) = False Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Una o più righe del documento risultano collegate ad uno o più documenti di nota di credito o nota di debito." & vbCrLf
    Testo = Testo & "Se si continua con questo comando potrebbero verificarsi delle incongruenze nella liquidazione" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio dati") = vbNo Then Exit Sub
End If
If NuovoRecordComm = 0 Then
    sSQL = "DELETE FROM RV_POCommissioniPerDoc "
    sSQL = sSQL & "WHERE IDRV_POCommissioniPerDoc=" & fnNotNullN(Me.GrigliaCommissioni("IDRV_POCommissioniPerDoc").Value)
    Cn.Execute sSQL
    
    fnGrigliaCommissioni

    Changed_Commissioni = True
End If

Exit Sub
ERR_cmdEliminaCommissione_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaCommissione_Click"
End Sub


Private Sub cmdNuovaCommissione_Click()
On Error Resume Next
    NuovoRecordComm = 1
    
    
    FLAG_COMMISSIONE_PER_PEDANA = 0
    LINK_TIPO_VALORE_COMMISSIONE = 1
    Me.cboCommissioni.WriteOn 0
    Me.txtPercCommissioni.Value = 0
    Me.txtImportoCommissioni.Value = 0 'Valore
    Me.txtImportoRigaComm.Value = 0 'Totale visualizzato
    
    Me.txtQuantitaCommissione.Value = 1
    
    Me.CDTipoPedanaCommissione.Load 0
    Me.chkCommissionePerPedana.Value = vbUnchecked
    
    
    Me.cboCommissioni.SetFocus
    
    ABILITA_CONTROLLI_COMMISSIONE 1
    
End Sub

Private Sub ABILITA_CONTROLLI_COMMISSIONE(TIpo As Long)
    chkCommissionePerPedana.Enabled = False
    txtPercCommissioni.Enabled = True
    txtImportoCommissioni.Enabled = False
    txtQuantitaCommissione.Enabled = False
    txtImportoRigaComm.Enabled = True
    CDTipoPedanaCommissione.Enabled = False

    Select Case TIpo
        Case 1 'Commissione normale
            chkCommissionePerPedana.Enabled = False
            txtPercCommissioni.Enabled = True
            txtImportoCommissioni.Enabled = False
            txtQuantitaCommissione.Enabled = False
            txtImportoRigaComm.Enabled = True
            CDTipoPedanaCommissione.Enabled = False
            Me.txtPercCommissioniOri.Enabled = False
        Case 2 'Commissione per pedana
            chkCommissionePerPedana.Enabled = True
            txtPercCommissioni.Enabled = False
            txtImportoCommissioni.Enabled = False
            txtQuantitaCommissione.Enabled = False
            txtImportoRigaComm.Enabled = True
            CDTipoPedanaCommissione.Enabled = True
            Me.txtPercCommissioniOri.Enabled = False
        Case 3 'Commissione per pallets
            chkCommissionePerPedana.Enabled = True
            txtPercCommissioni.Enabled = True
            txtImportoCommissioni.Enabled = True
            txtQuantitaCommissione.Enabled = True
            txtImportoRigaComm.Enabled = False
            CDTipoPedanaCommissione.Enabled = False
            Me.txtPercCommissioniOri.Enabled = False
        Case 4
            chkCommissionePerPedana.Enabled = False
            txtPercCommissioni.Enabled = False
            txtImportoCommissioni.Enabled = False
            txtQuantitaCommissione.Enabled = False
            txtImportoRigaComm.Enabled = False
            CDTipoPedanaCommissione.Enabled = False
            Me.txtPercCommissioniOri.Enabled = True
    End Select
End Sub


Private Sub cboCommissioni_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim CommissionePerPedana As Long


FLAG_COMMISSIONE_PER_PEDANA = 0
LINK_TIPO_VALORE_COMMISSIONE = 1

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & Me.cboCommissioni.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtPercCommissioni.Value = 0
    Me.txtPercCommissioniOri.Value = 0
    Me.chkCommissionePerPedana.Value = Unchecked
Else
    Me.txtPercCommissioni.Value = fnNotNullN(rs!Percentuale)
    Me.txtPercCommissioniOri.Value = fnNotNullN(rs!Percentuale)
    Me.chkCommissionePerPedana.Value = Abs(fnNotNullN(rs!CommissionePerPedana))
    FLAG_COMMISSIONE_PER_PEDANA = Abs(fnNotNullN(rs!CommissionePerPedana))
    If fnNotNullN(rs!IDRV_POTipoValoreDocumento) = 0 Then
        LINK_TIPO_VALORE_COMMISSIONE = 1
    Else
        LINK_TIPO_VALORE_COMMISSIONE = fnNotNullN(rs!IDRV_POTipoValoreDocumento)
    End If
    
End If

rs.CloseResultset
Set rs = Nothing

If ((FLAG_COMMISSIONE_PER_PEDANA = 0) And (LINK_TIPO_VALORE_COMMISSIONE = 1)) Then ABILITA_CONTROLLI_COMMISSIONE 1
If ((FLAG_COMMISSIONE_PER_PEDANA = 0) And (LINK_TIPO_VALORE_COMMISSIONE > 1)) Then ABILITA_CONTROLLI_COMMISSIONE 4
If ((FLAG_COMMISSIONE_PER_PEDANA = 1) And (Me.chkCommissionePerPedana.Value = vbChecked)) Then ABILITA_CONTROLLI_COMMISSIONE 2
If ((FLAG_COMMISSIONE_PER_PEDANA = 1) And (Me.chkCommissionePerPedana.Value = vbUnchecked)) Then ABILITA_CONTROLLI_COMMISSIONE 3

If LINK_TIPO_VALORE_COMMISSIONE = 1 Then
    If NuovoRecordComm = 1 Then
        txtPercCommissioni_LostFocus
    End If
Else
    If NuovoRecordComm = 1 Then
        txtPercCommissioniOri_LostFocus
    End If
End If
End Sub

Private Sub cmdSalvaCommissione_Click()
On Error GoTo ERR_cmdSalvaCommissione_Click
Dim sSQL As String
Dim Testo As String
Dim NumeroRecord As Long


If Me.cboCommissioni.CurrentID = 0 Then
    MsgBox "Inserire il tipo di commissione", vbInformation, "Salvataggio dati"
    Exit Sub
End If
If CONTROLLO_COLLEGAMENTO_DOC_COMPLETO(oDoc.IDOggetto, oDoc.IDTipoOggetto) = False Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Una o più righe del documento risultano collegate ad uno o più documenti di nota di credito o nota di debito." & vbCrLf
    Testo = Testo & "Se si continua con questo comando potrebbero verificarsi delle incongruenze nella liquidazione" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio dati") = vbNo Then Exit Sub

End If

'If (Me.chkCommissionePerPedana.Value = vbChecked) Then
'    If Me.CDTipoPedanaCommissione.KeyFieldID = 0 Then
'        MsgBox "Inserire il tipo di pedana", vbInformation, "Salvataggio dati"
'        Exit Sub
'    End If
'End If


If NuovoRecordComm = 1 Then
    NumeroRecord = Me.GrigliaCommissioni.ListCount
    
    sSQL = "INSERT INTO RV_POCommissioniPerDoc ("
    sSQL = sSQL & "IDRV_POCommissioniPerDoc, IDOggetto, IDRV_POTipoCommissione, Percentuale, Importo, ImportoRiga,"
    sSQL = sSQL & "Quantita, APercentuale, ImportoTotale, IDRV_POTipoPedana, PercentualeDaCommissione "
    sSQL = sSQL & ") "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc") & ", "
    sSQL = sSQL & oDoc.IDOggetto & ", "
    sSQL = sSQL & Me.cboCommissioni.CurrentID & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPercCommissioni.Value) & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & fnNormNumber(Me.txtImportoRigaComm.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtQuantitaCommissione.Value) & ", "
    sSQL = sSQL & Abs(Me.chkCommissionePerPedana.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtImportoCommissioni.Value) & ", "
    sSQL = sSQL & Me.CDTipoPedanaCommissione.KeyFieldID & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPercCommissioniOri.Value)
    
    sSQL = sSQL & ")"
    
Else
    NumeroRecord = Me.GrigliaCommissioni.ListIndex - 1
    sSQL = "UPDATE RV_POCommissioniPerDoc SET "
    sSQL = sSQL & "IDRV_POTipoCommissione=" & Me.cboCommissioni.CurrentID & ", "
    sSQL = sSQL & "Percentuale=" & fnNormNumber(Me.txtPercCommissioni.Value) & ", "
    sSQL = sSQL & "Importo=" & 0 & ", "
    sSQL = sSQL & "ImportoRiga=" & fnNormNumber(Me.txtImportoRigaComm.Value) & ", "
    sSQL = sSQL & "Quantita=" & fnNormNumber(Me.txtQuantitaCommissione.Value) & ", "
    sSQL = sSQL & "APercentuale=" & Abs(Me.chkCommissionePerPedana.Value) & ", "
    sSQL = sSQL & "ImportoTotale=" & fnNormNumber(Me.txtImportoCommissioni.Value) & ", "
    sSQL = sSQL & "IDRV_POTipoPedana=" & Me.CDTipoPedanaCommissione.KeyFieldID & ", "
    sSQL = sSQL & "PercentualeDaCommissione=" & fnNormNumber(Me.txtPercCommissioniOri.Value)
    sSQL = sSQL & " WHERE IDRV_POCommissioniPerDoc=" & fnNotNullN(Me.GrigliaCommissioni("IDRV_POCommissioniPerDoc").Value)
End If


Cn.Execute sSQL


fnGrigliaCommissioni



Me.GrigliaCommissioni.Recordset.Move NumeroRecord

Changed_Commissioni = True

Exit Sub
ERR_cmdSalvaCommissione_Click:
    MsgBox Err.Description, vbCritical, "cmdSalvaCommissione_Click"
End Sub


Private Sub Form_Load()

    Changed_Commissioni = False

    InitControlli
    
    fnGrigliaCommissioni
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsGriglia Is Nothing) Then
        rsGriglia.CloseResultset
        Set rsGriglia = Nothing
    End If
End Sub

Private Sub txtPercCommissioniOri_LostFocus()
Dim ImportoTotale As Double

Select Case LINK_TIPO_VALORE_COMMISSIONE
    Case 1
        ImportoTotale = TOTALE_MERCE
    Case 2
        ImportoTotale = TOTALE_MERCE_LORDA
    Case 3
        ImportoTotale = Totale_Documento_netto_iva
    Case 4
        ImportoTotale = Totale_Documento_lordo_iva
    Case Else
        ImportoTotale = TOTALE_MERCE
End Select

If ImportoTotale = 0 Then Exit Sub

txtImportoRigaComm.Value = ((ImportoTotale / 100) * Me.txtPercCommissioniOri.Value)

If TOTALE_MERCE > 0 Then
    Me.txtPercCommissioni.Value = (Me.txtImportoRigaComm.Value / TOTALE_MERCE) * 100
End If

End Sub
Private Sub txtPercCommissioni_LostFocus()
    If (Me.chkCommissionePerPedana.Value = vbChecked) Then Exit Sub
    If LINK_TIPO_VALORE_COMMISSIONE > 1 Then Exit Sub
    
    If TOTALE_MERCE > 0 Then
        Me.txtImportoRigaComm.Value = (TOTALE_MERCE / 100) * Me.txtPercCommissioni.Value
    End If
End Sub
Private Sub txtImportoCommissioni_Change()
    GET_TOTALE_COMMISIONE
End Sub

Private Sub txtImportoRigaComm_Change()
'    If (FLAG_COMMISSIONE_PER_PEDANA = 1) Then
'        GET_TOTALE_COMMISIONE
'        Exit Sub
'    End If
    GET_TOTALE_COMMISIONE
    
    If ((FLAG_COMMISSIONE_PER_PEDANA = 1) And (Me.chkCommissionePerPedana.Value = vbChecked)) Then Exit Sub
    
    If TOTALE_MERCE > 0 Then
        Me.txtPercCommissioni.Value = (Me.txtImportoRigaComm.Value / TOTALE_MERCE) * 100
    End If
End Sub
Private Sub GET_TOTALE_COMMISIONE()
    If FLAG_COMMISSIONE_PER_PEDANA = 1 Then
        If (Me.chkCommissionePerPedana.Value = vbChecked) Then
            txtImportoCommissioni.Value = 0
        Else
            txtImportoRigaComm.Value = Me.txtQuantitaCommissione.Value * Me.txtImportoCommissioni.Value
        End If
    End If
End Sub

Private Sub txtQuantitaCommissione_Change()
GET_TOTALE_COMMISIONE
End Sub

Private Function CONTROLLO_COLLEGAMENTO_DOC_COMPLETO(IDOggetto As Long, IDTipoOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NonBLoccare As Boolean


NonBLoccare = True
'''CONTROLLO SU NOTA DI CREDITO'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDOggetto FROM ValoriOggettoDettaglio0016 "
sSQL = sSQL & "WHERE RV_POIDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POIDTipoOggetto=" & IDTipoOggetto
'sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    
    NonBLoccare = False

Else
    NonBLoccare = True
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If NonBLoccare = True Then

    '''CONTROLLO SU NOTA DI DEBITO'''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT IDOggetto FROM ValoriOggettoDettaglio0007 "
    sSQL = sSQL & "WHERE RV_POIDOggetto=" & IDOggetto
    sSQL = sSQL & " AND RV_POIDTipoOggetto=" & IDTipoOggetto
'    sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
    
        NonBLoccare = False
    
    Else
        NonBLoccare = True
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End If

'NonBloccare = true vuol dire che non ci sono collegamenti alle righe del documento
'NonBloccare = false vuol dire che esiste almento una riga di nota riferito al codice lotto di vendita
CONTROLLO_COLLEGAMENTO_DOC_COMPLETO = NonBLoccare
End Function

Private Sub InitControlli()
    With Me.cboCommissioni
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoCommissione"
        .DisplayField = "TipoCommissione"
        .SQL = "SELECT * FROM RV_POTipoCommissione"
        .SQL = .SQL & " ORDER BY TipoCommissione"
    End With

    'Tipo di Pedana commissioni
    With Me.CDTipoPedanaCommissione
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceTipoPedana"
        .DescriptionField = "TipoPedana"
        .KeyField = "IDRV_POTipoPedana"
        .TableName = "RV_POIETipoPedana"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        
        .PropDescrizione.Caption = "Descrizione"
        
        .CodeCaption4Find = "Codice Articolo"
        
        .DescriptionCaption4Find = "Descrizione Articolo"

        .CodeIsNumeric = False
    End With
    
    Me.lblTotaleMercePerComm(0).Caption = "TOTALE MERCE DOCUMENTO: " & FormatNumber(TOTALE_MERCE, 2)
    Me.lblTotaleMercePerComm(1).Caption = "TOTALE MERCE DOCUMENTO (INCLUSO IMBALLO): " & FormatNumber(TOTALE_MERCE_LORDA, 2)

End Sub

Private Sub GrigliaCommissioni_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
        
    NuovoRecordComm = 0
    
    Me.cboCommissioni.WriteOn Me.GrigliaCommissioni("IDRV_POTipoCommissione").Value
    Me.txtPercCommissioni.Value = fnNotNullN(Me.GrigliaCommissioni("Percentuale").Value)
    Me.txtPercCommissioniOri.Value = fnNotNullN(Me.GrigliaCommissioni("PercentualeDaCommissione").Value)
    Me.chkCommissionePerPedana.Value = fnNotNullN(Me.GrigliaCommissioni("APercentuale").Value)
    Me.txtImportoCommissioni.Value = fnNotNullN(Me.GrigliaCommissioni("Importo").Value)
    Me.txtImportoRigaComm.Value = fnNotNullN(Me.GrigliaCommissioni("ImportoRiga").Value)
    Me.txtQuantitaCommissione.Value = fnNotNullN(Me.GrigliaCommissioni("Quantita").Value)
    Me.txtImportoCommissioni.Value = fnNotNullN(Me.GrigliaCommissioni("ImportoTotale").Value)
    Me.CDTipoPedanaCommissione.Load fnNotNullN(Me.GrigliaCommissioni("IDRV_POTipoPedana").Value)
    
End Sub

