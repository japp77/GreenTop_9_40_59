VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form FrmInserimentoRighe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inserimento nuove trattenute "
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
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
   ScaleHeight     =   3795
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5106
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
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   11520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   375
      Left            =   11520
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin DMTEDITNUMLib.dmtCurrency txtTrattenuta 
      Height          =   315
      Left            =   9720
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   " 0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      CurrencySymbol  =   ""
      AllowEmpty      =   0   'False
      DecFinalZeros   =   -1  'True
   End
   Begin VB.TextBox txtRiga 
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin DMTDataCmb.DMTCombo cboTipoTrattenutaAggiuntiva 
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
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
   End
   Begin DMTEDITNUMLib.dmtNumber txtPercentualeTrattenuta 
      Height          =   315
      Left            =   8880
      TabIndex        =   2
      Top             =   360
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo trattenuta"
      Height          =   255
      Index           =   17
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "%"
      Height          =   255
      Index           =   7
      Left            =   8880
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Trattenuta"
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Descrizione trattenuta"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "FrmInserimentoRighe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Nuovo As Integer
Private TOTALE_TRATTENUTE_LIBERE As Double



Private Sub cmdElimina_Click()
If MsgBox("Vuoi eliminare la trattenuta aggiuntiva?", vbQuestion + vbYesNo, "Elimina dati") = vbYes Then
    CnDMT.Execute "DELETE FROM RV_POTMPLiquidazioneRighe WHERE IDRV_POTMPLiquidazioneRighe=" & Me.Griglia.AllColumns("IDRV_POTMPLiquidazioneRighe").Value
    rsGriglia.Requery
    Griglia.Refresh
    
                    
End If
End Sub

Private Sub cmdNuovo_Click()
    Nuovo = 0
    Me.txtRiga.Text = ""
    Me.txtTrattenuta.Value = 0
    Me.txtPercentualeTrattenuta.Value = 0
    Me.cboTipoTrattenutaAggiuntiva.WriteOn 0
    
    Me.txtRiga.SetFocus
End Sub

Private Sub cmdSalva_Click()
    If Me.cboTipoTrattenutaAggiuntiva.CurrentID = 0 Then
        MsgBox "Inserire il tipo di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.cboTipoTrattenutaAggiuntiva.SetFocus
        Exit Sub
    End If
    If Me.cboTipoTrattenutaAggiuntiva.CurrentID = 3 Then
        MsgBox "Tipo di trattenuta incompatibile", vbInformation, "Controllo inserimento dati"
        Me.cboTipoTrattenutaAggiuntiva.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtRiga.Text)) = 0 Then
        MsgBox "Inserire una descrizione di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.txtRiga.SetFocus
        Exit Sub
    End If
    If Me.txtPercentualeTrattenuta.Value = 0 Then
        MsgBox "Inserire una percentuale di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.txtPercentualeTrattenuta.SetFocus
        Exit Sub
    End If
    
    
    If Nuovo = 0 Then
        sSQL = "INSERT INTO RV_POTMPLiquidazioneRighe ("
        sSQL = sSQL & "IDRV_POTMPLiquidazioneRighe, IDRV_POTMPLiquidazione, DescrizioneAggiuntiva, ImportoTrattenuta,"
        sSQL = sSQL & "IDRV_POTipoTrattenutaAggiuntiva, Percentuale) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POTMPLiquidazioneRighe", "IDRV_POTMPLiquidazioneRighe") & ", "
        sSQL = sSQL & LINK_DOCUMENTO_TMP_LIQ & ", "
        sSQL = sSQL & fnNormString(Me.txtRiga.Text) & ", "
        sSQL = sSQL & fnNormNumber(Me.txtTrattenuta.Value) & ", "
        sSQL = sSQL & Me.cboTipoTrattenutaAggiuntiva.CurrentID & ", "
        sSQL = sSQL & fnNormNumber(Me.txtPercentualeTrattenuta.Value) & ")"
    Else
        sSQL = "UPDATE RV_POTMPLiquidazioneRighe SET "
        sSQL = sSQL & "DescrizioneAggiuntiva=" & fnNormString(Me.txtRiga.Text) & ", "
        sSQL = sSQL & "ImportoTrattenuta=" & fnNormNumber(Me.txtTrattenuta.Value) & ", "
        sSQL = sSQL & "IDRV_POTipoTrattenutaAggiuntiva=" & Me.cboTipoTrattenutaAggiuntiva.CurrentID & ", "
        sSQL = sSQL & "Percentuale=" & fnNormNumber(Me.txtPercentualeTrattenuta.Value) & " "
        sSQL = sSQL & "WHERE IDRV_POTMPLiquidazioneRighe=" & Me.Griglia.AllColumns("IDRV_POTMPLiquidazioneRighe").Value
    End If
    
    CnDMT.Execute sSQL
    
    rsGriglia.Requery
    Griglia.Refresh


End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    'Tipo trattenute aggiuntive
    With Me.cboTipoTrattenutaAggiuntiva
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoTrattenutaAggiuntiva"
        .DisplayField = "TipoTrattenuta"
        .Sql = "SELECT * FROM RV_POTipoTrattenutaAggiuntiva ORDER BY TipoTrattenuta"
        .Fill
    End With
    
    
    fncGriglia



    Nuovo = 0
    Me.txtRiga.Text = ""
    Me.txtTrattenuta.Value = 0
    Me.cboTipoTrattenutaAggiuntiva.WriteOn 0
    Me.txtPercentualeTrattenuta.Value = 0
        
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sSQL As String
Dim TotaleTrattenutaAggiuntiva As Double
Dim TotaleTrattenutaAggiuntivaRiepilogo As Double


TotaleTrattenutaAggiuntiva = GET_TOTALE_TRATTENUTA_AGGIUNTIVA_TMP(LINK_DOCUMENTO_TMP_LIQ, 1)
TotaleTrattenutaAggiuntivaRiepilogo = GET_TOTALE_TRATTENUTA_AGGIUNTIVA_TMP(LINK_DOCUMENTO_TMP_LIQ, 2)

sSQL = "UPDATE RV_POTMPLiquidazione SET "
sSQL = sSQL & "TotaleTrattenuteAggiuntive=" & fnNormNumber(TotaleTrattenutaAggiuntiva) & ", "
sSQL = sSQL & "TotaleTrattenuteRiepilogo=" & fnNormNumber(TotaleTrattenutaAggiuntivaRiepilogo) & ", "
sSQL = sSQL & "NettoLiquidazione=" & fnNormNumber(FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleDocumento").Value - FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleTrattenuta").Value - (TotaleTrattenutaAggiuntiva + TotaleTrattenutaAggiuntivaRiepilogo)) & " "
sSQL = sSQL & "WHERE IDRV_POTMPLiquidazione=" & LINK_DOCUMENTO_TMP_LIQ

CnDMT.Execute sSQL

End Sub

Private Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader

    sSQL = "SELECT RV_POTMPLiquidazioneRighe.*, RV_POTipoTrattenutaAggiuntiva.TipoTrattenuta "
    sSQL = sSQL & "FROM RV_POTMPLiquidazioneRighe LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoTrattenutaAggiuntiva ON "
    sSQL = sSQL & "RV_POTMPLiquidazioneRighe.IDRV_POTipoTrattenutaAggiuntiva = RV_POTipoTrattenutaAggiuntiva.IDRV_POTipoTrattenutaAggiuntiva "
    sSQL = sSQL & "WHERE IDRV_POTMPLiquidazione = " & LINK_DOCUMENTO_TMP_LIQ

    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
        
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            
                .ColumnsHeader.Add "IDRV_POTMPLiquidazioneRighe", "IDRV_POTMPLiquidazioneRighe", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POTMPLiquidazione", "IDRV_POTMPLiquidazione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "DescrizioneAggiuntiva", "Descrizione", dgchar, True, 4500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POTipoTrattenutaAggiuntiva", "IDRV_POTipoTrattenutaAggiuntiva", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoTrattenuta", "Tipo trattenuta", dgchar, True, 2000, dgAlignleft, True, True, False
                .ColumnsHeader.Add "IDRV_POSegnoTrattenuta", "IDRV_POTipoTrattenutaAggiuntiva", dgInteger, False, 500, dgAlignleft
                Set cl = .ColumnsHeader.Add("Percentuale", "Percentuale", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("ImportoTrattenuta", "Importo", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."


                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not (rsGriglia Is Nothing) Then
        rsGriglia.Close
        Set rsGriglia = Nothing
    End If
End Sub


Private Sub Griglia_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    Me.txtRiga.Text = fnNotNull(Me.Griglia.AllColumns("DescrizioneAggiuntiva").Value)
    Me.txtTrattenuta.Value = fnNotNullN(Me.Griglia.AllColumns("ImportoTrattenuta").Value)
    Me.cboTipoTrattenutaAggiuntiva.WriteOn fnNotNullN(Me.Griglia.AllColumns("IDRV_POTipoTrattenutaAggiuntiva").Value)
    Me.txtPercentualeTrattenuta.Value = fnNotNullN(Me.Griglia.AllColumns("Percentuale").Value)
    Nuovo = 1
End Sub

Private Sub txtPercentualeTrattenuta_Change()
    If FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleDocumento").Value > 0 Then
        Me.txtTrattenuta.Value = (FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleDocumento").Value / 100) * Me.txtPercentualeTrattenuta.Value
    End If
End Sub

Private Sub txtTrattenuta_Change()
    If FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleDocumento").Value > 0 Then
        Me.txtPercentualeTrattenuta.Value = (Me.txtTrattenuta.Value / FrmVisualizzaLiquidazione.Griglia.AllColumns("TotaleDocumento").Value) * 100
    End If
End Sub
Private Function GET_TOTALE_TRATTENUTA_AGGIUNTIVA_TMP(IDLiquidazione As Long, IDTipoTrattenutaAggiutiva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sum(ImportoTrattenuta) As TotaleTrattenute "
sSQL = sSQL & "FROM RV_POTMPLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POTMPLiquidazione=" & IDLiquidazione
sSQL = sSQL & " AND IDRV_POTipoTrattenutaAggiuntiva=" & IDTipoTrattenutaAggiutiva

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_TRATTENUTA_AGGIUNTIVA_TMP = 0
Else
    GET_TOTALE_TRATTENUTA_AGGIUNTIVA_TMP = fnNotNullN(rs!TotaleTrattenute)
End If



rs.CloseResultset
Set rs = Nothing
End Function
