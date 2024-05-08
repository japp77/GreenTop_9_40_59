VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.0#0"; "DMTLblLinkCtl.ocx"
Begin VB.Form frmCambioValuta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cambio valuta"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAggiungi 
      Caption         =   "Aggiungi cambio"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin DMTLblLinkCtl.LabelLink Link_Cambio 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "Cambio valuta"
      Name            =   "LabelLink"
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6376
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
Attribute VB_Name = "frmCambioValuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub GET_GRIGLIA()
Dim sSQL As String

Dim OLDCursor As Long
Dim cl As dgColumnHeader


sSQL = "SELECT * FROM Cambio "
sSQL = sSQL & "WHERE IDValuta=" & frmMain.cboValuta.CurrentID
sSQL = sSQL & " AND IDValutaDiRiferimento=" & oDoc.DBDefaults.Link_Val_valuta_nazionale
sSQL = sSQL & " ORDER BY DataCambio DESC"
    
OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient
rsGriglia.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDCambio", "IDCambio", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "DataCambio", "Data cambio", dgDate, True, 1800, dgAligncenter
            Set cl = .ColumnsHeader.Add("Valore", "Valore", dgDouble, True, 1800, dgAlignRight)
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

End Sub

Private Sub cmdAggiungi_Click()
    Me.Link_Cambio.RunApplication
End Sub

Private Sub Form_Load()

    Set Me.Link_Cambio.Application = TheApp    'L’oggetto Application
    Link_Cambio.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    Link_Cambio.IDFunction = fncTrovaIDFunzione("VALUTA")


    Me.Caption = GET_VALUTA(frmMain.cboValuta.CurrentID) & " --> " & GET_VALUTA(oDoc.DBDefaults.Link_Val_valuta_nazionale)
    
    GET_GRIGLIA

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsGriglia Is Nothing) Then
        If rsGriglia.State > 0 Then
            rsGriglia.Close
        End If
        Set rsGriglia = Nothing
    End If
End Sub
Private Function GET_VALUTA(IDValuta As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Valuta FROM Valuta "
sSQL = sSQL & "WHERE IDValuta=" & IDValuta

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALUTA = ""
Else
    GET_VALUTA = fnNotNull(rs!Valuta)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GrigliaCorpo_DblClick()
    frmMain.cboCambioValuta.Refresh
    frmMain.cboCambioValuta.WriteOn fnNotNullN(Me.GrigliaCorpo.AllColumns("IDCambio").Value)
    Unload Me
End Sub
Private Function fncTrovaIDFunzione(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
'sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Link_Cambio_AfterRunServerApplication(ByVal lIDResultKey As Long)
    GET_GRIGLIA
End Sub

Private Sub Link_Cambio_BeforeRunServerApplication()
    Me.Link_Cambio.IDReturn = frmMain.cboValuta.CurrentID
End Sub

