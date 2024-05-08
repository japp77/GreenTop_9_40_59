VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmParticelle 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Particelle"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9990
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
   ScaleHeight     =   3675
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
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
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5530
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
Attribute VB_Name = "frmParticelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGrigliaAss As ADODB.Recordset

Private Sub cmdConferma_Click()
    RIPORTA_DATI_PARTICELLA = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    GET_PARTICELLE_SOCIO
    
    fnGrigliaParticelle
    
    RIPORTA_DATI_PARTICELLA = False
    
End Sub
Private Sub GET_PARTICELLE_SOCIO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
'''''''ELIMINAZIONE DATI TEMPORANEI'''''''''
sSQL = "DELETE FROM RV_PO01_TMPParticelleSocio "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''


sSQL = "SELECT RV_PO01_TerrenoRighe.FoglioCatastale, RV_PO01_TerrenoRighe.MappaCatastale, RV_PO01_TerrenoRighe.ParticellaCatastale, "
sSQL = sSQL & "RV_PO01_TerrenoRighe.SubCatastale, RV_PO01_TerrenoTesta.Terreno, RV_PO01_TerrenoTesta.Indirizzo, RV_PO01_TerrenoTesta.Cap,"
sSQL = sSQL & "Comune.Comune , Provincia.Provincia, RV_PO01_TerrenoRighe.SuperficieHA, RV_PO01_TerrenoRighe.SuperficieMQ, "
sSQL = sSQL & "RV_PO01_TerrenoTesta.IDSocio, RV_PO01_TerrenoRighe.IDRV_PO01_TerrenoRighe, RV_PO01_TerrenoTesta.IDRV_PO01_TerrenoTesta  "
sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
sSQL = sSQL & "RV_PO01_TerrenoTesta ON Comune.IDComune = RV_PO01_TerrenoTesta.IDComune RIGHT OUTER JOIN "
sSQL = sSQL & "RV_PO01_TerrenoRighe ON RV_PO01_TerrenoTesta.IDRV_PO01_TerrenoTesta = RV_PO01_TerrenoRighe.IDRV_PO01_TerrenoTesta "
sSQL = sSQL & "WHERE RV_PO01_TerrenoTesta.IDSocio=" & frmMain.CDSocio.KeyFieldID


Set rs = Cn.OpenResultset(sSQL)
Set rsNew = New ADODB.Recordset

rsNew.Open "RV_PO01_TMPParticelleSocio", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    If GET_ESISTENZA_ASSOCIAZIONE(LINK_SERRA_SOCIO, fnNotNullN(rs!IDRV_PO01_TerrenoRighe)) = False Then
        rsNew.AddNew
            rsNew!IDRV_PO01_TMPParticelleSocio = fnGetNewKey("RV_PO01_TMPParticelleSocio", "IDRV_PO01_TMPParticelleSocio")
            rsNew!IDUtente = TheApp.IDUser
            rsNew!IDRV_PO01_TerrenoRighe = fnNotNullN(rs!IDRV_PO01_TerrenoRighe)
            rsNew!IDRV_PO01_Terreno = fnNotNullN(rs!IDRV_PO01_TerrenoTesta)
            rsNew!Registra = False
            rsNew!MappaCatastale = Trim(fnNotNull(rs!MappaCatastale))
            rsNew!FoglioCatastale = Trim(fnNotNull(rs!FoglioCatastale))
            rsNew!PerticellaCatastale = Trim(fnNotNull(rs!ParticellaCatastale))
            rsNew!SubParticella = Trim(fnNotNull(rs!SubCatastale))
            rsNew!DimensioneMQ = fnNotNullN(rs!SuperficieMQ)
            rsNew!DimensioneHA = Trim(fnNotNull(rs!SuperficieHA))
            rsNew!Terreno = Trim(fnNotNull(rs!Terreno))
            rsNew!Comune = Trim(fnNotNull(rs!Terreno))
            rsNew!Provincia = Trim(fnNotNull(rs!Provincia))
            rsNew!Indirizzo = Trim(fnNotNull(rs!Indirizzo))
            rsNew!Cap = Trim(fnNotNull(rs!Cap))

            
        rsNew.Update
    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub fnGrigliaParticelle()
'On Error GoTo ERR_fnGrigliaAssegnazione
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT * FROM RV_PO01_TMPParticelleSocio "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        
        Set rsGrigliaAss = New ADODB.Recordset
        rsGrigliaAss.CursorLocation = adUseClient
        rsGrigliaAss.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
    
        With Me.GrigliaCorpo
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_PO01_TMPParticelleSocio", "IDRV_PO01_TMPParticelleSocio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_TerrenoRighe", "IDRV_PO01_TerrenoRighe", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_PO01_Terreno", "IDRV_PO01_Terreno", dgInteger, False, 500, dgAlignleft
                Set cl = .ColumnsHeader.Add("Registra", "Registra", dgBoolean, True, 1000, dgAligncenter)
                    cl.Editable = True
                .ColumnsHeader.Add "MappaCatastale", "Mappa", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "FoglioCatastale", "Foglio", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "PerticellaCatastale", "Particella", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "SubParticella", "Sub", dgchar, True, 1000, dgAlignleft
                Set cl = .ColumnsHeader.Add("DimensioneMQ", "Superficie M.q.", dgDouble, True, 1500, dgAlignRight)
                    cl.Editable = True
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "DimensioneHA", "Superficie H.a.", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "Indirizzo", "Indirizzo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "Comune", "Comune", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "Provincia", "Prov.", dgchar, True, 800, dgAlignleft
                .ColumnsHeader.Add "Cap", "C.A.P.", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "Terreno", "Terreno", dgchar, True, 2000, dgAlignleft
            Set .Recordset = rsGrigliaAss
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub
Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)


    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGrigliaAss.Fields("Registra").Value)), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    If Not rsGrigliaAss.EOF And Not rsGrigliaAss.BOF Then
                
        rsGrigliaAss.Fields("Registra").Value = Abs(CLng(Selected))
        'sbCheckSelected
        
        rsGrigliaAss.UpdateBatch
        Me.GrigliaCorpo.Refresh
    End If
End Sub
Private Function GET_ESISTENZA_ASSOCIAZIONE(IDSerraSocio As Long, IDParticella As Long) As Boolean
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT IDRV_PO01_SerraRighe FROM RV_PO01_SerraRighe "
sSQL = sSQL & "WHERE IDRV_PO01_Serra=" & IDSerraSocio
sSQL = sSQL & " AND IDRV_PO01_TerrenoRighe=" & IDParticella

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ASSOCIAZIONE = False
Else
    GET_ESISTENZA_ASSOCIAZIONE = True
End If


rs.CloseResultset
Set rs = Nothing

End Function
