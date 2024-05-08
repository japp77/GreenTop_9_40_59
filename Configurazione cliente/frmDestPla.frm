VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmDestPla 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Configurazione altre destinazioni per il planning degli ordini"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8670
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
   ScaleHeight     =   4335
   ScaleWidth      =   8670
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
      Left            =   6840
      TabIndex        =   1
      Top             =   3880
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6588
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
Attribute VB_Name = "frmDestPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsagg As ADODB.Recordset

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
Dim rsNew As ADODB.Recordset

''''ELIMINAZIONE DATI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POConfigurazioneClienteAD "
sSQL = sSQL & "WHERE IDAnagrafica=" & frmMain.CDCliente.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If (rsagg.EOF) And (rsagg.BOF) Then Exit Sub
    
rsagg.Filter = "Selezionato=1"

If (rsagg.EOF) And (rsagg.BOF) Then Exit Sub


rsagg.MoveFirst


sSQL = "SELECT * FROM RV_POConfigurazioneClienteAD "
sSQL = sSQL & "WHERE IDAnagrafica=" & frmMain.CDCliente.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsagg.EOF
    rsNew.AddNew
        rsNew!IDRV_POConfigurazioneClienteAD = fnGetNewKey("RV_POConfigurazioneClienteAD", "IDRV_POConfigurazioneClienteAD")
        rsNew!IDRV_POConfigurazioneCliente = LINK_CONFIGURAZIONE_CLIENTE
        rsNew!IDAnagrafica = frmMain.CDCliente.KeyFieldID
        rsNew!IDAzienda = TheApp.IDFirm
        rsNew!IDSitoPerAnagrafica = fnNotNullN(rsagg!IDSitoPerAnagrafica)
        rsNew!Abbreviazione = fnNotNull(rsagg!Abbreviazione)
    rsNew.Update
rsagg.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

Unload Me
Exit Sub

ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"

End Sub

Private Sub Form_Load()
    CREA_RECORDSET_TMP
    
    GET_GRIGLIA
End Sub
Private Sub CREA_RECORDSET_TMP()

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRiga As Long



If Not (rsagg Is Nothing) Then
    If rsagg.State > 0 Then
        rsagg.Close
    End If
    Set rsagg = Nothing
End If

Set rsagg = New ADODB.Recordset

rsagg.CursorLocation = adUseClient

rsagg.Fields.Append "NumeroRiga", adInteger, , adFldIsNullable
rsagg.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
rsagg.Fields.Append "Selezionato", adBoolean, , adFldIsNullable
rsagg.Fields.Append "SitoPerAnagrafica", adVarChar, 250, adFldIsNullable
rsagg.Fields.Append "Abbreviazione", adVarChar, 50, adFldIsNullable


rsagg.Open , , adOpenKeyset, adLockBatchOptimistic
NumeroRiga = 1

sSQL = "SELECT * FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & frmMain.CDCliente.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    rsagg.AddNew
        rsagg!NumeroRiga = NumeroRiga
        rsagg!IDSitoPerAnagrafica = fnNotNullN(rs!IDSitoPerAnagrafica)
        rsagg!SitoPerAnagrafica = fnNotNull(rs!SitoPerAnagrafica)
        PRELEVA_DATI_CONFIGURATI fnNotNullN(rs!IDSitoPerAnagrafica)
    rsagg.Update
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub PRELEVA_DATI_CONFIGURATI(IDSitoPerAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POConfigurazioneClienteAD "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsagg!Abbreviazione = ""
    rsagg!Selezionato = 0
Else
    rsagg!Abbreviazione = fnNotNull(rs!Abbreviazione)
    rsagg!Selezionato = 1

End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_GRIGLIA()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
                        
            .ColumnsHeader.Add "NumeroRiga", "NumeroRiga", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgInteger, False, 500, dgAlignleft
            Set cl = .ColumnsHeader.Add("Selezionato", "Selezionato", dgBoolean, True, 1500, dgAligncenter)
                cl.Editable = True
                
            .ColumnsHeader.Add "SitoPerAnagrafica", "Altra destinazione", dgchar, True, 3500, dgAlignleft
            Set cl = .ColumnsHeader.Add("Abbreviazione", "Abbreviazione", dgchar, True, 2000, dgAlignleft)
                cl.Editable = True
                
                
        Set .Recordset = rsagg
        .LoadUserSettings
        .Refresh
        
    End With
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsagg Is Nothing) Then
        If rsagg.State > 0 Then
            rsagg.Close
        End If
        Set rsagg = Nothing
    End If
End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
    rsagg.UpdateBatch
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)

        'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            If Not rsagg.EOF And Not rsagg.BOF Then
                sbSelectSelectedRow Not CBool(rsagg.Fields("Selezionato").Value), 2
            End If
            
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsagg.EOF And Not rsagg.BOF Then
        rsagg.Fields("Selezionato").Value = Abs(CLng(Selected))
        rsagg.UpdateBatch
        Me.Griglia.Refresh
    End If
End Sub

