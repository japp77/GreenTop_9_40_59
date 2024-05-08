VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.3#0"; "DmtGridCtl.ocx"
Begin VB.Form frmCoda 
   BorderStyle     =   0  'None
   Caption         =   "Coda dei salvataggi"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
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
   ScaleHeight     =   4890
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      Picture         =   "frmCoda.frx":0000
      ScaleHeight     =   1785
      ScaleWidth      =   7305
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      Begin VB.Label lblInfo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ATTENDERE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1500
         Width           =   7095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   2520
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
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
   Begin VB.Label Label1 
      Caption         =   "NON UTILIZZARE IL COMANDO CTRL-ALT-CANC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "ATTENDERE LA CODA DI REGISTRAZIONE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   5295
   End
End
Attribute VB_Name = "frmCoda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub SettaggioGriglia()
'On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_POTMP "
    'sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto
    sSQL = sSQL & "ORDER BY IDSessione, IDUtente "
    
    Set rsArt = Cn.OpenResultset(sSQL)
        Set rsEvent = rsArt.Data
    
    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDSessione", "Sessione", dgNumeric, True, 1000, dgAlignleft
            .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 3000, dgAlignleft
        Set .Recordset = rsArt.Data
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub

Private Sub cmdElimina_Click()
Dim sSQL As String
Dim Testo As String

If TheApp.IDUser <> Me.Griglia.AllColumns("IDUtente") Then
    Me.Timer1.Enabled = False
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Eliminando l'utente dalla coda di salvataggio si potrebbero avere degli effetti indesiderati." & vbCrLf
    Testo = Testo & "Continuare con il comando di eliminazione?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione dati di coda") = vbYes Then
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & Me.Griglia("IDUtente").Value
        'sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto
        Cn.Execute sSQL
        
        SettaggioGriglia
        
        Me.Timer1.Enabled = True
    End If
End If



End Sub

Private Sub Form_Activate()
    SettaggioGriglia
    Me.cmdElimina.Enabled = False
    DoEvents
    Me.Timer1.Enabled = True
End Sub

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
    If Me.cmdElimina.Enabled = False Then
        Me.cmdElimina.Enabled = True
    Else
        Me.cmdElimina.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rsArt Is Nothing) Then
    rsArt.CloseResultset
    Set rsArt = Nothing
End If

Me.Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    SettaggioGriglia
    DoEvents
End Sub
