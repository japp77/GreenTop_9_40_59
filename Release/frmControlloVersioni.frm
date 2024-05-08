VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmControlloVersioni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTROLLO VERSIONI CLIENT"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlloVersioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRicerca 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   9000
      Width           =   11895
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15690
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
Attribute VB_Name = "frmControlloVersioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private fso As FileSystemObject
Private Sub Form_Load()
On Error GoTo ERR_Form_Load

Dim fld As Folder
Dim fil As File
    
Set fso = New FileSystemObject
    
CREA_RECORDSET

Set fld = fso.GetFolder(MenuOptions.ProgramsPath)
For Each fil In fld.Files
    ADD_FILE fil
Next

Set fil = Nothing
Set fld = Nothing
Set fso = Nothing


GET_GRIGLIA

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub CREA_RECORDSET()
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "NomeFile", adVarChar, 255, adFldIsNullable
rsGriglia.Fields.Append "VersioneFile", adVarChar, 255, adFldIsNullable
rsGriglia.Fields.Append "UltimoAggiornamento", adVarChar, 255, adFldIsNullable


rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        If Len(Trim(Me.txtRicerca.Text)) > 0 Then
            rsGriglia.Filter = "NomeFile LIKE %" & Me.txtRicerca.Text & "%"
        End If
        
        With Me.Griglia
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            
            .ColumnsHeader.Clear

                .ColumnsHeader.Add "NomeFile", "Nome", dgchar, True, 4000, dgAlignleft
                .ColumnsHeader.Add "VersioneFile", "Versione", dgchar, True, 1500, dgAlignRight
                .ColumnsHeader.Add "UltimoAggiornamento", "Ultimo aggiormamento", dgchar, True, 1500, dgAligncenter
            
            Set .Recordset = rsGriglia
            .LoadUserSettings
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati"
End Sub

Private Sub txtRicerca_Change()
    GET_GRIGLIA
End Sub
Private Sub ADD_FILE(f As File)
On Error Resume Next
    If Mid(f.Name, 1, 5) = "RV_PO" Then
        x = fso.GetFileVersion(MenuOptions.ProgramsPath & "\" & f.Name)
        
        If Len(Trim(x)) > 0 Then
            rsGriglia.AddNew
                rsGriglia!NomeFile = f.Name
                rsGriglia!VersioneFile = x
                rsGriglia!UltimoAggiornamento = Format(f.DateLastModified, "dd/MM/yyyy hh:mm:ss")
            rsGriglia.Update
        End If
        
    End If


End Sub
