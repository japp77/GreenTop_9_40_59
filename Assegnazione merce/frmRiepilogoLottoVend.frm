VERSION 5.00
Begin VB.Form frmRiepilogoLottoVend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Riepilogo lavorazione"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiepilogoLottoVend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDocumentoCollegato 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Documento di vendita collegato"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmRiepilogoLottoVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.txtDocumentoCollegato.Text = GET_LAVORAZIONE_VENDITA(LINK_LAVORAZIONE_RIEP, 0)
    
End Sub
Private Function GET_LAVORAZIONE_VENDITA(IDAssegnazioneMerce As Long, IDProcessoIVGamma As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDMovimento, IDTipoOggetto, Oggetto, NumeroDocumento, DataDocumento FROM Movimento "
sSQL = sSQL & "WHERE (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
sSQL = sSQL & " AND RV_POIDAssegnazioneMerce=" & IDAssegnazioneMerce
sSQL = sSQL & " AND RV_POIDProcessoIVGamma=" & IDProcessoIVGamma
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LAVORAZIONE_VENDITA = ""
Else
    GET_LAVORAZIONE_VENDITA = fnNotNull(rs!Oggetto)
    GET_LAVORAZIONE_VENDITA = GET_LAVORAZIONE_VENDITA & " n° " & fnNotNull(rs!NumeroDocumento)
    
    GET_LAVORAZIONE_VENDITA = GET_LAVORAZIONE_VENDITA & " del " & fnNotNull(rs!DataDocumento)
End If

rs.CloseResultset
Set rs = Nothing
    
End Function
