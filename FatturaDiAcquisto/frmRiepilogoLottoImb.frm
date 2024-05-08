VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRiepilogoLottoImb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RIEPILOGO LOTTO IMBALLO"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRiepilogoLottoImb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRiepilogoFinale 
      Caption         =   "TOTALI DAI MOVIMENTI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   17775
      Begin VB.TextBox txtQtaDispMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   15240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtQtaScaricataMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   15240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtQtaCaricataMov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   15240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità disponibile"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   11640
         TabIndex        =   16
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità scaricata"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   11640
         TabIndex        =   15
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità caricata"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   11280
         TabIndex        =   14
         Top             =   240
         Width           =   3855
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   7646
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
   Begin VB.Frame fraRiepilogo 
      Caption         =   "Riepilogo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   17775
      Begin VB.CommandButton Command2 
         Height          =   285
         Left            =   14160
         Picture         =   "frmRiepilogoLottoImb.frx":4781A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salva riferimento esterno"
         Top             =   480
         Width           =   350
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   13820
         Picture         =   "frmRiepilogoLottoImb.frx":47DA4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Sblocca inserimento riferimento esterno"
         Top             =   480
         Width           =   350
      End
      Begin VB.TextBox txtRiferimentoEsterno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtQtaGiacenza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   16200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtQtaCaricata 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtImballo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtLottoImballo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Riferimento esterno"
         Height          =   255
         Index           =   4
         Left            =   9000
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Giacenza"
         Height          =   255
         Index           =   3
         Left            =   16200
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Q.tà caricata"
         Height          =   255
         Index           =   2
         Left            =   14640
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Imballo"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Lotto imballo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmRiepilogoLottoImb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset


Private Sub Command1_Click()
    Me.txtRiferimentoEsterno.Locked = Not Me.txtRiferimentoEsterno.Locked
End Sub

Private Sub Command2_Click()
On Error GoTo ERR_Command2_Click
Dim sSQL As String

sSQL = "UPDATE RV_POLottoImballo SET "
sSQL = sSQL & " RiferimentoEsterno=" & fnNormString(Me.txtRiferimentoEsterno.Text)
sSQL = sSQL & " WHERE IDRV_POLottoImballo=" & LINK_LOTTO_IMBALLO
Cn.Execute sSQL

MsgBox "Aggiornamento avvenuto con successo", vbInformation, "Salvataggio dati"

GET_GRIGLIA LINK_LOTTO_IMBALLO

Exit Sub
ERR_Command2_Click:
    MsgBox Err.Description, vbCritical, "Command2_Click"
End Sub

Private Sub Form_Load()
    GET_LOTTO_IMBALLO LINK_LOTTO_IMBALLO
    GET_GRIGLIA LINK_LOTTO_IMBALLO
    RiepilogoMovimenti LINK_LOTTO_IMBALLO
End Sub
Private Sub GET_GRIGLIA(IDLottoImballo As Long)
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

sSQL = "SELECT * FROM RV_POIERiepilogoLottoImballo "
sSQL = sSQL & " WHERE RV_POIDLottoImballo=" & IDLottoImballo
sSQL = sSQL & " ORDER BY DataMovimento DESC"

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient
rsGriglia.Open sSQL, Cn.InternalConnection
    
With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDMovimento", "IDLetteraIntento", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDMagazzino", "IDMagazzino", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDFunzione", "IDFunzione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDProcessoPerFunzione", "IDProcessoPerFunzione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "DataMovimento", "Data movimento", dgDate, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Oggetto", "Documento", dgchar, True, 4000, dgAlignleft
        
        .ColumnsHeader.Add "DataDocumento", "DataDocumento", dgDate, False, 1800, dgAlignleft
        .ColumnsHeader.Add "NumeroDocumento", "NumeroDocumento", dgchar, False, 1200, dgAlignRight
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneArticolo", "Descrizione imballo", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisura", "Unità di misura", dgchar, False, 1200, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoAnagrafica", "IDTipoAnagrafica", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Anagrafica", "Anagrafica", dgchar, False, 2000, dgAlignleft
        .ColumnsHeader.Add "Nome", "Nome", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "PartitaIva", "PartitaIva", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "CodiceFiscale", "CodiceFiscale", dgchar, False, 1500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDLottoImballo", "RV_POIDLottoImballo", dgNumeric, False, 500, dgAlignleft
        
        .ColumnsHeader.Add "LottoImballo", "LottoImballo", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "RiferimentoEsterno", "Rif. esterno", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "Funzione", "Funzione", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "PartecipaGiacenza", "Segno", dgchar, True, 1200, dgAligncenter
        .ColumnsHeader.Add "Processo", "Processo", dgchar, False, 1200, dgAlignleft
        
        
        Set cl = .ColumnsHeader.Add("QuantitaTotale", "Quantita", dgDouble, True, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        .ColumnsHeader.Add "PartecipaDisponibilita", "Segno disp.", dgchar, False, 1200, dgAlignleft
        .ColumnsHeader.Add "TipoDocumentoGT", "Tipo oggetto GT", dgchar, False, 1200, dgAlignleft
        


    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub RiepilogoMovimenti(IDLottoImballo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QuantitaCaricata As Double
Dim QuantitaScaricata As Double
Dim QuantitaDisponibile As Double

QuantitaCaricata = 0
QuantitaScaricata = 0
QuantitaDisponibile = 0


sSQL = "SELECT * FROM RV_POIERiepilogoLottoImballo "
sSQL = sSQL & "WHERE RV_POIDLottoImballo=" & IDLottoImballo

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If fnNotNull(rs!PartecipaGiacenza) = "+" Then
        QuantitaCaricata = QuantitaCaricata + fnNotNullN(rs!QuantitaTotale)
        
    End If
    If fnNotNull(rs!PartecipaGiacenza) = "-" Then
        QuantitaScaricata = QuantitaScaricata + fnNotNullN(rs!QuantitaTotale)
        
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

QuantitaDisponibile = QuantitaCaricata - QuantitaScaricata

Me.txtQtaCaricataMov.Text = FormatNumber(QuantitaCaricata, 2)
Me.txtQtaScaricataMov.Text = FormatNumber(QuantitaScaricata, 2)
Me.txtQtaDispMov.Text = FormatNumber(QuantitaDisponibile, 2)


End Sub
Private Sub GET_LOTTO_IMBALLO(IDLottoImballo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIELottoImballo "
sSQL = sSQL & "WHERE IDRV_POLottoImballo=" & IDLottoImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtLottoImballo.Text = ""
    Me.txtImballo.Text = ""
    Me.txtQtaCaricata.Text = "0"
    Me.txtQtaGiacenza.Text = "0"
    Me.txtRiferimentoEsterno.Text = ""
Else
    Me.txtLottoImballo.Text = fnNotNull(rs!LottoImballo)
    Me.txtImballo.Text = fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo)
    Me.txtQtaCaricata.Text = fnNotNullN(rs!QuantitaCaricata)
    Me.txtQtaGiacenza.Text = fnNotNullN(rs!Giacenza)
    Me.txtRiferimentoEsterno.Text = fnNotNull(rs!RiferimentoEsterno)
    
End If

rs.CloseResultset
Set rs = Nothing

Me.txtQtaCaricata.Text = FormatNumber(Me.txtQtaCaricata.Text, 2)
Me.txtQtaGiacenza.Text = FormatNumber(Me.txtQtaGiacenza.Text, 2)

End Sub

