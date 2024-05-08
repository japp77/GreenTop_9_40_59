VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmNuovoOrdine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NUOVO ORDINE"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10950
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
   ScaleHeight     =   6330
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   10695
      Begin VB.CommandButton cmdLetteraIntento 
         Height          =   315
         Left            =   720
         Picture         =   "frmNuovoOrdine.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Lettere di intento del cliente"
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdEliminaRifLetInt 
         Height          =   315
         Left            =   360
         Picture         =   "frmNuovoOrdine.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Elimina riferimento lettera intento"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtDescrizioneRigaDoc 
         Height          =   525
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   4680
         Width           =   10215
      End
      Begin VB.TextBox txtAnnotazioniInterna 
         Height          =   645
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   3720
         Width           =   10215
      End
      Begin VB.TextBox txtNumeroOrdineCliente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7680
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtAnnotazioniOrdine 
         Height          =   525
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2880
         Width           =   10215
      End
      Begin DMTDataCmb.DMTCombo cboDestinazione 
         Height          =   315
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "CREA ORDINE"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   5520
         Width           =   1935
      End
      Begin DmtCodDescCtl.DmtCodDesc cdCliente 
         Height          =   615
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         PropCodice      =   $"frmNuovoOrdine.frx":0B14
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmNuovoOrdine.frx":0B62
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmNuovoOrdine.frx":0BB4
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
      Begin DMTDATETIMELib.dmtDate txtDataOrdine 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   1050
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboPagamento 
         Height          =   315
         Left            =   3000
         TabIndex        =   9
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
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
      Begin DMTDataCmb.DMTCombo cboSezionale 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
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
      Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   1050
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboVettore 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   1050
         Width           =   2655
         _ExtentX        =   4683
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
      Begin DMTDATETIMELib.dmtDate txtDataPartenza 
         Height          =   315
         Left            =   9240
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataOrdineCliente 
         Height          =   315
         Left            =   6120
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboVettoreSuccessivo 
         Height          =   315
         Left            =   5640
         TabIndex        =   16
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDATETIMELib.dmtDate txtDataArrivoMerceL 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataArrivoMerce 
         Height          =   315
         Left            =   8280
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtTime txtOraArrivoMerce 
         Height          =   315
         Left            =   9600
         TabIndex        =   3
         Top             =   480
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboTipoTrasporto 
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   1050
         Width           =   2175
         _ExtentX        =   3836
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
      Begin DMTDATETIMELib.dmtTime txtOraArrivoMerceL 
         Height          =   315
         Left            =   4560
         TabIndex        =   15
         Top             =   2280
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   1485
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTDATETIMELib.dmtDate txtDataLetteraIntento 
         Height          =   315
         Left            =   1560
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtNLetteraIntento 
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Top             =   1680
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboIvaCliente 
         Height          =   315
         Left            =   360
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lblLetteraIntento 
         Caption         =   "Lettera d'intento"
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Data arrivo"
         Height          =   255
         Index           =   12
         Left            =   8280
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ora arrivo"
         Height          =   255
         Index           =   20
         Left            =   9600
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo trasporto"
         Height          =   255
         Index           =   13
         Left            =   5640
         TabIndex        =   38
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Data arrivo"
         Height          =   255
         Index           =   14
         Left            =   3240
         TabIndex        =   37
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ora arrivo"
         Height          =   255
         Index           =   21
         Left            =   4560
         TabIndex        =   36
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Luogo di presa merce"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   35
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Vettore successivo"
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   34
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Annotazioni finali del corpo del documento di evasione"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   33
         Top             =   4440
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Annotazioni interne"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   32
         Top             =   3480
         Width           =   6135
      End
      Begin VB.Label Label5 
         Caption         =   "Num. ordine cli."
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Data ordine cli."
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Annotazioni di fatturazione"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   29
         Top             =   2640
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Data partenza"
         Height          =   255
         Index           =   5
         Left            =   9240
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Vettore"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   27
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Numero"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Pagamento "
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   24
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Destinazione diversa"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Data ordine"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmNuovoOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VarNumeroDoc As String


Private Sub cboSezionale_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_PERIODO_IVA As Long

If Me.txtDataOrdine.Value = 0 Then Me.txtDataOrdine.Value = Date



LINK_PERIODO_IVA = fnGetPeriodoIVA(Me.txtDataOrdine.Text)

sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
sSQL = sSQL & " AND IDTipoModulo=4"
sSQL = sSQL & " AND IDPeriodoIVA=" & LINK_PERIODO_IVA

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "INSERT INTO ProgressivoSezionale ("
    sSQL = sSQL & "IDProgressivoSezionale, IDTipoModulo, IDPeriodoIVA, IDSezionale, "
    sSQL = sSQL & "ProgressivoDisponibile, DataUltimaVariazione, IDUtenteUltimaVariazione, "
    sSQL = sSQL & "VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
    sSQL = sSQL & 4 & ", "
    sSQL = sSQL & LINK_PERIODO_IVA & ", "
    sSQL = sSQL & Me.cboSezionale.CurrentID & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ")"
    
    Cn.Execute sSQL
    
    Me.txtNumeroOrdine.Value = 1
Else
    Me.txtNumeroOrdine.Value = fnNotNullN(rs!ProgressivoDisponibile)
End If


rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub cdCliente_ChangeElement()
Dim LINK_CLIENTE_IVA As Long
    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica WHERE IDAnagrafica=" & Me.cdCliente.KeyFieldID
        .Fill
    End With
    
    Me.cboPagamento.WriteOn GET_IDPAGAMENTO_DEFAULT
        
    Me.cboVettore.WriteOn GET_VETTORE_DEFAULT(Me.cdCliente.KeyFieldID)

    LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdCliente.KeyFieldID)
    
    If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(Me.cdCliente.KeyFieldID, TheApp.IDFirm, Year(Me.txtDataOrdine.Text)) = 1 Then
        Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO(Me.cdCliente.KeyFieldID, TheApp.IDFirm, Year(Me.txtDataOrdine.Text))
        LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
    End If
    
    Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA

    GET_INFO_CLIENTE_FIDO_BLOCCO Me.cdCliente.KeyFieldID
    
End Sub
Private Function GET_VETTORE_DEFAULT(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_VETTORE As Long
Dim LINK_TIPO_SPEDIZIONE As Long

sSQL = "SELECT IDTipoSpedizione FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_SPEDIZIONE = 0
Else
    LINK_TIPO_SPEDIZIONE = fnNotNullN(rs!IDTipoSpedizione)
End If

rs.CloseResultset
Set rs = Nothing


If LINK_TIPO_SPEDIZIONE = 0 Then
    sSQL = "SELECT IDTipoSpedizione FROM ConfigurazioneVendite "
    'sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        LINK_TIPO_SPEDIZIONE = 0
    Else
        LINK_TIPO_SPEDIZIONE = fnNotNullN(rs!IDTipoSpedizione)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If LINK_TIPO_SPEDIZIONE <> 3 Then
        GET_VETTORE_DEFAULT = 0
        Exit Function
    End If
    
    
Else
    If LINK_TIPO_SPEDIZIONE <> 3 Then
        GET_VETTORE_DEFAULT = 0
        Exit Function
    End If
End If

''''VETTORE DI DAFAUL DEL CLIENTE'''''''''''''''''''''''''''
sSQL = "SELECT IDVettoreDefault FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaCliente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_VETTORE = 0
Else
    LINK_VETTORE = fnNotNullN(rs!IDVettoreDefault)
End If


rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If LINK_VETTORE = 0 Then
    sSQL = "SELECT IDVettore FROM ConfigurazioneVendite "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        LINK_VETTORE = 0
    Else
        LINK_VETTORE = fnNotNullN(rs!IDVettore)
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing
End If


Me.cboTipoTrasporto.WriteOn LINK_TIPO_SPEDIZIONE
GET_VETTORE_DEFAULT = LINK_VETTORE

End Function


Private Sub cmdEliminaRifLetInt_Click()
On Error GoTo ERR_cmdEliminaRifLetInt_Click
Dim Testo As String
Dim LINK_CLIENTE_IVA As Long

If Me.txtIDLetteraIntento.Value = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il riferimento alla lettera d'intento?" & vbCrLf
If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento lettera d'intento") = vbNo Then Exit Sub

Me.txtIDLetteraIntento.Value = 0

LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdCliente.KeyFieldID)

LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
    
Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA

Exit Sub
ERR_cmdEliminaRifLetInt_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifLetInt_Click"
End Sub

Private Sub cmdLetteraIntento_Click()
On Error GoTo ERR_cmdLetteraIntento_Click
Dim LINK_CLIENTE_IVA As Long

LINK_CLIENTE_IVA = 0
'If Me.txtIDOrdine.Value = 0 Then Exit Sub

Set Control_Return = Me.txtIDLetteraIntento
Set Control_Return_Cliente = Me.cdCliente
Set Control_Return_Data_Ordine = Me.txtDataOrdine

frmLetteraIntento.Show vbModal

LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdCliente.KeyFieldID)

If Me.txtIDLetteraIntento.Value > 0 Then
    
    LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
        
    Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
    
End If
    
Exit Sub
ERR_cmdLetteraIntento_Click:
    MsgBox Err.Description, vbCritical, "cmdLetteraIntento_Click"
End Sub

Private Sub cmdRicerca_Click()
    If CHECK_ABILITAZIONE_DIAMANTE = False Then Exit Sub

    If Me.cdCliente.KeyFieldID = 0 Then
        MsgBox "Inserire il cliente", vbInformation, "Inserimento ordine"
        Me.cdCliente.SetFocus
        Exit Sub
    End If

    If Me.cboPagamento.CurrentID = 0 Then
        MsgBox "Inserire il pagamento", vbInformation, "Inserimento ordine"
        Me.cboPagamento.SetFocus
        Exit Sub
    End If

    If Me.cboSezionale.CurrentID = 0 Then
        MsgBox "Inserire il sezionale", vbInformation, "Inserimento ordine"
        Me.cboSezionale.SetFocus
        Exit Sub
    End If

    If Me.txtDataOrdine.Value = 0 Then
        MsgBox "Inserire la data del documento", vbInformation, "Inserimento ordine"
        Me.txtDataOrdine.SetFocus
        Exit Sub
    End If
    
    If LINK_BLOCCO_CLIENTE = 1 Then
        MsgBox "Impossibile salvare il documento corrente poichè il cliente risulta bloccato", vbInformation, "Inserimento ordine"
        Me.cdCliente.SetFocus
        Exit Sub
    End If
    
    If GET_CONTROLLO_FIDO_CLIENTE = False Then
        AVVIA_FIDO_DOPO_CONTROLLO = False
        frmFido.Show vbModal
        If AVVIA_FIDO_DOPO_CONTROLLO = False Then
            Me.cdCliente.SetFocus
            Exit Sub
        End If
    End If
    
Screen.MousePointer = 11
    FN_CREA_ORDINE
Screen.MousePointer = 0

frmMain.txtIDOrdineCliente.Value = LINK_ORDINE_RIF
frmMain.cdCliente.Load Me.cdCliente.KeyFieldID
frmMain.txtDataOrdine.Value = Me.txtDataOrdine.Value
frmMain.txtNumeroOrdine.Value = CLng(VarNumeroDoc)

Unload Me
End Sub


Private Sub Form_Activate()
    Me.cboSezionale.WriteOn GET_IDSEZIONALE_ORDINE(15)
    Me.cdCliente.SetFocus
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

     With Me.cdCliente
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Clienti"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    With Me.cboPagamento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT * FROM Pagamento"
        .Fill
    End With

    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica WHERE IDAnagrafica=" & Me.cdCliente.KeyFieldID
        .Fill
    End With

    With Me.cboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT * FROM Sezionale WHERE IDFiliale=" & TheApp.Branch & " AND IDRegistroIva=8"
        
        .Fill
    End With
    
    With Me.cboVettore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT * FROM Vettore ORDER BY Vettore"
        
        .Fill
    End With

    With Me.cboTipoTrasporto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDTipoSpedizione"
        .DisplayField = "TipoSpedizione"
        .SQL = "SELECT * FROM TipoSpedizione ORDER BY TipoSpedizione"
        .Fill
    End With
    
    With Me.cboVettoreSuccessivo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT * FROM Vettore ORDER BY Vettore"
        .Fill
    End With

    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica  "
        .SQL = .SQL & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
    
    With Me.cboIvaCliente
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva"
        .Fill
    End With
End Sub
Private Function GET_IDPAGAMENTO_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDPagamentoDefault FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & Me.cdCliente.KeyFieldID
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    sSQL = "SELECT IDPagamentoDefault FROM PersonalizzazionePerFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_IDPAGAMENTO_DEFAULT = 0
    Else
        GET_IDPAGAMENTO_DEFAULT = fnNotNullN(rs!IDPagamentoDefault)
    End If
    
    
Else
    GET_IDPAGAMENTO_DEFAULT = fnNotNullN(rs!IDPagamentoDefault)
End If


rs.CloseResultset
Set rs = Nothing

End Function


Public Function GET_IDSEZIONALE_ORDINE(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDSezionale "
    sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
    sSQL = sSQL & " WHERE IDFiliale = " & TheApp.Branch
    sSQL = sSQL & " AND IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_IDSEZIONALE_ORDINE = fnNotNullN(rs!IDSezionale)
    Else
        GET_IDSEZIONALE_ORDINE = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Function


Private Sub txtDataOrdine_LostFocus()
    If Me.txtDataOrdine.Value = 0 Then Me.txtDataOrdine.Value = Date
    cboSezionale_Click
End Sub
Private Function FN_CREA_ORDINE() As String
On Error GoTo ERR_FN_CREA_ORDINE
Dim ObjDoc As DmtDocs.cDocument
Dim cDefault As Collection


If Not (ObjDoc Is Nothing) Then
    Set ObjDoc = Nothing
End If

Set ObjDoc = New DmtDocs.cDocument
 
    With ObjDoc
        Set .Connection = Cn
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = GET_ATTIVITA_AZIENDA
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto 15
        .IDFunzione = 128
        
        .IDEsercizio = fncEsercizio(Me.txtDataOrdine.Text)
        .IDSezionale = Me.cboSezionale.CurrentID
        .IDTipoAnagrafica = 2
        .IDUtente = TheApp.IDUser
        .Descrizione = "Ordine da cliente"
        .DataEmissione = Me.txtDataOrdine.Text
        .Numero = Me.txtNumeroOrdine.Value
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto 15
        Else
            .ClearValues
        End If
        
        ObjDoc.Tables("ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
        
        .Field "Link_Doc_magazzino", frmMain.CboMagazzinoVend.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Doc_sezionale", ObjDoc.IDSezionale, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Val_valuta", 9, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_data", ObjDoc.DataEmissione, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_numero", Me.txtNumeroOrdine.Value, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_data_prevista_evasione", Me.txtDataPartenza.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_data_presso_nom", Me.txtDataOrdineCliente.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_numero_presso_nom", Me.txtNumeroOrdineCliente.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_annotazioni_variazio", Mid(Me.txtAnnotazioniOrdine.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_POAnnotazioniInterna", Mid(Me.txtAnnotazioniInterna.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_PODescrizioneCorpoDocEv", Mid(Me.txtDescrizioneRigaDoc.Text, 1, 250), "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_POIDLuogoPresaMerce", Me.cboLuogoPresaMerce.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_PODataArrivoMerceLuogo", Me.txtDataArrivoMerceL.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_POOraArrivoMerceLuogo", Me.txtOraArrivoMerceL.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)


        .Field "RV_POIDTrasportatoreSuccessivo", Me.cboVettoreSuccessivo.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        .ReadDataFromCliFo Me.cdCliente.KeyFieldID
        
        .Field "Link_Doc_pagamento", Me.cboPagamento.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_ordine_chiuso", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Nom_lettera_intento", Me.txtIDLetteraIntento.Value, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Nom_IVA", Me.cboIvaCliente.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        .ReadDataFromCliFoSite Me.cboDestinazione.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_PODataArrivoMerce", Me.txtDataArrivoMerce.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_POOraArrivoMerce", Me.txtOraArrivoMerce.Text, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        
        .Field "Link_doc_spedizione", Me.cboTipoTrasporto.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        If Me.cboTipoTrasporto.CurrentID = 3 Then
            If Me.cboVettore.CurrentID > 0 Then
                .ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
            End If
        End If
        
        SetDefault
    
        Set ObjDoc.Scadenze = Nothing
        ObjDoc.PerformDocument Nothing
        VarNumeroDoc = "0"
        VarNumeroDoc = ObjDoc.Insert
        ObjDoc.Update
        If VarNumeroDoc > 0 Then
            
            LINK_ORDINE_RIF = fncTrovaDocumento(VarNumeroDoc)

        Else
            LINK_ORDINE_RIF = 0
            
        End If
        
    End With
    
    Set ObjDoc = Nothing
   
 Exit Function
ERR_FN_CREA_ORDINE:
 
    MsgBox Err.Description, vbCritical, "FN_CREA_ORDINE"
    LINK_ORDINE_RIF = 0
End Function
Private Sub SetDefault()
    'Imposta i default per i campi relativi alla valuta corrente
    'al tipo di arrotondamento, alle spese e ai bolli
    'Questi default sono obbligatori per il calcolo del documento
    Set cDefault = New Collection
    'Valore di arrotondamento per la valuta corrente
    cDefault.Add 1, "Val_arrotondamento"
    'Tipo di arrotondomento per la valuta corrente
    cDefault.Add 1, "Link_Val_tipo_arrotondamento"
    'ID della valuta corrente
    cDefault.Add 0, "Link_Val_valuta"
    'Spesse incassi in percentuale
    cDefault.Add 0, "Nom_spese_incassi_perc"
    'Importo del bollo
    cDefault.Add 0, "Nom_bollo_esente"
    'Importo limite per il pagamento del bollo
    cDefault.Add 0, "Nom_bollo_esente_limite"
    'ID del contratto bancario azienda
    cDefault.Add 0, "Link_Doc_contratto_bancario_az"
    'ID della natura delle scadenze
    cDefault.Add 0, "IDNaturaScadenza"
End Sub
Private Function fncTrovaDocumento(NumeroDoc As String) As Long
On Error GoTo ERR_fncTrovaDocumento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto From Oggetto WHERE ("
sSQL = sSQL & "(IDTipoOggetto=15) "
sSQL = sSQL & "AND (Numero=" & fnNormString(NumeroDoc) & ") "
sSQL = sSQL & "AND (DataEmissione=" & fnNormDate(Me.txtDataOrdine.Text) & ") "
sSQL = sSQL & "AND (IDAzienda=" & TheApp.IDFirm & ")) "
sSQL = sSQL & "ORDER BY IDOggetto DESC"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDocumento = rs!IDOggetto
Else
    fncTrovaDocumento = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_fncTrovaDocumento:
    MsgBox Err.Description, vbCritical, "Impossibile stampare"
    fncTrovaDocumento = 0
End Function
Private Function GET_CONTROLLO_FIDO_CLIENTE() As Boolean
'Consideriamo di aver istanziato un oggetto oServices di tipo DmtCliFu.Services
Dim oServices As DMTCliFu.Services

Set oServices = New DMTCliFu.Services

'Passo una connessione al database di DMT (di tipo DmtOleDbLib.adoConnection)

Set oServices.Connection = Cn
'Identificativo dell'Azienda

oServices.IDAzienda = TheApp.IDFirm
'Identificativo dell'anagrafica del cliente

oServices.IDAnagrafica = Me.cdCliente.KeyFieldID
'Totale del documento corrente

oServices.TotDocumento = 0
'IDPagamaneto del documento corrente

oServices.IDPagamento = Me.cboPagamento.CurrentID

'Identificativo univoco dell'oggetto del documento da cui leggere il totale netto a pagare nella valuta dell'azienda
'oServices.IDDocumento = oDoc.IDOggetto

'Tabella di testata del tipo di documento che si sta prendendo in considerazione per la determinazione del netto a pagare
'oServices.DocTableName = sTabellaTestata

GET_CONTROLLO_FIDO_CLIENTE = oServices.CheckFido

Set oServices = Nothing
End Function

Private Function GET_INFO_CLIENTE_FIDO_BLOCCO(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


''''''''''FIDO PER CLIENTE
sSQL = "SELECT BloccoEmissioneDoc, IDTipoBloccoFido, PasswordSbloccoFido, DataFineSbloccoFido "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_BLOCCO_CLIENTE = 0
    LINK_TIPO_FIDO_CLIENTE = 0
    PASSWORD_FIDO_CLIENTE = ""
    DATA_SBLOCCO_FIDO_CLIENTE = ""
    
Else
    LINK_BLOCCO_CLIENTE = fnNotNullN(rs!BloccoEmissioneDoc)
    LINK_TIPO_FIDO_CLIENTE = fnNotNullN(rs!IDTipoBloccoFido)
    PASSWORD_FIDO_CLIENTE = fnCryptString(fnNotNull(rs!PasswordSbloccoFido))
    DATA_SBLOCCO_FIDO_CLIENTE = fnNotNull(rs!DataFineSbloccoFido)
End If

rs.CloseResultset
Set rs = Nothing

'''''''''''FIDO PER AZIENDA
sSQL = "SELECT IDTipoBloccoFido, PasswordSbloccoFido "
sSQL = sSQL & "FROM ConfigurazioneGenerale "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_FIDO_AZIENDA = 0
    PASSWORD_FIDO_AZIENDA = ""
Else
    LINK_TIPO_FIDO_AZIENDA = fnNotNullN(rs!IDTipoBloccoFido)
    PASSWORD_FIDO_AZIENDA = fnCryptString(fnNotNull(rs!PasswordSbloccoFido))
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Function CHECK_ABILITAZIONE_DIAMANTE() As Boolean
Dim I As Integer
Dim swChk As DmtSwChk.SwCheck

Set swChk = New DmtSwChk.SwCheck

Set swChk.Connection = Cn.InternalConnection
swChk.SwComponentName = "MODBASE_"
I = swChk.CheckSwComponent

Select Case I
    Case 0
        CHECK_ABILITAZIONE_DIAMANTE = False
        MsgBox swChk.LastError, vbCritical, "Abilitazione programma"
    Case -1
        CHECK_ABILITAZIONE_DIAMANTE = True
    Case -2
        CHECK_ABILITAZIONE_DIAMANTE = False
        MsgBox swChk.LastError, vbCritical, "Abilitazione programma"
End Select

Set swChk = Nothing

End Function

Private Sub txtIDLetteraIntento_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If IsEmpty(Me.txtIDLetteraIntento.Value) Then Me.txtIDLetteraIntento.Value = 0

sSQL = "SELECT * FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & Me.txtIDLetteraIntento.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNLetteraIntento.Value = 0
    Me.txtDataLetteraIntento.Value = 0
    Me.lblLetteraIntento.ToolTipText = ""
Else
    Me.txtNLetteraIntento.Value = fnNotNullN(rs!Numero)
    Me.txtDataLetteraIntento.Value = fnNotNullN(rs!Data)
    Me.lblLetteraIntento.ToolTipText = "Prot. N° " & fnNotNull(rs!NumeroCliFor) & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_CONTROLLO_NUMERO_LETTERE_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDLetteraIntento) AS NumeroRecord "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = 0
Else
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_NUMERO_LETTERE_INTENTO"
End Function
Private Function GET_LINK_LETTERA_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_LINK_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLetteraIntento "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LETTERA_INTENTO = 0
Else
    GET_LINK_LETTERA_INTENTO = fnNotNullN(rs!IDLetteraIntento)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_LETTERA_INTENTO"
End Function
Private Function GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento As Long, IDIvaCliente As Long) As Long
On Error GoTo ERR_GET_LINK_IVA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
Else
    If fnNotNullN(rs!IDIva) > 0 Then
        GET_LINK_IVA_LETTERA_INTENTO = fnNotNullN(rs!IDIva)
    Else
        GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_IVA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_IVA_LETTERA_INTENTO"
End Function

Private Function GET_LINK_IVA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE = 0
Else
    GET_LINK_IVA_CLIENTE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function

End Function

