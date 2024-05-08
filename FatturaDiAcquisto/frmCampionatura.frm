VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.7#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmCampionatura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMPIONATURA"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15495
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
   ScaleHeight     =   7500
   ScaleWidth      =   15495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminaArticolo 
      Caption         =   "ELIMINA ARTICOLO"
      Height          =   495
      Left            =   13680
      TabIndex        =   33
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ARTICOLI DERIVATI"
      Height          =   495
      Left            =   13680
      TabIndex        =   31
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAggiungi 
      Caption         =   "AGGIUNGI ARTICOLO"
      Height          =   495
      Left            =   11880
      TabIndex        =   30
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAggiorna 
      Caption         =   "RIPRISTINA"
      Height          =   495
      Left            =   11880
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "ELIMINA"
      Height          =   495
      Left            =   11760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame FraTotali 
      Caption         =   "RIEPILOGO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin DMTEDITNUMLib.dmtNumber txtPercCamp 
         Height          =   315
         Left            =   2160
         TabIndex        =   23
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleInsCamp 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDiffCamp 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTotaleImponibile 
         Height          =   315
         Left            =   4920
         TabIndex        =   26
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
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
      Begin DMTEDITNUMLib.dmtNumber txtTotaleIVA 
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Top             =   600
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
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
      Begin DMTEDITNUMLib.dmtNumber txtTotaleDocumento 
         Height          =   315
         Left            =   4920
         TabIndex        =   28
         Top             =   960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
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
      Begin VB.Label Label10 
         Caption         =   "Totale "
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "I.V.A."
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   3600
         X2              =   3600
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "% Camp."
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Differenza"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Imponibile"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Campionatura"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13680
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8916
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
   Begin VB.Frame FraTesta 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkBloccata 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloccata"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaCamp 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   65535
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTDATETIMELib.dmtDate txtDataCampionatura 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCampionatura 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtAnnoCampionatura 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaConf 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
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
      Begin DMTEDITNUMLib.dmtNumber txtTotalePercCamp 
         Height          =   315
         Left            =   3480
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "% Campionata"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Q.tà conferita"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Q.tà campionata"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Numerazione"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Data"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCampionatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private LINK_LISTINO_ACQ As Long

Private LINK_CAMPIONATURA As Long

Private Camp_Change As Boolean

Private CONFERMA_SALVATAGGIO As Boolean

Private Sub cmdAggiorna_Click()
On Error GoTo ERR_cmdAggiorna_Click
    'GET_PRELEVA_DATI_RIGHE
    
    CARICA_ARTICOLI_CAMPIONATI Link_RigaConferimento, TheApp.IDUser, True
    
    GET_GRIGLIA
    
    TOTALI_CAMPIONATURA
Exit Sub
ERR_cmdAggiorna_Click:
    MsgBox Err.Description, vbCritical, "cmdAggiorna_Click"
    
End Sub

Private Sub cmdAggiungi_Click()
    frmAddArticoloFiglio.Show vbModal
    
    GET_GRIGLIA
    
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
    
    CONFERMA_SALVATAGGIO = SALVATAGGIO_TESTA
    
    If CONFERMA_SALVATAGGIO = True Then
    
        SALVATAGGIO_RIGHE
        
        GET_PRELEVA_DATI_TESTA
        
        'GET_PRELEVA_DATI_RIGHE
        CARICA_ARTICOLI_CAMPIONATI Link_RigaConferimento, TheApp.IDUser, True
        GET_GRIGLIA
        
        TOTALI_CAMPIONATURA
        
        Camp_Change = False
        
        GET_CHANGE
        
    End If
Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Private Sub cmdElimina_Click()
On Error GoTo ERR_cmdElimina_Click
Dim Testo As String
Dim sSQL As String

If LINK_CAMPIONATURA = 0 Then Exit Sub

Testo = "Sei sicuro di eliminare la campionatura di questo conferimento"

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione campionatura") = vbNo Then Exit Sub


sSQL = "DELETE FROM RV_POCampionaturaRighe "
sSQL = sSQL & " WHERE IDRV_POCampionatura=" & LINK_CAMPIONATURA
Cn.Execute sSQL

sSQL = "DELETE FROM RV_POCampionatura "
sSQL = sSQL & " WHERE IDRV_POCampionatura=" & LINK_CAMPIONATURA
Cn.Execute sSQL

GET_PRELEVA_DATI_TESTA

GET_PRELEVA_DATI_RIGHE
    
GET_GRIGLIA

TOTALI_CAMPIONATURA

Camp_Change = False
Exit Sub
ERR_cmdElimina_Click:
    MsgBox Err.Description, vbCritical, "Eliminazione campionatura"
End Sub

Private Sub cmdEliminaArticolo_Click()
Dim sSQL As String
Dim Testo As String

Testo = "Sei sicuro di voler eliminare la riga di campionatura"
If MsgBox(Testo, vbQuestion + vbYesNo, "Elimazione riga di campionatura") = vbNo Then Exit Sub

If ((fnNotNullN(Me.Griglia.AllColumns("IDArticoloQuadratura").Value) > 0) And (fnNotNullN(Me.Griglia.AllColumns("IDCollegamento").Value) > 0)) Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Questa riga di quadratura è stata generata da un'elaborazione" & vbCrLf
    Testo = Testo & "Sei sicuro di voler continuare?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Elimazione riga di campionatura") = vbNo Then Exit Sub
End If

If ((fnNotNullN(Me.Griglia.AllColumns("IDArticoloQuadratura").Value) = 0) And (fnNotNullN(Me.Griglia.AllColumns("IDCollegamento").Value) > 0)) Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Questa riga di invenduto è stata generata da un'elaborazione" & vbCrLf
    Testo = Testo & "Sei sicuro di voler continuare?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Elimazione riga di campionatura") = vbNo Then Exit Sub
End If


If ((fnNotNullN(Me.Griglia.AllColumns("IDArticoloQuadratura").Value) = 0) And (fnNotNullN(Me.Griglia.AllColumns("IDCollegamento").Value) = 0)) Then
    If GET_CONTROLLO_ESISTENZA_COLLEGAMENTI(fnNotNullN(Me.Griglia.AllColumns("IDRV_POCampionaturaRighe").Value)) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Questa riga di campionatura è collegata ad righe di campionatura precedentemente elaborate" & vbCrLf
        Testo = Testo & "Sei sicuro di voler continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Elimazione riga di campionatura") = vbNo Then Exit Sub
    End If
End If


sSQL = "DELETE FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDRV_POCampionaturaRigheTMP=" & fnNotNullN(Me.Griglia("IDRV_POCampionaturaRigheTMP").Value)
Cn.Execute sSQL

GET_GRIGLIA

TOTALI_CAMPIONATURA

End Sub
Private Function GET_CONTROLLO_ESISTENZA_COLLEGAMENTI(IDRigaCampionatura As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCampionaturaRighe FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDCollegamento=" & IDRigaCampionatura

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_COLLEGAMENTI = False
Else
    GET_CONTROLLO_ESISTENZA_COLLEGAMENTI = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub Command1_Click()
Dim Testo As String

    Testo = "Con questa procedura verranno importati tutti gli articoli derivati dell'articolo conferito" & vbCrLf
    Testo = Testo & "Vuoi continuare?"

    If MsgBox(Testo, vbQuestion + vbYesNo, "Inserimento articoli derivati") = vbNo Then Exit Sub
    
    GET_PRELEVA_DATI_RIGHE
    
    GET_GRIGLIA
    
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE32)
    ParametroListinoCampionatura
    Me.txtQtaConf.Value = frmMain.txtQta_UM.Value

    Me.Caption = "Campionatura dell'articolo " & frmMain.CDArticolo.Code & " - " & frmMain.CDArticolo.Description
    
    GET_PRELEVA_DATI_TESTA
    
    'GET_PRELEVA_DATI_RIGHE
    
    CARICA_ARTICOLI_CAMPIONATI Link_RigaConferimento, TheApp.IDUser, True
    
    GET_GRIGLIA
    
    TOTALI_CAMPIONATURA
    
    Camp_Change = False
    
    Me.txtQtaCamp.Tag = Me.txtQtaCamp.Value

End Sub
Private Sub GET_PRELEVA_DATI_RIGHE()
On Error GoTo ERR_GET_PRELEVA_DATI_RIGHE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsArt As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long
Dim Unita_progresso As Double


''''''CONTEGGIO RECORD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT Count(IDRV_POArticoloFiglio) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POArticoloFiglio "
sSQL = sSQL & "WHERE IDArticolo=" & LINK_ARTICOLO_CONFERITO

Set rsArt = Cn.OpenResultset(sSQL)

If rsArt.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rsArt!NumeroRecord)
End If

rsArt.CloseResultset
Set rsArt = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If NumeroRecord = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
Me.ProgressBar1.Visible = True

Unita_progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)



sSQL = "SELECT * FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT RV_POArticoloFiglio.IDRV_POArticoloFiglio, RV_POArticoloFiglio.IDArticoloFiglio, RV_POArticoloFiglio.IDArticolo, RV_POArticoloFiglio.PesoPerOrdinamento, "
sSQL = sSQL & "RV_POArticoloFiglio.IDRV_POTipoLavorazione, RV_POTipoLavorazione.TipoLavorazione, Articolo.CodiceArticolo, Articolo.Articolo,"
sSQL = sSQL & "RV_POTipoCategoria.IDRV_POTipoCategoria , RV_POTipoCategoria.TipoCategoria, RV_POCalibro.IDRV_POCalibro, RV_POCalibro.Calibro,  Articolo.DescrizioneArticoloRidotta "
sSQL = sSQL & "FROM RV_POTipoCategoria RIGHT OUTER JOIN "
sSQL = sSQL & "Articolo ON RV_POTipoCategoria.IDRV_POTipoCategoria = Articolo.RV_POIDTipoCategoria LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCalibro ON Articolo.RV_POIDCalibro = RV_POCalibro.IDRV_POCalibro RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POArticoloFiglio ON Articolo.IDArticolo = RV_POArticoloFiglio.IDArticoloFiglio LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoLavorazione ON RV_POArticoloFiglio.IDRV_POTipoLavorazione = RV_POTipoLavorazione.IDRV_POTipoLavorazione "
sSQL = sSQL & "WHERE RV_POArticoloFiglio.IDArticolo=" & LINK_ARTICOLO_CONFERITO
sSQL = sSQL & " ORDER BY RV_POArticoloFiglio.PesoPerOrdinamento, Articolo.CodiceArticolo "

Set rsArt = Cn.OpenResultset(sSQL)

While Not rsArt.EOF
    If GET_ESISTENZA_ARTICOLO_DERIVATO(fnNotNullN(rsArt!IDArticoloFiglio), Link_RigaConferimento) = False Then
    
        rsNew.AddNew
            rsNew!IDRV_POCampionaturaRigheTMP = fnGetNewKey("RV_POCampionaturaRigheTMP", "IDRV_POCampionaturaRigheTMP")
            rsNew!IDUtente = TheApp.IDUser
            rsNew!IDRV_POCaricoMerceRighe = Link_RigaConferimento
            rsNew!IDArticolo = fnNotNullN(rsArt!IDArticoloFiglio)
            rsNew!QuantitaCampionata = 0
            rsNew!QuantitaDefinitiva = 0
            rsNew!ImportoUnitario = GET_PREZZO_ARTICOLO(LINK_LISTINO_ACQ, fnNotNullN(rsArt!IDArticoloFiglio))
            rsNew!ImportoNettoRiga = 0
            rsNew!ImportoImpostaRiga = 0
            rsNew!ImportoLordoRiga = 0
            rsNew!Invenduto = 0
            rsNew!DescrizioneAggiuntiva = ""
            rsNew!IDArticoloQuadratura = 0
            rsNew!IDCollegamento = 0
            rsNew!IDRV_POCampionaturaRighe = 0
            'GET_INFO_CAMPIONATURA rsNew, fnNotNullN(rsArt!IDArticoloFiglio), Link_RigaConferimento
            GET_INFO_ARTICOLO rsNew, fnNotNullN(rsArt!IDArticoloFiglio)
            GET_INFO_ARTICOLO_QUAD rsNew, fnNotNullN(rsNew!IDArticoloQuadratura)
        rsNew.Update
    End If
    If (Me.ProgressBar1.Value + Unita_progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_progresso
    End If
    
    DoEvents
    
rsArt.MoveNext
Wend

rsArt.CloseResultset
Set rsArt = Nothing

Exit Sub
ERR_GET_PRELEVA_DATI_RIGHE:
    MsgBox Err.Description, vbCritical, "GET_PRELEVA_DATI_RIGHE"
End Sub

Private Sub GET_INFO_CAMPIONATURA(rsTmp As ADODB.Recordset, IDArticolo As Long, IDRigaConferimento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsTmp!QuantitaCampionata = 0
    rsTmp!QuantitaDefinitiva = 0
    rsTmp!ImportoUnitario = 0 'GET_PREZZO_ARTICOLO(LINK_LISTINO_ACQ, IDArticolo)
    rsTmp!ImportoNettoRiga = 0
    rsTmp!ImportoImpostaRiga = 0
    rsTmp!ImportoLordoRiga = 0
    
Else
    rsTmp!QuantitaCampionata = fnNotNullN(rs!QuantitaCampionata)
    rsTmp!QuantitaDefinitiva = fnNotNullN(rs!QuantitaDefinitiva)
    rsTmp!ImportoUnitario = fnNotNullN(rs!ImportoUnitario)
    rsTmp!ImportoNettoRiga = fnNotNullN(rs!ImportoNettoRiga)
    rsTmp!ImportoImpostaRiga = fnNotNullN(rs!ImportoImpostaRiga)
    rsTmp!ImportoLordoRiga = fnNotNullN(rs!ImportoLordoRiga)

End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_INFO_ARTICOLO(rsTmp As ADODB.Recordset, IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM IERepArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsTmp!CodiceArticolo = ""
    rsTmp!Articolo = ""
    rsTmp!IDUnitaDiMisura = 0
    rsTmp!UnitaDiMisura = ""
    rsTmp!IDIva = 0
    rsTmp!AliquotaIva = 0
Else
    rsTmp!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
    rsTmp!Articolo = fnNotNull(rs!Articolo)
    rsTmp!IDUnitaDiMisura = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    rsTmp!UnitaDiMisura = fnNotNull(rs!UnitaDiMisuraAcquisto)
    rsTmp!IDIva = fnNotNullN(rs!IDIvaAcquisto)
    rsTmp!AliquotaIva = GET_ALIQUOTA_IVA(fnNotNullN(rs!IDIvaAcquisto))
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_INFO_ARTICOLO_QUAD(rsTmp As ADODB.Recordset, IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM IERepArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rsTmp!CodiceArticoloQuadratura = ""
    rsTmp!ArticoloQuadratura = ""
Else
    rsTmp!CodiceArticoloQuadratura = fnNotNull(rs!CodiceArticolo)
    rsTmp!ArticoloQuadratura = fnNotNull(rs!Articolo)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub ParametroListinoCampionatura()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoCampionatura FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_LISTINO_ACQ = fnNotNullN(rs!IDListinoCampionatura)
Else
    LINK_LISTINO_ACQ = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_PREZZO_ARTICOLO(IDListino As Long, IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT PrezzoNettoIVA "
sSQL = sSQL & "FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDListino=" & IDListino
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_ARTICOLO = 0
Else
    GET_PREZZO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
Else
    GET_ALIQUOTA_IVA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_GRIGLIA()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    sSQL = "SELECT * FROM RV_POCampionaturaRigheTMP "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & Link_RigaConferimento

    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic
    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDRV_POCampionaturaRigheTMP", "IDRV_POCampionaturaRigheTMP", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgInteger, False, 500, dgAlignleft
                
                .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Articolo", "Articolo", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "UnitaDiMisura", "U.M.", dgchar, True, 1700, dgAlignleft
                Set cl = .ColumnsHeader.Add("QuantitaCampionata", "Q.tà camp.", dgDouble, True, 1800, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("QuantitaDefinitiva", "Q.tà conf.", dgDouble, True, 1800, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                
                Set cl = .ColumnsHeader.Add("ImportoUnitario", "Imp. Uni.", dgDouble, True, 1800, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "IDIva", "IDIva", dgNumeric, False, 500, dgAlignleft
                Set cl = .ColumnsHeader.Add("AliquotaIva", "% I.V.A.", dgDouble, False, 1100, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                
                
                Set cl = .ColumnsHeader.Add("ImportoNettoRiga", "Netto riga", dgDouble, True, 1800, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("ImportoImpostaRiga", "I.V.A.", dgDouble, True, 1800, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("ImportoLordoRiga", "Totale riga", dgDouble, True, 1800, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "Invenduto", "Invenduto", dgBoolean, False, 1500, dgAligncenter
                .ColumnsHeader.Add "DescrizioneAggiuntiva", "Annotazioni", dgchar, False, 2500, dgAlignleft
                .ColumnsHeader.Add "IDArticoloQuadratura", "IDArticoloQuadratura", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticoloQuadratura", "Codice Art. quad.", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "ArticoloQuadratura", "Articolo quad.", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDCollegamento", "IDCollegamento", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDRV_POCaricoMerceRigheOLD", "IDRV_POCaricoMerceRigheOLD", dgInteger, False, 500, dgAlignleft
        Set .Recordset = rsGriglia
        .LoadUserSettings
        .Refresh
        
    End With
    
    'Cn.CursorLocation = OLDCursor

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Testo As String


If Camp_Change = True Then
    Testo = "Vuoi salvare la campionatura?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Uscita campionatura") = vbYes Then
        cmdConferma_Click
        If CONFERMA_SALVATAGGIO = False Then
            Testo = "Il documento non è stato salvato" & vbCrLf
            Testo = Testo & "Vuoi uscire comunque dalla campionatura?"
            If MsgBox(Testo, vbQuestion + vbYesNo, "Uscita campionatura") = vbNo Then Cancel = 1
        End If
    End If
End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsGriglia Is Nothing) Then
        rsGriglia.Close
        Set rsGriglia = Nothing
    End If
End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As dmtgridctl.dgColumnHeader, ByVal Value As Variant)
    If Me.txtQtaCamp.Value > 0 Then
        rsGriglia("QuantitaDefinitiva").Value = (fnNotNullN(rsGriglia("QuantitaCampionata").Value) / Me.txtQtaCamp.Value) * Me.txtQtaConf.Value
        rsGriglia("ImportoNettoRiga").Value = fnNotNullN(rsGriglia("ImportoUnitario").Value) * fnNotNullN(rsGriglia("QuantitaDefinitiva").Value)
        rsGriglia("ImportoImpostaRiga").Value = (fnNotNullN(rsGriglia("ImportoNettoRiga").Value) / 100) * fnNotNullN(rsGriglia("AliquotaIva").Value)
        rsGriglia("ImportoLordoRiga").Value = fnNotNullN(rsGriglia("ImportoNettoRiga").Value) + fnNotNullN(rsGriglia("ImportoImpostaRiga").Value)
    End If
    
    Me.Griglia.Refresh
    rsGriglia.UpdateBatch
    
    TOTALI_CAMPIONATURA
    
    Camp_Change = True
    
    GET_CHANGE
    
End Sub
Private Sub TOTALI_CAMPIONATURA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtTotalePercCamp.Value = ((Me.txtQtaCamp.Value / Me.txtQtaConf.Value) * 100)

sSQL = "SELECT SUM(QuantitaCampionata) AS SommaQuantitaCampionata, "
sSQL = sSQL & "SUM(ImportoNettoRiga) AS SommaImportoNettoRiga, "
sSQL = sSQL & "SUM(ImportoImpostaRiga) AS SommaImportoImpostaRiga, "
sSQL = sSQL & "SUM(ImportoLordoRiga) AS SommaImportoLordoRiga "
sSQL = sSQL & "FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotaleInsCamp.Value = 0
    Me.txtTotaleImponibile.Value = 0
    Me.txtTotaleIVA.Value = 0
    Me.txtTotaleDocumento.Value = 0
Else
    Me.txtTotaleInsCamp.Value = fnNotNullN(rs!SommaQuantitaCampionata)
    Me.txtTotaleImponibile.Value = fnNotNullN(rs!SommaImportoNettoRiga)
    Me.txtTotaleIVA.Value = fnNotNullN(rs!SommaImportoImpostaRiga)
    Me.txtTotaleDocumento.Value = fnNotNullN(rs!SommaImportoLordoRiga)
End If

rs.CloseResultset
Set rs = Nothing

Me.txtDiffCamp.Value = Me.txtQtaCamp.Value - Me.txtTotaleInsCamp.Value
If Me.txtQtaCamp.Value > 0 Then
    Me.txtPercCamp.Value = ((Me.txtTotaleInsCamp.Value / Me.txtQtaCamp.Value) * 100)
Else
    Me.txtPercCamp.Value = 0
End If

End Sub

Private Sub txtDataCampionatura_LostFocus()
    If Me.txtDataCampionatura.Value = 0 Then Me.txtDataCampionatura.Value = Date
    Me.txtAnnoCampionatura.Value = Year(Me.txtDataCampionatura.Text)
    Me.txtNumeroCampionatura.Value = GET_NUMERO_DOCUMENTO(Me.txtAnnoCampionatura.Value)

    Camp_Change = True
    
End Sub

Private Sub txtQtaCamp_Change()
    TOTALI_CAMPIONATURA
    Camp_Change = True
    GET_CHANGE
    
End Sub

Private Sub GET_CHANGE()
If Camp_Change = True Then
    cmdConferma.Enabled = True
Else
    cmdConferma.Enabled = False
End If

End Sub
Private Function GET_PRELEVA_DATI_TESTA()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT * FROM RV_POCampionatura "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_CAMPIONATURA = 0
    Me.txtDataCampionatura.Value = Date
    Me.txtAnnoCampionatura.Value = Year(Date)
    Me.txtNumeroCampionatura.Value = GET_NUMERO_DOCUMENTO(Me.txtAnnoCampionatura.Value)
    Me.txtQtaCamp.Value = 0
    Me.chkBloccata.Value = vbUnchecked
Else
    LINK_CAMPIONATURA = fnNotNullN(rs!IDRV_POCampionatura)
    Me.txtDataCampionatura.Value = fnNotNullN(rs!DataCampionatura)
    Me.txtAnnoCampionatura.Value = fnNotNullN(rs!AnnoCampionatura)
    Me.txtNumeroCampionatura.Value = fnNotNullN(rs!NumeroCampionatura)
    Me.txtQtaCamp.Value = fnNotNullN(rs!QuantitaCampionata)
    Me.chkBloccata.Value = Abs(fnNotNullN(rs!Bloccata))
End If

rs.CloseResultset
Set rs = Nothing

Camp_Change = False
GET_CHANGE
End Function
Private Function SALVATAGGIO_TESTA() As Boolean
On Error GoTo ERR_SALVATAGGIO_TESTA
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Testo As String

SALVATAGGIO_TESTA = False

    If Me.txtDataCampionatura.Value = 0 Then
        MsgBox "La data di campionatura deve essere inserita", vbCritical, "Salvataggio"
        Exit Function
    End If
    If Me.txtQtaCamp.Value = 0 Then
        MsgBox "La quantità cmpionata non può essere uguale a zero", vbCritical, "Salvataggio"
        Exit Function
    End If
    If Me.txtQtaCamp.Value > Me.txtQtaConf.Value Then
        MsgBox "La quantità cmpionata non può essere maggiore della quantità conferita", vbCritical, "Salvataggio"
        Exit Function
    End If
    
    If Me.txtDiffCamp.Value <> 0 Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "La differenza di campionatura non può essere diversa da zero" & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio") = vbNo Then
            Exit Function
        End If
    End If


sSQL = "SELECT * FROM RV_POCampionatura "
sSQL = sSQL & "WHERE IDRV_POCampionatura=" & LINK_CAMPIONATURA

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    LINK_CAMPIONATURA = fnGetNewKey("RV_POCampionatura", "IDRV_POCampionatura")
    Me.txtNumeroCampionatura.Value = GET_NUMERO_DOCUMENTO(Me.txtAnnoCampionatura.Value)
    
End If
    rs!IDRV_POCampionatura = LINK_CAMPIONATURA
    rs!IDRV_POCaricoMerceRighe = Link_RigaConferimento
    rs!DataCampionatura = Me.txtDataCampionatura.Text
    rs!AnnoCampionatura = Me.txtAnnoCampionatura.Value
    rs!NumeroCampionatura = Me.txtNumeroCampionatura.Value
    rs!QuantitaCampionata = Me.txtQtaCamp.Value
    rs!IDUtente = TheApp.IDUser
    rs!Bloccata = Abs(Me.chkBloccata.Value)
rs.Update

rs.Close
Set rs = Nothing

SALVATAGGIO_TESTA = True
Exit Function
ERR_SALVATAGGIO_TESTA:
    MsgBox Err.Description, vbCritical, "Campionatura"
    SALVATAGGIO_TESTA = False
End Function
Private Sub SALVATAGGIO_RIGHE()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsTmp As DmtOleDbLib.adoResultset


'ELIMINAZIONE DATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCampionatura=" & LINK_CAMPIONATURA
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCampionatura=" & LINK_CAMPIONATURA

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & Link_RigaConferimento
sSQL = sSQL & " AND QuantitaCampionata<>0"
sSQL = sSQL & " ORDER BY IDCollegamento"
Set rsTmp = Cn.OpenResultset(sSQL)

While Not rsTmp.EOF
    rs.AddNew
        rs!IDRV_POCampionaturaRighe = fnGetNewKey("RV_POCampionaturaRighe", "IDRV_POCampionaturaRighe")
        rs!IDRV_POCampionatura = LINK_CAMPIONATURA
        rs!IDRV_POCaricoMerceRighe = Link_RigaConferimento
        rs!IDArticolo = fnNotNullN(rsTmp!IDArticolo)
        rs!IDUnitaDiMisura = fnNotNullN(rsTmp!IDUnitaDiMisura)
        rs!QuantitaCampionata = fnNotNullN(rsTmp!QuantitaCampionata)
        rs!QuantitaDefinitiva = fnNotNullN(rsTmp!QuantitaDefinitiva)
        rs!ImportoUnitario = fnNotNullN(rsTmp!ImportoUnitario)
        rs!ImportoNettoRiga = fnNotNullN(rsTmp!ImportoNettoRiga)
        rs!IDIva = fnNotNullN(rsTmp!IDIva)
        rs!ImportoImpostaRiga = fnNotNullN(rsTmp!ImportoImpostaRiga)
        rs!ImportoLordoRiga = fnNotNullN(rsTmp!ImportoLordoRiga)
        rs!Invenduto = fnNotNullN(rsTmp!Invenduto)
        rs!DescrizioneAggiuntiva = fnNotNull(rsTmp!DescrizioneAggiuntiva)
        rs!IDArticoloQuadratura = fnNotNullN(rsTmp!IDArticoloQuadratura)
        rs!IDRV_POCampionaturaRigheOLD = fnNotNullN(rsTmp!IDRV_POCampionaturaRighe)
        rs!IDCollegamento = GET_LINK_COLLEGAMENTO(fnNotNullN(rsTmp!IDCollegamento))
    rs.Update
rsTmp.MoveNext
Wend

rs.Close
Set rs = Nothing

rsTmp.CloseResultset
Set rsTmp = Nothing
End Sub
Private Function GET_NUMERO_DOCUMENTO(Anno As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroCampionatura) AS NumeroDocumento "
sSQL = sSQL & "FROM RV_POCampionatura "
sSQL = sSQL & "WHERE AnnoCampionatura=" & Anno

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_DOCUMENTO = 1
Else
    GET_NUMERO_DOCUMENTO = fnNotNullN(rs!NumeroDocumento) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtQtaCamp_LostFocus()
    If Me.txtQtaCamp.Value = 0 Then Exit Sub
    If Me.txtQtaCamp.Tag = Me.txtQtaCamp.Value Then Exit Sub
    
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        rsGriglia.MoveFirst
        While Not rsGriglia.EOF
            rsGriglia("QuantitaDefinitiva").Value = (fnNotNullN(rsGriglia("QuantitaCampionata").Value) / Me.txtQtaCamp.Value) * Me.txtQtaConf.Value
            rsGriglia("ImportoNettoRiga").Value = fnNotNullN(rsGriglia("ImportoUnitario").Value) * fnNotNullN(rsGriglia("QuantitaDefinitiva").Value)
            rsGriglia("ImportoImpostaRiga").Value = (fnNotNullN(rsGriglia("ImportoNettoRiga").Value) / 100) * fnNotNullN(rsGriglia("AliquotaIva").Value)
            rsGriglia("ImportoLordoRiga").Value = fnNotNullN(rsGriglia("ImportoNettoRiga").Value) + fnNotNullN(rsGriglia("ImportoImpostaRiga").Value)
            rsGriglia.UpdateBatch
        rsGriglia.MoveNext
        Wend
        
        Me.Griglia.Refresh
        
        TOTALI_CAMPIONATURA

        Camp_Change = True
        GET_CHANGE
        Me.txtQtaCamp.Tag = Me.txtQtaCamp.Value
    End If
End Sub
Private Sub CARICA_ARTICOLI_CAMPIONATI(IDRigaConferimento As Long, IDUtente As Long, Elimina As Boolean)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset


If Elimina = True Then
    '''''ELIMINAZIONE DATI DALLA TABELLA TEMPORANEA'''''''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POCampionaturaRigheTMP "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If
sSQL = "SELECT * FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POCampionaturaRigheTMP = fnGetNewKey("RV_POCampionaturaRigheTMP", "IDRV_POCampionaturaRigheTMP")
        rsNew!IDUtente = IDUtente
        rsNew!IDRV_POCaricoMerceRighe = IDRigaConferimento
        rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsNew!QuantitaCampionata = fnNotNullN(rs!QuantitaCampionata)
        rsNew!QuantitaDefinitiva = fnNotNullN(rs!QuantitaDefinitiva)
        rsNew!ImportoUnitario = fnNotNullN(rs!ImportoUnitario)
        rsNew!ImportoNettoRiga = fnNotNullN(rs!ImportoNettoRiga)
        rsNew!ImportoImpostaRiga = fnNotNullN(rs!ImportoImpostaRiga)
        rsNew!ImportoLordoRiga = fnNotNullN(rs!ImportoLordoRiga)
        rsNew!Invenduto = fnNotNullN(rs!Invenduto)
        rsNew!DescrizioneAggiuntiva = fnNotNull(rs!DescrizioneAggiuntiva)
        rsNew!IDArticoloQuadratura = fnNotNullN(rs!IDArticoloQuadratura)
        rsNew!IDCollegamento = fnNotNullN(rs!IDCollegamento)
        rsNew!IDRV_POCampionaturaRighe = fnNotNullN(rs!IDRV_POCampionaturaRighe)
        
        GET_INFO_ARTICOLO rsNew, fnNotNullN(rs!IDArticolo)
        
        GET_INFO_ARTICOLO_QUAD rsNew, fnNotNullN(rs!IDArticoloQuadratura)
        
    rsNew.Update
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_ESISTENZA_ARTICOLO_DERIVATO(IDArticolo As Long, IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POCampionaturaRigheTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & "  AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ARTICOLO_DERIVATO = False
Else
    GET_ESISTENZA_ARTICOLO_DERIVATO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_LINK_COLLEGAMENTO(IDCollegamento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCampionaturaRighe FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCampionaturaRigheOLD=" & IDCollegamento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_COLLEGAMENTO = 0
Else
    GET_LINK_COLLEGAMENTO = fnNotNullN(rs!IDRV_POCampionaturaRighe)
End If

rs.CloseResultset
Set rs = Nothing
End Function
