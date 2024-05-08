VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmImputazionePrezzi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMPUTAZIONE PREZZI"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18570
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImputazionePrezzi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   18570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   7440
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10935
      Left            =   0
      ScaleHeight     =   10905
      ScaleWidth      =   18465
      TabIndex        =   2
      Top             =   0
      Width           =   18495
      Begin VB.Frame Frame2 
         Caption         =   "IMPUTAZIONE PREZZI"
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
         Height          =   2655
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   18375
         Begin VB.CheckBox chkAggiornaIvaImballo 
            Caption         =   "Aggiorna codice I.V.A. dell'imballo"
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   2280
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaIvaMerce 
            Caption         =   "Aggiorna il codice I.V.A. della merce"
            Height          =   255
            Left            =   3720
            TabIndex        =   28
            Top             =   2040
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaDaListino 
            Caption         =   "Aggiorna l'importo unitario della merce come da listino selezionato"
            Height          =   255
            Left            =   3720
            TabIndex        =   27
            Top             =   1800
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaPrezziAZero 
            Caption         =   "Aggiorna Importo unitario della merce quando è zero"
            Height          =   255
            Left            =   3720
            TabIndex        =   9
            Top             =   1320
            Width           =   9015
         End
         Begin VB.CheckBox chkAggiornaPrezzoImballoDaListino 
            Caption         =   "Aggiorna l'importo unitario imballo da listino quando è zero"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1080
            Width           =   9015
         End
         Begin VB.CommandButton cmdAggiorna 
            Caption         =   "AGGIORNA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   15000
            TabIndex        =   7
            Top             =   1080
            Width           =   3255
         End
         Begin VB.CheckBox chkMerceInclusoImballo 
            Caption         =   "Imposta a tutte le righe l'indicazione di merce incluso imballo"
            Height          =   255
            Left            =   3720
            TabIndex        =   6
            Top             =   1560
            Width           =   9135
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
            Height          =   375
            Left            =   3720
            TabIndex        =   10
            Top             =   480
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecimalPlaces   =   5
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioImballo 
            Height          =   375
            Left            =   6240
            TabIndex        =   11
            Top             =   480
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecimalPlaces   =   5
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboListino 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTEDITNUMLib.dmtNumber txtSconto1 
            Height          =   375
            Left            =   8760
            TabIndex        =   13
            Top             =   480
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtSconto2 
            Height          =   375
            Left            =   10560
            TabIndex        =   14
            Top             =   480
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboListinoMerce 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTDataCmb.DMTCombo cboIvaMerce 
            Height          =   315
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTDataCmb.DMTCombo cboIvaImballo 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2280
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTEDITNUMLib.dmtNumber txtVariazioneManuale 
            Height          =   375
            Left            =   12480
            TabIndex        =   38
            Top             =   480
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecimalPlaces   =   5
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Variazione manuale"
            Height          =   255
            Index           =   4
            Left            =   12480
            TabIndex        =   39
            Top             =   240
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            X1              =   3720
            X2              =   18240
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            X1              =   3600
            X2              =   3600
            Y1              =   240
            Y2              =   2520
         End
         Begin VB.Label Label3 
            Caption         =   "I.V.A. imballo"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "I.V.A. merce"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Listino merce"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 2 %"
            Height          =   255
            Index           =   3
            Left            =   10560
            TabIndex        =   19
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 1 %"
            Height          =   255
            Index           =   2
            Left            =   8760
            TabIndex        =   18
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Listino imballi"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario imballo"
            Height          =   255
            Index           =   1
            Left            =   6240
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario articolo"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraRicerca 
         Caption         =   "PARAMETRI"
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
         Height          =   975
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   18375
         Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
            Height          =   615
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":4781A
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":47872
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":478D2
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
         Begin DmtCodDescCtl.DmtCodDesc CDImballo 
            Height          =   615
            Left            =   4920
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":4792C
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":47983
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":479E2
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
         Begin DmtCodDescCtl.DmtCodDesc CDSocio 
            Height          =   615
            Left            =   9240
            TabIndex        =   23
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1085
            PropCodice      =   $"frmImputazionePrezzi.frx":47A3C
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmImputazionePrezzi.frx":47A8A
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmImputazionePrezzi.frx":47AE4
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
         Begin DMTDataCmb.DMTCombo cboCategoriaMerceologica 
            Height          =   315
            Left            =   13320
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
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
         Begin DMTDataCmb.DMTCombo cboCategoriaFiscale 
            Height          =   315
            Left            =   15840
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
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
         Begin VB.Label Label3 
            Caption         =   "Categoria fiscale"
            Height          =   255
            Index           =   5
            Left            =   15840
            TabIndex        =   37
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Categoria merceologica"
            Height          =   255
            Index           =   4
            Left            =   13320
            TabIndex        =   35
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CONFERMA AGGIORNAMENTO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   15240
         TabIndex        =   3
         Top             =   10080
         Width           =   3135
      End
      Begin DmtGridCtl.DmtGrid GrigliaCorpo 
         Height          =   6375
         Left            =   0
         TabIndex        =   4
         Top             =   3600
         Width           =   18375
         _ExtentX        =   32411
         _ExtentY        =   11245
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
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   10080
         Width           =   12495
      End
   End
End
Attribute VB_Name = "frmImputazionePrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsImp As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim rsClone As ADODB.Recordset

Private Sub cboCategoriaFiscale_Click()
AGGIORNA_FILTRO
End Sub

Private Sub cboCategoriaMerceologica_Click()
    AGGIORNA_FILTRO
End Sub

Private Sub CDArticolo_ChangeElement()
    AGGIORNA_FILTRO
End Sub

Private Sub CDImballo_ChangeElement()
'Dim sSQL As String
'
'    sSQL = "RV_POTipoRiga=1 AND RV_PORigaCompleta=1"
'
'    If Me.CDArticolo.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND Link_art_articolo=" & Me.CDArticolo.KeyFieldID
'    End If
'    If Me.CDImballo.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND RV_POIDImballo=" & Me.CDImballo.KeyFieldID
'    End If
'    If Me.CDSocio.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND RV_POIDSocio=" & Me.CDSocio.KeyFieldID
'    End If
'
'    rsNew.Filter = sSQL
'
'
'    'rsNew.Requery
'
'    If Not (rsNew.EOF And rsNew.BOF) Then
'        Me.GrigliaCorpo.Requery
'    End If
'    Me.GrigliaCorpo.LoadUserSettings

AGGIORNA_FILTRO
End Sub

Private Sub CDSocio_ChangeElement()
'Dim sSQL As String
'
'    sSQL = "RV_POTipoRiga=1 AND RV_PORigaCompleta=1"
'
'    If Me.CDArticolo.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND Link_art_articolo=" & Me.CDArticolo.KeyFieldID
'    End If
'    If Me.CDImballo.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND RV_POIDImballo=" & Me.CDImballo.KeyFieldID
'    End If
'    If Me.CDSocio.KeyFieldID > 0 Then
'        sSQL = sSQL & " AND RV_POIDSocio=" & Me.CDSocio.KeyFieldID
'    End If
'
'    rsNew.Filter = sSQL
'
'
'
'
'    If Not (rsNew.EOF And rsNew.BOF) Then
'        Me.GrigliaCorpo.Requery
'    End If
'    Me.GrigliaCorpo.LoadUserSettings

AGGIORNA_FILTRO
End Sub

Private Sub cmdAggiorna_Click()
On Error GoTo ERR_cmdAggiorna_Click
Dim ImportoImballo As Double
Dim Testo As String

GrigliaCorpo.UpdatePosition = False

If Me.chkAggiornaDaListino.Value = vbChecked Then
    If Me.cboListinoMerce.CurrentID = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "È stato indicato di aggiornare i prezzi da un listino, ma il listino non è stato selezionato" & vbclrf
        Testo = Testo & "Impossibile continuare"
        
        Exit Sub
        
    End If
End If


    If Not (rsNew.BOF And rsNew.EOF) Then
        rsNew.MoveFirst
                
        While Not rsNew.EOF
                
            Screen.MousePointer = 11
            Me.lblInfo.Caption = "Aggiornamento in corso..."
            DoEvents
            
            If Me.chkAggiornaDaListino.Value = vbChecked Then
                rsNew("Art_prezzo_unitario_neutro").Value = GET_IMPORTO_DA_LISTINO(fnNotNullN(rsNew!Link_Art_articolo), cboListinoMerce.CurrentID)
            Else
                If Me.txtImportoUnitarioArticolo.Value > 0 Then
                    If Me.chkAggiornaPrezziAZero.Value = vbChecked Then
                        If fnNotNullN(rsNew("Art_prezzo_unitario_neutro").Value) = 0 Then
                            rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                        End If
                    Else
                        rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                    End If
                End If
            End If
            
            If Me.txtSconto1.Value > 0 Then
                rsNew("Art_sco_in_percentuale_1").Value = Me.txtSconto1.Value
            End If
            If Me.txtSconto2.Value > 0 Then
                rsNew("Art_sco_in_percentuale_2").Value = Me.txtSconto2.Value
            End If
            
            If Me.chkMerceInclusoImballo.Value = vbChecked Then
                rsNew("RV_POImportoImballoInArticolo").Value = Me.chkMerceInclusoImballo.Value
            End If
           
            If Me.chkAggiornaIvaMerce.Value = vbChecked Then
                If Me.cboIvaMerce.CurrentID > 0 Then
                    rsNew!Link_Art_IVA = Me.cboIvaMerce.CurrentID
                    rsNew!Art_aliquota_iva = GET_ALIQUOTA_IVA(Me.cboIvaMerce.CurrentID)
                End If
            End If
            
            
            rsNew!Art_pre_uni_net_sco_net_IVA = rsNew("Art_prezzo_unitario_neutro").Value - ((Value / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
            rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
            rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
            rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
            rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
            
            rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
            rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
            
            rsNew!Art_prezzo_unitario_netto_IVA = rsNew("Art_prezzo_unitario_neutro").Value
            rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
            rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
            rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
            rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
            rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA

            'ImportoImballo = fnNotNullN(rsNew!RV_POImportoUnitarioImballo)
            ImportoImballo = fnNotNullN(rsNew!RV_POImportoImballoSel)
            
            
            If Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked Then
                If rsNew.Fields("RV_POImportoImballoInArticolo").Value = False Then
                    rsNew.Fields("RV_POImportoUnitarioImballo").Value = GET_PREZZO_IMBALLO(ImportoImballo)
                Else
                    rsNew.Fields("RV_POImportoUnitarioImballo").Value = 0
                End If
            End If
            
            rsNew!RV_POImportoDaLiq = sbCalcolaImportoVariazioneLiquidazione(ImportoImballo)
            rsNew!RV_POVariazionePrezzoImballo = GET_VARIAZIONE_PREZZO_IMBALLO(ImportoImballo)
            rsNew!RV_POImportoMerceNetta = rsNew!Art_prezzo_unitario_netto_IVA - rsNew!RV_POVariazionePrezzoImballo
            rsNew!RV_POVariazionePrezzoManuale = Me.txtVariazioneManuale.Value
            
            rsNew.UpdateBatch
            Screen.MousePointer = 0
            DoEvents
        rsNew.MoveNext
        Wend

    End If
    
GrigliaCorpo.UpdatePosition = True

Screen.MousePointer = 0
Me.lblInfo.Caption = "Aggiornamento effettuato!"


MsgBox "Aggiornamento delle modifiche effettuato con successo!", vbInformation, "Aggiornamento prezzi documento"

Me.GrigliaCorpo.Refresh

Exit Sub
ERR_cmdAggiorna_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "cmdAggiorna_Click"
    
End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
Dim I As Integer

rsNew.UpdateBatch

rsNew.Filter = adFilterNone

Me.Command1.Enabled = False
Me.lblInfo.Caption = "RECUPERO DATI IN CORSO........................."
DoEvents

Me.GrigliaCorpo.UpdatePosition = False

If Not (rsNew.BOF And rsNew.EOF) Then
    rsNew.MoveFirst
    
    While Not rsNew.EOF
        rsClone.AddNew
            For I = 0 To rsNew.Fields.Count - 1
                rsClone.Fields(rsNew.Fields(I).Name).Value = rsNew.Fields(I).Value
            Next
        DoEvents
        rsClone.Update
        DoEvents
    rsNew.MoveNext
    Wend
End If

oDoc.Tables(sTabellaDettaglio).RemoveAllRetail

rsNew.MoveFirst
Me.lblInfo.Caption = "AGGIORNAMENTO DATI IN CORSO........................."
DoEvents

While Not rsNew.EOF
    
    Screen.MousePointer = 11
    DoEvents
    
    If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail 1
    Else
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
    End If
    
    For I = 0 To rsNew.Fields.Count - 1
        If (rsNew.Fields(I).Name <> "IDValoriOggettoDettaglio") And (rsNew.Fields(I).Name <> "IDOggetto") And (rsNew.Fields(I).Name <> "IDTipoOggetto") Then
            oDoc.Field rsNew.Fields(I).Name, rsNew.Fields(I).Value, sTabellaDettaglio
        End If
    Next
    
    If (rsNew!RV_PORigaCompleta = 1) And (rsNew!RV_POTipoRiga = 2) Then
        SCRIVI_RIGA_IMBALLO
    End If
    
    Screen.MousePointer = 0
    DoEvents

rsNew.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

DoEvents

Me.lblInfo.Caption = "REGISTRAZIONE DATI NEL DOCUMENTO........................."
DoEvents
oDoc.PerformTable sTabellaDettaglio, True

Me.lblInfo.Caption = ""

Unload Me

Exit Sub


ERR_Command1_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aggiornamento prezzi"
    Me.Command1.Enabled = True

End Sub
Private Sub SCRIVI_RIGA_IMBALLO()
Dim sSQL As String

rsClone.Filter = "RV_POTipoRiga=1 AND RV_POLinkRiga=" & rsNew!RV_POLinkRiga

If Not rsClone.EOF Then
    If Me.chkAggiornaIvaImballo.Value = vbChecked Then
        If Me.cboIvaImballo.CurrentID > 0 Then
            rsNew!Link_Art_IVA = Me.cboIvaImballo.CurrentID
            rsNew!Art_aliquota_iva = GET_ALIQUOTA_IVA(Me.cboIvaImballo.CurrentID)
            
            oDoc.Field "Link_art_IVA", rsNew!Link_Art_IVA, sTabellaDettaglio
            oDoc.Field "Art_aliquota_iva", rsNew!Art_aliquota_iva, sTabellaDettaglio
            
        End If
    End If
    
    oDoc.Field "Art_sco_in_percentuale_1", 0, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_2", 0, sTabellaDettaglio
    oDoc.Field "Art_importo_sconto_netto_IVA", 0, sTabellaDettaglio
    oDoc.Field "Art_importo_totale_lordo_IVA", (fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)) + (((fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)) / 100) * fnNotNullN(rsNew!Art_aliquota_iva)), sTabellaDettaglio
    oDoc.Field "Art_importo_totale_netto_IVA", (fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)), sTabellaDettaglio
                
    oDoc.Field "Art_prezzo_unitario_netto_IVA", fnNotNullN(rsClone!RV_POImportoUnitarioImballo), sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_lordo_IVA", fnNotNullN(rsClone!RV_POImportoUnitarioImballo) + ((fnNotNullN(rsClone!RV_POImportoUnitarioImballo) / 100) * fnNotNullN(rsNew!Art_aliquota_iva)), sTabellaDettaglio
    
    oDoc.Field "Art_pre_uni_net_sco_net_IVA", fnNotNullN(rsClone!RV_POImportoUnitarioImballo), sTabellaDettaglio
    oDoc.Field "Art_pre_uni_net_sco_lor_IVA", fnNotNullN(rsClone!RV_POImportoUnitarioImballo) + ((fnNotNullN(rsClone!RV_POImportoUnitarioImballo) / 100) * fnNotNullN(rsNew!Art_aliquota_iva)), sTabellaDettaglio
               
    oDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rsClone!RV_POImportoUnitarioImballo), sTabellaDettaglio
    oDoc.Field "Art_importo_totale_neutro", (fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)), sTabellaDettaglio
    
    oDoc.Field "Art_Importo_netto_IVA", (fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)), sTabellaDettaglio
                
    oDoc.Field "Art_importo_totale_netto_IVA", (fnNotNullN(rsClone!RV_POImportoUnitarioImballo) * fnNotNullN(rsClone!Art_numero_colli)), sTabellaDettaglio
    
    oDoc.Field "RV_POImportoImballoInArticolo", rsClone!RV_POImportoImballoInArticolo, sTabellaDettaglio
End If

End Sub

'Private Sub Command2_Click()
'    Me.GrigliaCorpo.SaveUserSettings
'    Me.GrigliaCorpo.Refresh
'End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With
    Me.lblInfo.Caption = ""
    INIT_CONTROLLI
    CREA_TABELLA_TEMPORANEA
    GET_GRIGLIA
    
    Me.cboListino.WriteOn GET_LISTINO_DEFAULT(frmMain.cdAnagrafica.KeyFieldID)
    
    If Me.cboListino.CurrentID > 0 Then
        Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked
    End If
    Me.chkAggiornaPrezziAZero.Value = vbChecked
    Me.chkMerceInclusoImballo.Value = GET_MERCE_INCLUSO_IMBALLO_CLIENTE(frmMain.cdAnagrafica.KeyFieldID)
    
    
    BLoandingPrezzatura = 1
    
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub CREA_TABELLA_TEMPORANEA()
Dim sSQL As String
Dim I As Long

sSQL = "SELECT * FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
'sSQL = sSQL & " AND RV_PORigaCompleta=1"

Set RsImp = New ADODB.Recordset
Set rsNew = New ADODB.Recordset
Set rsClone = New ADODB.Recordset

RsImp.Open sSQL, Cn.InternalConnection

rsNew.CursorLocation = adUseClient
rsClone.CursorLocation = adUseClient

''''CREA TABELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For I = 0 To RsImp.Fields.Count - 1
    Select Case RsImp.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsNew.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, RsImp.Fields(I).DefinedSize, RsImp.Fields(I).Attributes
            rsClone.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, RsImp.Fields(I).DefinedSize, RsImp.Fields(I).Attributes
            
        Case adInteger
            rsNew.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
            rsClone.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
        
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsNew.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
            rsClone.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
        
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsNew.Fields.Append RsImp.Fields(I).Name, adBoolean, , RsImp.Fields(I).Attributes
            rsClone.Fields.Append RsImp.Fields(I).Name, adBoolean, , RsImp.Fields(I).Attributes
        
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsNew.Fields.Append RsImp.Fields(I).Name, adDouble, , RsImp.Fields(I).Attributes
            rsClone.Fields.Append RsImp.Fields(I).Name, adDouble, , RsImp.Fields(I).Attributes
    End Select
Next

rsNew.Fields.Append "PesoNetto", adDouble, , adFldIsNullable
rsNew.Fields.Append "IDCategoriaMerceologica", adInteger, , adFldIsNullable
rsNew.Fields.Append "IDCategoriaFiscale", adInteger, , adFldIsNullable
rsNew.Fields.Append "DescrizioneCategoriaFiscale", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "DescrizioneCategoriaMerceologica", adVarChar, 250, adFldIsNullable

rsClone.Fields.Append "PesoNetto", adDouble, , adFldIsNullable
rsClone.Fields.Append "IDCategoriaMerceologica", adInteger, , adFldIsNullable
rsClone.Fields.Append "IDCategoriaFiscale", adInteger, , adFldIsNullable
rsClone.Fields.Append "DescrizioneCategoriaFiscale", adVarChar, 250, adFldIsNullable
rsClone.Fields.Append "DescrizioneCategoriaMerceologica", adVarChar, 250, adFldIsNullable

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rsNew.Open , , adOpenKeyset, adLockBatchOptimistic
rsClone.Open , , adOpenKeyset, adLockBatchOptimistic
While Not RsImp.EOF
    rsNew.AddNew
    rsClone.AddNew
        For I = 0 To RsImp.Fields.Count - 1
            rsNew.Fields(RsImp.Fields(I).Name).Value = RsImp.Fields(I).Value
        Next
        rsNew!PesoNetto = fnNotNullN(RsImp!Art_peso) - fnNotNullN(RsImp!Art_tara)
        GET_PROP_ARTICOLO fnNotNullN(RsImp!Link_Art_articolo), rsNew
    rsNew.Update
    rsClone.Update
RsImp.MoveNext
Wend

RsImp.Close
Set RsImp = Nothing
End Sub
Private Sub GET_GRIGLIA()
Dim OLD_Cursor As Long

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

rsNew.Filter = "RV_POTipoRiga=1 AND RV_PORigaCompleta=1"

    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
                .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
                .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Raggr. ord.", dgchar, True, 3000, dgAlignleft
                
                Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Imp. uni. merce", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."

                Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_1", "% Sc1", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_2", "% Sc2", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."


                Set cl = .ColumnsHeader.Add("Art_pre_uni_net_sco_net_IVA", "Importo netto IVA", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    cl.BackColor = vbRed
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    
                 Set cl = .ColumnsHeader.Add("RV_POVariazionePrezzoManuale", "Var. man.", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."

                Set cl = .ColumnsHeader.Add("RV_POImportoImballoInArticolo", "Incluso Imballo", dgBoolean, True, 1300, dgAligncenter)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                
                Set cl = .ColumnsHeader.Add("RV_POImportoDaLiq", "Variazione", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    cl.BackColor = vbRed
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 1100, dgAlignRight
                .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, True, 3000, dgAlignleft
                Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioImballo", "Imp. uni. imballo", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("Art_numero_colli", "Colli", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 1
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 1
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("Art_peso", "Peso lordo", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
                    'cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POIDSocio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POCodiceSocio", "codice socio/fornitore", dgchar, True, 1100, dgAlignleft
                .ColumnsHeader.Add "RV_POSocio", "Anagrafica socio/fornitore", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "RV_PONomeSocio", "Nome anagrafica socio/fornitore", dgchar, False, 3000, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaMerceologica", "IDCategoriaMerceologica", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "DescrizioneCategoriaMerceologica", "Categoria merceologica", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaFiscale", "IDCategoriaFiscale", dgInteger, True, 500, dgAlignRight
                .ColumnsHeader.Add "DescrizioneCategoriaFiscale", "Categoria fiscale", dgchar, False, 2000, dgAlignleft
                
        Set .Recordset = rsNew
        .LoadUserSettings
        .Refresh
    End With

Cn.CursorLocation = OLDCursor

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    

        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
     BLoandingPrezzatura = 0
End Sub

Private Sub GrigliaCorpo_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
    If rsNew!RV_POTipoRiga = 1 Then
        Select Case Column.FieldName
            Case "Art_prezzo_unitario_neutro"
                rsNew!Art_pre_uni_net_sco_net_IVA = Value - ((Value / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                
                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                
                rsNew!Art_prezzo_unitario_netto_IVA = Value
                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
                
            Case "Art_sco_in_percentuale_1"
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_prezzo_unitario_neutro - ((rsNew!Art_prezzo_unitario_neutro / 100) * Value)
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                
                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (Value + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                
                rsNew!Art_prezzo_unitario_netto_IVA = rsNew!Art_prezzo_unitario_neutro
                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
                
            Case "Art_sco_in_percentuale_2"
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_prezzo_unitario_neutro - ((rsNew!Art_prezzo_unitario_neutro / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * Value)
                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                
                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + Value))
                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                
                rsNew!Art_prezzo_unitario_netto_IVA = rsNew!Art_prezzo_unitario_neutro
                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
            Case "RV_POImportoImballoInArticolo"

        End Select
    
    End If


    If rsNew!RV_POTipoRiga = 2 Then
        Select Case Column.FieldName
            Case "Art_prezzo_unitario_neutro"
                rsNew!Art_pre_uni_net_sco_net_IVA = Value - ((Value / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                
                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                
                rsNew!Art_prezzo_unitario_netto_IVA = Value
                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
        End Select

    End If


    rsNew.UpdateBatch
    
    Me.GrigliaCorpo.Refresh

End Sub
Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsNew.Fields("RV_POImportoImballoInArticolo").Value))
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean)
Dim ImportoImballo As Double

    If Not rsNew.EOF And Not rsNew.BOF Then
                
        rsNew.Fields("RV_POImportoImballoInArticolo").Value = Abs(CLng(Selected))
    
        ImportoImballo = fnNotNullN(rsNew!RV_POImportoUnitarioImballo)
        
        If rsNew.Fields("RV_POImportoImballoInArticolo").Value = False Then
            rsNew.Fields("RV_POImportoUnitarioImballo").Value = GET_PREZZO_IMBALLO(ImportoImballo)
        Else
            rsNew.Fields("RV_POImportoUnitarioImballo").Value = 0
        End If
        
        rsNew!RV_POImportoDaLiq = sbCalcolaImportoVariazioneLiquidazione(ImportoImballo)
        
        rsNew.UpdateBatch
                
        Me.GrigliaCorpo.Refresh

    End If

End Sub
Private Function sbCalcolaImportoVariazioneLiquidazione(ImpImballo As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoImballo As Double
Dim ImportoImballoUnitario As Double
Dim LINK_UM_LIQUIDAZIONE As Long
Dim LINK_UM_COOP_ARTICOLO As Long
Dim MOLTIPLICATORE_ARTICOLO As Long


LINK_UM_COOP_ARTICOLO = fnGetUMCoop(rsNew!Link_Art_unita_di_misura)

sSQL = "SELECT RV_POMoltiplicatore, RV_POIDUnitaDiMisuraLiq "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & rsNew!Link_Art_articolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_UM_LIQUIDAZIONE = 0
    MOLTIPLICATORE_ARTICOLO = 1
Else
    LINK_UM_LIQUIDAZIONE = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
    MOLTIPLICATORE_ARTICOLO = fnNotNullN(rs!RV_POMoltiplicatore)
End If

rs.CloseResultset
Set rs = Nothing

ImportoImballo = ImpImballo

ImportoImballo = GET_PREZZO_IMBALLO(ImportoImballo)

If rsNew!Art_quantita_totale > 0 Then
    ImportoImballoUnitario = (ImportoImballo * rsNew!Art_numero_colli) / rsNew!Art_quantita_totale

    If rsNew!RV_POImportoImballoInArticolo = True Then
        'Me.txtImportoLiq.Value = Me.txtImponibileUnitario.Value - ImportoImballoUnitario
        sbCalcolaImportoVariazioneLiquidazione = -ImportoImballoUnitario
    Else
        'Me.txtImportoLiq.Value = Me.txtImponibileUnitario.Value
        sbCalcolaImportoVariazioneLiquidazione = 0
    End If
Else
       sbCalcolaImportoVariazioneLiquidazione = 0
End If

If LINK_UM_LIQUIDAZIONE = LINK_UM_COOP_ARTICOLO Then
    If Moltiplicatore > 0 Then
        sbCalcolaImportoVariazioneLiquidazione = sbCalcolaImportoVariazioneLiquidazione / Moltiplicatore
    Else
        sbCalcolaImportoVariazioneLiquidazione = sbCalcolaImportoVariazioneLiquidazione
    End If
Else
    If rsNew!RV_POQuantitaLiq > 0 Then
        sbCalcolaImportoVariazioneLiquidazione = (sbCalcolaImportoVariazioneLiquidazione * rsNew!Art_quantita_totale) / rsNew!RV_POQuantitaLiq
    Else
       sbCalcolaImportoVariazioneLiquidazione = sbCalcolaImportoVariazioneLiquidazione
    End If
End If

End Function
Private Function GET_PREZZO_IMBALLO(ImportoImballo As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If ImportoImballo = 0 Then

    If Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked Then
        sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
        sSQL = sSQL & "WHERE ("
        sSQL = sSQL & "(IDListino=" & Me.cboListino.CurrentID & ") "
        sSQL = sSQL & "AND (IDArticolo=" & rsNew!RV_POIDImballo & "))"
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF = False Then
            GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIva)
        Else
            GET_PREZZO_IMBALLO = 0
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        If GET_PREZZO_IMBALLO = 0 Then
            GET_PREZZO_IMBALLO = Me.txtImportoUnitarioImballo.Value
        End If
        
    Else
        GET_PREZZO_IMBALLO = Me.txtImportoUnitarioImballo.Value
    End If
Else
    GET_PREZZO_IMBALLO = ImportoImballo
End If
End Function
Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = rs!RV_POIDUnitaDiMisuraCoop
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub INIT_CONTROLLI()
     With Me.CDArticolo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


    'Imballo
    With Me.CDImballo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
     With Me.CDSocio
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Socio\Fornitore"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Socio\Fornitore"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Listino
    With Me.cboListino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT IDListino, Listino "
        .SQL = .SQL & "FROM Listino "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino=0"
        .Fill
    End With

    'Listino merce
    With Me.cboListinoMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT IDListino, Listino "
        .SQL = .SQL & "FROM Listino "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino=0"
        .Fill
    End With

    With Me.cboIvaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY Iva"
    End With
    
    With Me.cboIvaImballo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY Iva"
    End With

    With Me.cboCategoriaMerceologica
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaMerceologica"
        .DisplayField = "CategoriaMerceologica"
        .SQL = "SELECT IDCategoriaMerceologica, CategoriaMerceologica FROM CategoriaMerceologica"
        .SQL = .SQL & " ORDER BY CategoriaMerceologica"
    End With

    With Me.cboCategoriaFiscale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaFiscale"
        .DisplayField = "CategoriaFiscale"
        .SQL = "SELECT IDCategoriaFiscale, CategoriaFiscale FROM CategoriaFiscale"
        .SQL = .SQL & " ORDER BY CategoriaFiscale"
    End With
End Sub

Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Function GET_LISTINO_DEFAULT(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset

GET_LISTINO_DEFAULT = 0

sSQL = "SELECT IDListinoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDefault)
End If

rs.CloseResultset
Set rs = Nothing

If GET_LISTINO_DEFAULT > 0 Then Exit Function

sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
End If

rs.CloseResultset
Set rs = Nothing

If GET_LISTINO_DEFAULT > 0 Then Exit Function

GET_LISTINO_DEFAULT = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA

End Function
Private Function GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoImballiDefault "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = 0
Else
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = fnNotNullN(rs!IDListinoImballiDefault)
    
End If
rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_MERCE_INCLUSO_IMBALLO_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoInclusoImballo "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_MERCE_INCLUSO_IMBALLO_CLIENTE = 0
Else
    GET_MERCE_INCLUSO_IMBALLO_CLIENTE = Abs(fnNotNullN(rs!PrezzoInclusoImballo))
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_VARIAZIONE_PREZZO_IMBALLO(ImpImballo As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoImballo As Double
Dim ImportoImballoUnitario As Double
Dim LINK_UM_LIQUIDAZIONE As Long
Dim LINK_UM_COOP_ARTICOLO As Long
Dim MOLTIPLICATORE_ARTICOLO As Long




ImportoImballo = ImpImballo

ImportoImballo = GET_PREZZO_IMBALLO(ImportoImballo)

If rsNew!Art_quantita_totale > 0 Then
    ImportoImballoUnitario = (ImportoImballo * rsNew!Art_numero_colli) / rsNew!Art_quantita_totale

    If rsNew!RV_POImportoImballoInArticolo = True Then
        'Me.txtImportoLiq.Value = Me.txtImponibileUnitario.Value - ImportoImballoUnitario
        GET_VARIAZIONE_PREZZO_IMBALLO = ImportoImballoUnitario
    Else
        'Me.txtImportoLiq.Value = Me.txtImponibileUnitario.Value
        GET_VARIAZIONE_PREZZO_IMBALLO = 0
    End If
Else
       GET_VARIAZIONE_PREZZO_IMBALLO = 0
End If


End Function
Private Function GET_IMPORTO_DA_LISTINO(IDArticolo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDListino=" & IDListino

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_DA_LISTINO = 0
Else
    GET_IMPORTO_DA_LISTINO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Iva "
sSQL = sSQL & "WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ALIQUOTA_IVA = 0
Else
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_PROP_ARTICOLO(IDArticolo As Long, rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo.IDArticolo, Articolo.IDCategoriaMerceologica, CategoriaMerceologica.CategoriaMerceologica, Articolo.IDCategoriaFiscale, CategoriaFiscale.CategoriaFiscale "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "CategoriaMerceologica ON Articolo.IDCategoriaMerceologica = CategoriaMerceologica.IDCategoriaMerceologica LEFT OUTER JOIN "
sSQL = sSQL & "CategoriaFiscale ON Articolo.IDCategoriaFiscale = CategoriaFiscale.IDCategoriaFiscale "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    rstmp!IDCategoriaMerceologica = fnNotNullN(rs!IDCategoriaMerceologica)
    rstmp!DescrizioneCategoriaMerceologica = fnNotNull(rs!CategoriaMerceologica)
    rstmp!IDCategoriaFiscale = fnNotNullN(rs!IDCategoriaFiscale)
    rstmp!DescrizioneCategoriaFiscale = fnNotNull(rs!CategoriaFiscale)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub AGGIORNA_FILTRO()
Dim sSQL As String

    sSQL = "RV_POTipoRiga=1 AND RV_PORigaCompleta=1"

    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND Link_art_articolo=" & Me.CDArticolo.KeyFieldID
    End If
    If Me.CDImballo.KeyFieldID > 0 Then
        sSQL = sSQL & " AND RV_POIDImballo=" & Me.CDImballo.KeyFieldID
    End If
    If Me.CDSocio.KeyFieldID > 0 Then
        sSQL = sSQL & " AND RV_POIDSocio=" & Me.CDSocio.KeyFieldID
    End If
    If Me.cboCategoriaMerceologica.CurrentID > 0 Then
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & Me.cboCategoriaMerceologica.CurrentID
    End If
    If Me.cboCategoriaFiscale.CurrentID > 0 Then
        sSQL = sSQL & " AND IDCategoriaFiscale=" & Me.cboCategoriaFiscale.CurrentID
    End If
    
    
    rsNew.Filter = sSQL
    
    'rsNew.Requery
    
    If Not (rsNew.EOF And rsNew.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
Me.GrigliaCorpo.LoadUserSettings

End Sub
