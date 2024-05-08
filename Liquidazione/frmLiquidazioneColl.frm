VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmLiquidazioneColl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidazione collegata"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLiquidazioneColl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleLordoDocumento 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleIvaDocumento 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleDocumento 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenute 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtNettoLiquidazione 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenutePerLavorazioni 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteGenerali 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteAgg 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4680
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattAggRiep 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   5280
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteConferimento 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   " 0"
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
         BorderStyle     =   1
         Enabled         =   0   'False
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Totale tratt. agg. riepilogo"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Totale tratt. aggiuntive"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Totale netto documento"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Totale trattenute"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Importo di liquidazione"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Totale I.V.A. documento"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Totale lordo documento"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Trattenute per lavorazioni"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Trattenute generali"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Trattenute conferimento"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLiquidazioneColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    GET_TOTALI_LIQ
End Sub
Private Sub GET_TOTALI_LIQ()
On Error GoTo ERR_GET_TOTALI_LIQ
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLiquidazione "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & LINK_LIQUIDAZIONE_COLL

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtTotaleDocumento.Value = fnNotNullN(rs!TotaleDocumento)
    txtTotaleIvaDocumento.Value = fnNotNullN(rs!TotaleIva)
    txtTotaleLordoDocumento.Value = fnNotNullN(rs!TotaleDocumentoLordoIva)
    txtTotaleTrattenutePerLavorazioni.Value = fnNotNullN(rs!TotaleTrattenutaPerLavorazione)
    txtTotaleTrattenuteGenerali.Value = fnNotNullN(rs!TotaleTrattenutaGenerale)
    txtTotaleTrattenuteConferimento.Value = fnNotNullN(rs!TotaleTrattenuteConferimento)
    txtTotaleTrattenute.Value = fnNotNullN(rs!TotaleTrattenuta)
    txtTotaleTrattenuteAgg.Value = fnNotNullN(rs!TotaleTrattenuteAggiuntive)
    txtTotaleTrattAggRiep.Value = fnNotNullN(rs!TotaleTrattenuteRiepilogo)
    txtNettoLiquidazione.Value = fnNotNullN(rs!NettoLiquidazione)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_TOTALI_LIQ:
    MsgBox Err.Description, vbCritical, "GET_TOTALI_LIQ"
End Sub
