VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaRigaOrd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LAVORAZIONE MERCE"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelezionaRigaOrd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   16635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraTotaliOrdine 
      Caption         =   "Totali merce lavorata"
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
      Height          =   4095
      Left            =   0
      TabIndex        =   16
      Top             =   4080
      Width           =   3015
      Begin VB.TextBox txtTotalePesoNettoOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePedaneOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleColliOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePesoLordoOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleTaraOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePezziOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleTipoPedaneOrdPrep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Peso Netto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   2880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label8 
         Caption         =   "N° pedane"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Colli"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Peso lordo"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Tara"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   2880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   120
         X2              =   2880
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   120
         X2              =   2880
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label8 
         Caption         =   "Pezzi"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   2880
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label8 
         Caption         =   "N° Tipo ped."
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   120
         X2              =   2880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   120
         X2              =   2880
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.Frame FraTotaliOrdineSmist 
      Caption         =   "Totale merce ordinata"
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
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtTotalePesoNettoOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePedaneOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleColliOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePesoLordoOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleTaraOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtTotalePezziOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtTotaleTipoPedaneOrdSmist 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Peso Netto"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   120
         X2              =   2880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label8 
         Caption         =   "N° pedane"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Colli"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Peso lordo"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Tara"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   120
         X2              =   2880
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   120
         X2              =   2880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   120
         X2              =   2880
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label8 
         Caption         =   "Pezzi"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   975
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   2880
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label8 
         Caption         =   "N° Tipo ped."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   120
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.CheckBox chkRicerca 
      BackColor       =   &H00FF0000&
      Caption         =   "Solo articolo selezionato"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3040
      TabIndex        =   0
      Top             =   30
      Width           =   3495
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3855
      Left            =   3000
      TabIndex        =   31
      Top             =   4320
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   6800
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
   Begin DmtGridCtl.DmtGrid GrigliaCorpoOrdine 
      Height          =   3735
      Left            =   3000
      TabIndex        =   32
      Top             =   240
      Width           =   13575
      _ExtentX        =   23945
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "MERCE ORDINATA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   34
      Top             =   0
      Width           =   13575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "LAVORAZIONE MERCE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   33
      Top             =   4080
      Width           =   13575
   End
End
Attribute VB_Name = "frmSelezionaRigaOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsGriglia As DmtOleDbLib.adoResultset
Public rsGrigliaOrd As DmtOleDbLib.adoResultset

'L'oggetto PaintNotify usato per la gestione dei campi calcolati
Public gPaintNotify As PaintNotify

Private Sub fncGriglia()
On Error GoTo ERR_fncGriglia
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
sSQL = "SELECT * FROM RV_POIEControlloAssOrdini "
sSQL = sSQL & " WHERE RV_POIEControlloAssOrdini.IDOggettoOrdinePadre=" & FrmMain.txtIDOrdinePadre.Value
If chkRicerca.Value = vbChecked Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
End If


    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    Set rsGriglia = CnDMT.OpenResultset(sSQL)
        Set rsEvent = rsGriglia.Data

    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POPedana", "IDRV_POPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodicePedana", "Codice pedana", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "IDRV_POTipoPedana", "IDRV_POPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "IDUnitaDiMisura", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisura", "U.M.", dgchar, True, 1000, dgAlignleft
            Set cl = .ColumnsHeader.Add("Colli", "Colli", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "IDUnitaDiMisuraCoop", "IDUnitaDiMisuraCoop", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisuraCoop", "U.M. coop.", dgchar, True, 1000, dgAlignleft
            Set cl = .ColumnsHeader.Add("Qta_UM", "Q.tà Mov.", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "IDImballoVendita", "IDImballoVendita", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imballo", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "IDAnagraficaSocio", "IDAnagraficaSocio", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceSocio", "Codice Socio", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "NumeroConferimento", "N° conf.", dgNumeric, True, 1800, dgAlignRight
            .ColumnsHeader.Add "DataDocumento", "Data Lav.", dgDate, True, 1800, dgAlignleft
            .ColumnsHeader.Add "OraLavorazione", "Ora Lav.", dgchar, False, 2000, dgAlignleft

            

                        
            Set .Recordset = rsGriglia.Data
            .LoadUserSettings
            .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor

Exit Sub
ERR_fncGriglia:
    MsgBox Err.Description, vbCritical, "fncGriglia"
End Sub

Private Sub chkRicerca_Click()

    fncGriglia
    GET_TOTALI_ORDINE_DA_PREPARARE FrmMain.txtIDOrdinePadre.Value
    GET_TOTALI_ORDINE FrmMain.txtIDOrdinePadre.Value
    
End Sub

Private Sub Form_Activate()
    INIT_CONTROLLI
    fncGriglia
    fncGrigliaOrdine
    GET_TOTALI_ORDINE_DA_PREPARARE FrmMain.txtIDOrdinePadre.Value
    GET_TOTALI_ORDINE FrmMain.txtIDOrdinePadre.Value
    RIGA_ORDINE_SELEZIONATA = 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    'Me.Top = frmMain.cmdLavorazioni.Top
    'Me.Left = frmMain.cmdLavorazioni.Left + frmMain.cmdLavorazioni.Width + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsGriglia.CloseResultset
    Set rsGriglia = Nothing
    rsGrigliaOrd.CloseResultset
    Set rsGrigliaOrd = Nothing

End Sub
Private Sub INIT_CONTROLLI()
    Set gPaintNotify = New PaintNotify
    
'    With Me.cboTipoRaggr
'        Set .Database = TheApp.Database.Connection
'        .AddFieldKey "IDRV_POTipoRaggrOrdLav"
'        .DisplayField = "TipoRaggrOrdLav"
'        .Sql = "SELECT * FROM RV_POTipoRaggrOrdLav ORDER BY IDRV_POTipoRaggrOrdLav"
'    End With

End Sub
Private Function GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_PEDANE_LAVORATE = 0

sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
If chkRicerca.Value = vbChecked Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
End If
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_PEDANE_LAVORATE = GET_NUMERO_PEDANE_LAVORATE + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_TIPO_PEDANE_LAVORATE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_TIPO_PEDANE_LAVORATE = 0

sSQL = "SELECT RV_POTipoPedana.CodiceTipoPedana, RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POPedana ON RV_POTipoPedana.IDRV_POTipoPedana = RV_POPedana.IDRV_POTipoPedana RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POAssegnazioneMerce ON RV_POPedana.IDRV_POPedana = RV_POAssegnazioneMerce.IDRV_POPedana "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
If chkRicerca.Value = vbChecked Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
End If
sSQL = sSQL & "GROUP BY RV_POTipoPedana.CodiceTipoPedana, RV_POTipoPedana.IDRV_POTipoPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_TIPO_PEDANE_LAVORATE = GET_NUMERO_TIPO_PEDANE_LAVORATE + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TOTALI_ORDINE_DA_PREPARARE(IDOggettoOrdine As Long) As Double
On Error GoTo ERR_GET_TOTALI_ORDINE_DA_PREPARARE
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Colli) as NumeroColli,"
sSQL = sSQL & "SUM(PesoLordo) as PesoLordo, "
sSQL = sSQL & "SUM(PesoNetto) as PesoNetto, "
sSQL = sSQL & "SUM(Tara) as Tara, "
sSQL = sSQL & "SUM(Pezzi) as Pezzi "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
If chkRicerca.Value = vbChecked Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglioRigaOrd=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
End If
Set rs = CnDMT.OpenResultset(sSQL)
    
If Not rs.EOF Then
    Me.txtTotaleColliOrdPrep.Text = FormatNumber(fnNotNullN(rs!NumeroColli), 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdPrep.Text = FormatNumber(fnNotNullN(rs!PesoLordo), 2, , , vbTrue)
    Me.txtTotaleTaraOrdPrep.Text = FormatNumber(fnNotNullN(rs!Tara), 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdPrep.Text = FormatNumber(fnNotNullN(rs!PesoNetto), 2, , , vbTrue)
    Me.txtTotalePezziOrdPrep.Text = FormatNumber(fnNotNullN(rs!Pezzi), 2, , , vbTrue)
Else
    Me.txtTotaleColliOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotaleTaraOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePezziOrdPrep.Text = FormatNumber(0, 2, , , vbTrue)
End If

rs.CloseResultset
Set rs = Nothing
    
Me.txtTotalePedaneOrdPrep.Text = FormatNumber(GET_NUMERO_PEDANE_LAVORATE(IDOggettoOrdine), 2, , , vbTrue)
Me.txtTotaleTipoPedaneOrdPrep.Text = FormatNumber(GET_NUMERO_TIPO_PEDANE_LAVORATE(IDOggettoOrdine), 2, , , vbTrue)

Exit Function
ERR_GET_TOTALI_ORDINE_DA_PREPARARE:
    MsgBox Err.Description, vbCritical, "GET_TOTALI_ORDINE_DA_PREPARARE"
End Function
Private Sub fncGrigliaOrdine()
On Error GoTo ERR_fncGrigliaOrdine
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    

sSQL = "SELECT * FROM  ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & FrmMain.txtIDOrdinePadre.Value
sSQL = sSQL & " AND RV_POTipoRiga=1"

    
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    Set rsGrigliaOrd = CnDMT.OpenResultset(sSQL)
        Set rsEvent = rsGrigliaOrd.Data

    With Me.GrigliaCorpoOrdine
        .ColumnsHeader.Clear
        Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Link_art_articolo", "Link_art_articolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Art_descrizione", "Art_descrizione", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Tipo pedana", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione pedana", dgchar, False, 2000, dgAlignleft
            Set cl = .ColumnsHeader.Add("RV_POQuantitaPedana", "Q.ta pedane", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "RV_POIDArticoloPedana", "RV_POIDArticoloPedana", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POCodiceArticoloPedana", "Codice articolo pedana", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_PODescrizioneArticoloPedana", "Descrizione articolo pedana", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_POIDTipoUMOrdine", "RV_POIDTipoUMOrdine", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POTipoUMOrdine", "U.M. Riga Ordine", dgchar, True, 2000, dgAlignleft
            
            
            Set cl = .ColumnsHeader.Add("art_numero_colli", "Colli", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Art_Peso", "Peso lordo", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Art_tara", "Tara", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, True, 900, dgAlignRight, True, True, True)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "Art_sigla_unita_di_misura", "U.M. Mov.", dgchar, False, 1000, dgAlignleft
            Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà Mov.", dgDouble, True, 900, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "RV_POCodiceImballo", "Codice articolo Imballo", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione articolo imballo", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Raggrupp. ord.", dgchar, True, 2500, dgAlignleft
            
            Set .Recordset = rsGrigliaOrd.Data
            .LoadUserSettings
            .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor

Exit Sub
ERR_fncGrigliaOrdine:
    MsgBox Err.Description, vbCritical, "fncGrigliaOrdine"
End Sub
Private Function GET_TOTALI_ORDINE(IDOggettoOrdine As Long) As Double
On Error GoTo ERR_GET_TOTALI_ORDINE
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Art_numero_Colli) as NumeroColli,"
sSQL = sSQL & "SUM(Art_Peso) as PesoLordo, "
'sSQL = sSQL & "SUM(PesoNetto) as PesoNetto, "
sSQL = sSQL & "SUM(Art_Tara) as Tara, "
sSQL = sSQL & "SUM(Art_quantita_Pezzi) as Pezzi, "
sSQL = sSQL & "SUM(RV_POQuantitaPedana) as NumeroPedane "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1"
If chkRicerca.Value = vbChecked Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & fnNotNullN(GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
End If
Set rs = CnDMT.OpenResultset(sSQL)
    
If Not rs.EOF Then
    Me.txtTotaleColliOrdSmist.Text = FormatNumber(fnNotNullN(rs!NumeroColli), 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdSmist.Text = FormatNumber(fnNotNullN(rs!PesoLordo), 2, , , vbTrue)
    Me.txtTotaleTaraOrdSmist.Text = FormatNumber(fnNotNullN(rs!Tara), 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdSmist.Text = FormatNumber((fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara)), 2, , , vbTrue)
    Me.txtTotalePezziOrdSmist.Text = FormatNumber(fnNotNullN(rs!Pezzi), 2, , , vbTrue)
    Me.txtTotalePedaneOrdSmist.Text = FormatNumber(fnNotNullN(rs!NumeroPedane), 2, , , vbTrue)
Else
    Me.txtTotaleColliOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoLordoOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotaleTaraOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePesoNettoOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePezziOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
    Me.txtTotalePedaneOrdSmist.Text = FormatNumber(0, 2, , , vbTrue)
End If

rs.CloseResultset
Set rs = Nothing
    
Me.txtTotaleTipoPedaneOrdSmist.Text = FormatNumber(GET_NUMERO_TIPO_PEDANE_ORDINATE(IDOggettoOrdine), 2, , , vbTrue)
Exit Function
ERR_GET_TOTALI_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_TOTALI_ORDINE"
End Function
Private Function GET_NUMERO_TIPO_PEDANE_ORDINATE(IDOggettoOrdine As Long) As Double
On Error GoTo ERR_GET_NUMERO_TIPO_PEDANE_ORDINATE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_TIPO_PEDANE_ORDINATE = 0

sSQL = "SELECT RV_POIDTipoPedana "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & "GROUP BY RV_POIDTipoPedana"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_TIPO_PEDANE_ORDINATE = GET_NUMERO_TIPO_PEDANE_ORDINATE + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_NUMERO_TIPO_PEDANE_ORDINATE:
    MsgBox Err.Description, vbCritical, "GET_NUMERO_TIPO_PEDANE_ORDINATE"
    
End Function
Private Sub GrigliaCorpoOrdine_DblClick()
    frmAssegnazioneMerce.txtIDRigaOrdine.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
    
    RIGA_ORDINE_SELEZIONATA = 1
    Unload Me
End Sub

Private Sub GrigliaCorpoOrdine_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    fncGriglia
    GET_TOTALI_ORDINE_DA_PREPARARE FrmMain.txtIDOrdinePadre.Value
    GET_TOTALI_ORDINE FrmMain.txtIDOrdinePadre.Value
End Sub

