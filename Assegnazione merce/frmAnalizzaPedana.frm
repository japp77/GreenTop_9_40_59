VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmAnalizzaPedana 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalizzaPedana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraRiepilogo 
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
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   12255
      Begin DMTEDITNUMLib.dmtNumber txtQta_UM 
         Height          =   285
         Left            =   9840
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin DMTEDITNUMLib.dmtNumber txtPezzi 
         Height          =   285
         Left            =   8040
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin DMTEDITNUMLib.dmtNumber txtTara 
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin DMTEDITNUMLib.dmtNumber txtColli 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin VB.Label Label2 
         Caption         =   "Q.tà Mov."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   9840
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Peso netto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tara "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Pezzi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8040
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Peso lordo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Colli"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8281
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
Attribute VB_Name = "frmAnalizzaPedana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    
    Me.Top = frmMain.Frame1.Height + 1300
    Me.Left = frmMain.fraRighe.Left
    Me.Width = frmMain.fraRighe.Width
    
    Me.Caption = "Pedana: " & frmMain.txtCodicePedana.Text & " (" & frmMain.CDTipoPedana.Description & ")"
    
    SettaggioGriglia
    SettaggioRiepilogo

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
    
End Sub
Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT RV_POAssegnazioneMerce.*, RV_POCaricoMerceRighe.IDArticolo AS IDArticoloConferito,"
    sSQL = sSQL & "RV_POCaricoMerceRighe.CodiceArticolo AS CodiceArticolo_Conferito, RV_POCaricoMerceRighe.Articolo AS Articolo_Conferito,"
    sSQL = sSQL & "ValoriOggettoPerTipo000F.Doc_ordine_chiuso, ValoriOggettoPerTipo000F.Doc_data, ValoriOggettoPerTipo000F.Doc_numero,"
    sSQL = sSQL & "ValoriOggettoPerTipo000F.Nom_codice, ValoriOggettoPerTipo000F.Nom_ragione_sociale_o_cognome, ValoriOggettoPerTipo000F.Nom_nome,"
    sSQL = sSQL & "Oggetto.IDAzienda, ValoriOggettoPerTipo000F.RV_PONumeroOrdinePadre, ValoriOggettoPerTipo000F.RV_PODataOrdinePadre,  "
    sSQL = sSQL & "ValoriOggettoPerTipo000F.RV_PONumeroListaPrelievo, ValoriOggettoPerTipo000F.RV_POIDOrdinePadre "
    sSQL = sSQL & "FROM Oggetto INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo000F ON Oggetto.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto AND "
    sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo000F.IDTipoOggetto RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_POAssegnazioneMerce ON Oggetto.IDOggetto = RV_POAssegnazioneMerce.IDOggettoOrdine LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceTesta RIGHT OUTER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta ON "
    sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDRV_POPedana=" & frmMain.txtIDPedana.Value

    
        Set rsArt = Cn.OpenResultset(sSQL)
            Set rsEvent = rsArt.data
        
        With Me.GrigliaCorpo
            .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POAssegnazioneMerce", "ID", dgInteger, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceArticolo", "Codice Art.", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Qta_UM", "Quantità", dgDouble, True, 1100, dgAlignRight, True, True, False
            .ColumnsHeader.Add "Colli", "Colli", dgDouble, True, 1100, dgAlignRight, True, True, False
            Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Tara", "Tara", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Pezzi", "Pezzi", dgDouble, False, 1100, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "IDImballoVendita", "IDImballo", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceImballoVendita", "Codice Imb.", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodicePedana", "CodicePedana", dgchar, True, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDAnagraficaSocio", "IDSocio", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceSocio", "Codice socio", dgchar, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NumeroConferimento", "N° Conf.", dgInteger, True, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceLottoVendita", "Lotto di vendita", dgchar, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDCliente", "IDCliente", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Cliente", dgchar, True, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "RV_PODataOrdinePadre", "Data ord.", dgDate, True, 1500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "RV_PONumeroOrdinePadre", "N° ord.", dgNumeric, True, 1500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "RV_PONumeroListaPrelievo", "N° lista", dgNumeric, True, 1500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDArticolo_Conferito", "IDArticolo_Conferito", dgInteger, False, 2000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceArticolo_conferito", "Cod. Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Articolo_conferito", "Art. Conf.", dgchar, False, 1700, dgAlignleft, True, True, False
            Set .Recordset = rsArt.data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio analisi pedana"
    
End Sub

Private Sub Form_Resize()
On Error GoTo ERR_Form_Resize
If Me.WindowState = 1 Then Exit Sub
If Me.ScaleHeight < 1000 Then Exit Sub
    Me.fraRiepilogo.Top = Me.ScaleHeight - Me.fraRiepilogo.Height - 100
    Me.fraRiepilogo.Left = (Me.ScaleWidth / 2) - (Me.fraRiepilogo.Width / 2)
    
    
    Me.GrigliaCorpo.Width = Me.Width - 200
    Me.GrigliaCorpo.Height = Me.ScaleHeight - Me.fraRiepilogo.Height - 200
Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsArt Is Nothing) Then
        rsArt.CloseResultset
        Set rsArt = Nothing
    End If
End Sub
Private Sub SettaggioRiepilogo()
On Error GoTo ERR_SettaggioRiepilogo
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Colli) AS TotaleColli, "
sSQL = sSQL & "SUM(PesoLordo) AS TotalePesoLordo, "
sSQL = sSQL & "SUM(Tara) AS TotaleTara, "
sSQL = sSQL & "SUM(PesoNetto) AS TotalePesoNetto, "
sSQL = sSQL & "SUM(Pezzi) AS TotalePezzi, "
sSQL = sSQL & "SUM(Qta_UM) AS TotaleQuantitaMovimentata "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POPedana=" & frmMain.txtIDPedana

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtColli.Value = fnNotNullN(rs!TotaleColli)
    Me.txtPesoLordo.Value = fnNotNullN(rs!TotalePesoLordo)
    Me.txtTara.Value = fnNotNullN(rs!TotaleTara)
    Me.txtPesoNetto.Value = fnNotNullN(rs!TotalePesoNetto)
    Me.txtPezzi.Value = fnNotNullN(rs!TotalePezzi)
    Me.txtQta_UM.Value = fnNotNullN(rs!TotaleQuantitaMovimentata)
End If



rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_SettaggioRiepilogo:
    MsgBox Err.Description, vbCritical, "Settaggio riepilogo pedana"
End Sub

