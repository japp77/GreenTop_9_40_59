VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#10.15#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRicerca 
   Caption         =   "Ricerca articoli"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin DmtGridCtl.DmtGrid GridRicerca 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12726
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      UpdatePosition  =   0   'False
      UseUserSettings =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmRicerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As DmtADOLib.adoResultset
Public gPaintNotify As PaintNotify

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Click_Ricerca
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    Me.Width = frmMain.Width - 100
    Me.Top = frmMain.Height / 2
    Me.Height = frmMain.Height / 2 - 500
    Me.Left = 0
    
    Set gPaintNotify = New PaintNotify
    GetGrigliaRicerca
    Me.GridRicerca.Recordset.Requery
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub GetGrigliaRicerca()
On Error GoTo ERR_GetGrigliaLavorazione
Dim sSQL As String
Dim cl As dgColumnHeader




    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        
    Set rsGriglia = Cn.OpenResultset(sSQL_Ricerca)
            Set rsEvent = rsGriglia.Data
        
    
        With Me.GridRicerca
            Set .PaintNotifyObj = gPaintNotify
            .ColumnsHeader.Clear
                    .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignlef
                    .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2000, dgAlignleft
                    .ColumnsHeader.Add "IDCodiceLotto_Vendita", "IDLotto", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceLottoVendita", "Codice lotto", dgchar, True, 2000, dgAlignleft
                    .ColumnsHeader.Add "DescrizioneLottoVendita", "Descrizione lotto", dgchar, False, 2000, dgAlignleft
                    Set cl = .ColumnsHeader.Add("ColliDisponibili", "Colli disponibili", dgDouble, True, 1000, dgAlignRight, , , True)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("ColliImpeganti", "Colli impegnati", dgDouble, True, 1000, dgAlignRight, , , True)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Colli", "Colli caricati", dgCurrency, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("QtaColliVenduti", "Colli venduti", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("DisponibilitaLotto", "Disp. lotto", dgDouble, True, 1000, dgAlignRight, , , True)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("PesoLordo", "Peso lordo lav.", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto lav.", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Tara", "Tara lav.", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    Set cl = .ColumnsHeader.Add("Qta_UM", "Q.tà mov.", dgDouble, True, 1000, dgAlignRight)
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    .ColumnsHeader.Add "CodiceImballoVendita", "Codice imballo", dgchar, False, 2000, dgAlignleft
                    .ColumnsHeader.Add "ImballoVendita", "Imballo", dgchar, False, 2000, dgAlignleft
                    .ColumnsHeader.Add "Conf_Anagrafica", "Socio", dgchar, False, 2000, dgAlignleft
                    .ColumnsHeader.Add "Conf_Nome", "Nome socio", dgchar, False, 1000, dgAlignleft
                    .ColumnsHeader.Add "Conf_DataDocumento", "Data conf", dgDate, False, 1500, dgAlignleft
                    .ColumnsHeader.Add "Conf_NumeroDocumento", "Numero Doc.", dgchar, False, 1000, dgAlignleft
                    .ColumnsHeader.Add "Conf_CodiceLotto", "Cod. lotto Conf.", dgchar, False, 1000, dgAlignleft
                    .ColumnsHeader.Add "Conf_DescrizioneLotto", "Lotto Conf.", dgchar, False, 1000, dgAlignleft
                    .ColumnsHeader.Add "Conf_CodiceImballo", "Cod. imballo", dgchar, False, 1000, dgAlignleft
                    .ColumnsHeader.Add "Conf_DescrizioneImballo", "Imballo", dgchar, False, 1000, dgAlignleft
                    
                    
            Set .Recordset = rsGriglia.Data
            .Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GetGrigliaLavorazione:
    MsgBox Err.Description, vbCritical, "Griglia ricerca"
End Sub



Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Me.GridRicerca.Width = Me.Width - 150
        Me.GridRicerca.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsGriglia.CloseResultset
    Set rsGriglia = Nothing
End Sub


Private Sub Click_Ricerca()
On Error GoTo ERR_GridRicerca_DblClick

Exit Sub
ERR_GridRicerca_DblClick:
    MsgBox Err.Description, vbCritical, "GridRicerca_DblClick"
End Sub

