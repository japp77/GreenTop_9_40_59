VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPrezziVeloci 
   Caption         =   "OPERAZIONE MASSIVA DEGLI IMPORTI UNITARI DELLA LIQUIDAZIONE"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrezziVeloci.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   18675
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   840
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10215
      Left            =   0
      ScaleHeight     =   10185
      ScaleWidth      =   18585
      TabIndex        =   0
      Top             =   0
      Width           =   18615
      Begin VB.CommandButton Command1 
         Caption         =   "CONFERMA AGGIORNAMENTO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   9360
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Valori per le operazioni massive"
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
         Height          =   1215
         Left            =   4080
         TabIndex        =   12
         Top             =   0
         Width           =   14415
         Begin VB.CommandButton Command5 
            Caption         =   "DESEL. TUTTI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   2295
         End
         Begin VB.CommandButton Command4 
            Caption         =   "SEL. TUTTI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   0
            Width           =   2295
         End
         Begin VB.CommandButton cmdAggiorna 
            Caption         =   "AGGIORNA DATI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   12000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   2295
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   600
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
            Caption         =   "Importo unitario articolo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame FraRicerca 
         Caption         =   "Parametri di ricerca"
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
         Height          =   9255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         Begin VB.CommandButton Command2 
            Caption         =   "RICERCA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            Caption         =   "PULISCI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtCodiceArticoloRic 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   1560
            Width           =   3615
         End
         Begin VB.TextBox txtDescriArticoloRic 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   2160
            Width           =   3615
         End
         Begin VB.TextBox txtCatMercRic 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   2760
            Width           =   3615
         End
         Begin VB.CheckBox chkNonScarti 
            Caption         =   "Non visualizzare gli scarti"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   3240
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            X1              =   120
            X2              =   3720
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label2 
            Caption         =   "Codice articolo"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "Descrizione articolo"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "Categoria merceologica"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   2520
            Width           =   3615
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   9840
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin DmtGridCtl.DmtGrid GrigliaCorpo 
         Height          =   7935
         Left            =   4080
         TabIndex        =   17
         Top             =   1320
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13996
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
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   18
         Top             =   9480
         Width           =   14415
      End
   End
End
Attribute VB_Name = "frmPrezziVeloci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsRigheElaTMP As ADODB.Recordset

Private Sub cmdAggiorna_Click()
On Error GoTo ERR_cmdAggiorna_Click
Dim FiltroPrec As String

Me.GrigliaCorpo.UpdatePosition = False
FiltroPrec = rsRigheElaTMP.Filter
rsRigheElaTMP.Filter = "Selezionato=1"

While Not rsRigheElaTMP.EOF
    rsRigheElaTMP!ImportoUnitarioDaReg = Me.txtImportoUnitarioArticolo.Value
    rsRigheElaTMP!ImponibileDaReg = rsRigheElaTMP!ImportoUnitarioDaReg * fnNotNullN(rsRigheElaTMP!QuantitaVenduta)
    rsRigheElaTMP!Selezionato = 0
    rsRigheElaTMP.UpdateBatch

rsRigheElaTMP.MoveNext
Wend

If FiltroPrec <> "0" Then
    rsRigheElaTMP.Filter = FiltroPrec
Else
    rsRigheElaTMP.Filter = vbNullString
End If

Me.GrigliaCorpo.Refresh
Me.GrigliaCorpo.UpdatePosition = True
Exit Sub
ERR_cmdAggiorna_Click:
    MsgBox Err.Description, vbCritical, "cmdAggiorna_Click"
End Sub

Private Sub Command1_Click()
    
    ELABORAZIONE_LIQUIDAZIONE LINK_LIQUIDAZIONE
    
End Sub

Private Sub Command2_Click()
    GET_GRIGLIA
End Sub

Private Sub Command3_Click()
    Me.txtCatMercRic.Text = ""
    Me.txtCodiceArticoloRic.Text = ""
    Me.txtDescriArticoloRic.Text = ""
    Me.chkNonScarti.Value = vbUnchecked
End Sub

Private Sub Command4_Click()
Me.GrigliaCorpo.UpdatePosition = False

rsRigheElaTMP.MoveFirst

While Not rsRigheElaTMP.EOF
    rsRigheElaTMP!Selezionato = 1
rsRigheElaTMP.MoveNext
Wend
rsRigheElaTMP.MoveFirst
Me.GrigliaCorpo.UpdatePosition = True

Me.GrigliaCorpo.Refresh

End Sub

Private Sub Command5_Click()
    Me.GrigliaCorpo.UpdatePosition = False
    
    rsRigheElaTMP.MoveFirst
    
    While Not rsRigheElaTMP.EOF
        rsRigheElaTMP!Selezionato = 0
    rsRigheElaTMP.MoveNext
    Wend
    rsRigheElaTMP.MoveFirst
    Me.GrigliaCorpo.UpdatePosition = True
    
    Me.GrigliaCorpo.Refresh
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    
        If Me.Width > 18615 Then
            Me.Pic1.Width = Me.Width - 250
            Me.GrigliaCorpo.Width = Me.Pic1.Width - 100 - Me.GrigliaCorpo.Left
        End If
        
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
Private Sub Form_Load()
    CONFERMA_RIC_LIQ_VEND = 0
    CREA_RECORDSET_TMP
End Sub
Private Sub CREA_RECORDSET_TMP()
On Error GoTo ERR_CREA_RECORDSET_TMP
Dim sSQL As String
Dim I As Long
Dim RsImp As ADODB.Recordset

Screen.MousePointer = 11

If Not (rsRigheElaTMP Is Nothing) Then
    Set rsRigheElaTMP = Nothing
End If

Set rsRigheElaTMP = New ADODB.Recordset
rsRigheElaTMP.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_POIELiquidazioneRigheElaModVel "
sSQL = sSQL & " WHERE IDRV_POLiquidazioneRigheEla=0"
sSQL = sSQL & " AND TipoRiga<>2"
Set RsImp = New ADODB.Recordset
RsImp.Open sSQL, Cn.InternalConnection

''''CREA TABELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For I = 0 To RsImp.Fields.Count - 1
    Select Case RsImp.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsRigheElaTMP.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, RsImp.Fields(I).DefinedSize, RsImp.Fields(I).Attributes
        Case adInteger
            rsRigheElaTMP.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsRigheElaTMP.Fields.Append RsImp.Fields(I).Name, RsImp.Fields(I).Type, , RsImp.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsRigheElaTMP.Fields.Append RsImp.Fields(I).Name, adBoolean, , RsImp.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsRigheElaTMP.Fields.Append RsImp.Fields(I).Name, adDouble, , RsImp.Fields(I).Attributes
    End Select
Next

RsImp.Close
Set RsImp = Nothing


rsRigheElaTMP.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable
rsRigheElaTMP.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIELiquidazioneRigheElaModVel "
sSQL = sSQL & " WHERE IDRV_POLiquidazione=" & LINK_LIQUIDAZIONE
sSQL = sSQL & " AND TipoRiga<>2"
Set RsImp = New ADODB.Recordset
RsImp.Open sSQL, Cn.InternalConnection

While Not RsImp.EOF
    rsRigheElaTMP.AddNew
        For I = 0 To RsImp.Fields.Count - 1
            rsRigheElaTMP.Fields(RsImp.Fields(I).Name).Value = RsImp.Fields(I).Value
        Next I
        rsRigheElaTMP!Selezionato = 0
    rsRigheElaTMP.Update
RsImp.MoveNext
Wend

RsImp.Close
Set RsImp = Nothing

Screen.MousePointer = 0

GET_GRIGLIA

Exit Sub
ERR_CREA_RECORDSET_TMP:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_TMP"

End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long
Dim sSQL As String

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = ""

If Len(Me.txtCodiceArticoloRic.Text) > 0 Then
    If Len(sSQL) > 0 Then sSQL = sSQL & " AND "
    
    sSQL = sSQL & "CodiceArticolo LIKE " & fnNormString("%" & Me.txtCodiceArticoloRic.Text & "%")
    
End If
If Len(Me.txtDescriArticoloRic.Text) > 0 Then
    If Len(sSQL) > 0 Then sSQL = sSQL & " AND "
    sSQL = sSQL & "Articolo LIKE " & fnNormString("%" & Me.txtDescriArticoloRic.Text & "%")
    
End If
If Len(Me.txtCatMercRic.Text) > 0 Then
    If Len(sSQL) > 0 Then sSQL = sSQL & " AND "
    sSQL = sSQL & "CategoriaMerceologica LIKE " & fnNormString("%" & Me.txtCatMercRic.Text & "%")
End If
If Me.chkNonScarti.Value = vbChecked Then
    If Len(sSQL) > 0 Then sSQL = sSQL & " AND "
    sSQL = sSQL & "TipoRiga <> 2"
End If

rsRigheElaTMP.Filter = sSQL

With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
    .LoadUserSettings
    
    
    Set cl = .ColumnsHeader.Add("Selezionato", "Sel.", dgBoolean, True, 1300, dgAligncenter)
        cl.Editable = True
    .ColumnsHeader.Add "IDRV_POLiquidazioneRigheEla", "IDRV_POLiquidazioneRigheEla", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POLiquidazione", "IDRV_POLiquidazione", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_POLiquidazionePeriodo", "IDRV_POLiquidazionePeriodo", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "TipoRiga", "TipoRiga", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDValoriOggettoDettaglioArticolo", "IDValoriOggettoDettaglioArticolo", dgInteger, False, 500, dgAlignRight
    
    Set cl = .ColumnsHeader.Add("QuantitaVenduta", "Quantità", dgDouble, True, 2000, dgAlignRight)
        'cl.Editable = True
        'cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."

    Set cl = .ColumnsHeader.Add("ImportoUnitarioDaReg", "Importo unitario", dgDouble, True, 2000, dgAlignRight)
        cl.Editable = True
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    
    Set cl = .ColumnsHeader.Add("ImponibileDaReg", "Imponibile", dgDouble, True, 2000, dgAlignRight)
        'cl.Editable = True
        'cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 3
        cl.FormatOptions.FormatNumericThousandSep = "."
    
    
    .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1100, dgAlignleft
    .ColumnsHeader.Add "Articolo", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
    .ColumnsHeader.Add "CodiceLottoArticolo", "Lotto di vendita", dgchar, False, 3000, dgAlignleft
    .ColumnsHeader.Add "CategoriaMerceologica", "Categoria merceologica", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "DataDocumentoVendita", "Data doc. vendita", dgDate, True, 2000, dgAlignleft
    .ColumnsHeader.Add "Oggetto", "Descrizione documento di vendita", dgchar, False, 2500, dgAlignleft
    .ColumnsHeader.Add "CodiceArticolo_Conf", "Codice articolo conf.", dgchar, False, 1100, dgAlignleft
    .ColumnsHeader.Add "Articolo_Conf", "Descrizione articolo conf.", dgchar, False, 3000, dgAlignleft
    .ColumnsHeader.Add "DataConferimento", "Data conferimento", dgDate, True, 2000, dgAlignleft
                
    Set .Recordset = rsRigheElaTMP
    
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaCorpo_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsRigheElaTMP.Fields("Selezionato").Value), 2
        End If
    End If
Exit Sub
ERR_GrigliaCorpo_KeyPress:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_KeyPress"
End Sub

Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaCorpo_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaCorpo.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaCorpo.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaCorpo.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsRigheElaTMP.Fields("Selezionato").Value), 2
            End If
        End If
    End If
    
Exit Sub
ERR_GrigliaCorpo_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_MouseUp"
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsRigheElaTMP.EOF And Not rsRigheElaTMP.BOF Then
        rsRigheElaTMP.Fields("Selezionato").Value = Abs(CLng(Selected))
        rsRigheElaTMP.UpdateBatch
        Me.GrigliaCorpo.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
Private Sub GrigliaCorpo_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
On Error GoTo ERR_GrigliaCorpo_AfterChangeFieldValue
    Select Case Column.FieldName
        Case "ImportoUnitarioDaReg"
            rsRigheElaTMP!ImponibileDaReg = Value * fnNotNullN(rsRigheElaTMP!QuantitaVenduta)
    End Select

    rsRigheElaTMP.UpdateBatch
    
    Me.GrigliaCorpo.Refresh
Exit Sub
ERR_GrigliaCorpo_AfterChangeFieldValue:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_AfterChangeFieldValue"
End Sub
Private Sub ELABORAZIONE_LIQUIDAZIONE(idliquidazione As Long)
On Error GoTo ERR_ELABORAZIONE_LIQUIDAZIONE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim NumeroRecord As Long
Dim UnitaProgresso As Double
Dim NumeroElaborazione As Long

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
NumeroRecord = 0
NumeroElaborazione = 1

sSQL = "SELECT COUNT(IDRV_POLiquidazioneRigheEla) as NRecord "
sSQL = sSQL & " FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & " WHERE IDRV_POLiquidazione=" & idliquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

If Not rs.EOF Then
    NumeroRecord = fnNotNullN(rs!NRecord)
End If

rs.Close
Set rs = Nothing

If NumeroRecord = 0 Then
    
    Exit Sub
End If

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 2)

sSQL = "SELECT * FROM RV_POLiquidazioneRigheEla"
sSQL = sSQL & " WHERE IDRV_POLiquidazione=" & idliquidazione

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


rsRigheElaTMP.Filter = vbNullString

Me.GrigliaCorpo.UpdatePosition = False

While Not rs.EOF
    Me.lblInfo.Caption = "ELABORAZIONE " & NumeroElaborazione & " di " & NumeroRecord
    DoEvents
    
    rsRigheElaTMP.Filter = "IDRV_POLiquidazioneRigheEla=" & fnNotNullN(rs!IDRV_POLiquidazioneRigheEla)
    
    If Not rsRigheElaTMP.EOF Then
        If fnNotNullN(rsRigheElaTMP!ImportoUnitarioDaReg) <> fnNotNullN(rs!ImportoUnitarioDaReg) Then
            rs!ImportoUnitarioDaReg = rsRigheElaTMP!ImportoUnitarioDaReg
            rs!ImponibileDaReg = rsRigheElaTMP!ImportoUnitarioDaReg * rs!QuantitaVenduta
            rs!ImpostaDaReg = (rsRigheElaTMP!ImponibileDaReg / 100) * fnNotNullN(rs!AliquotaIva_per_Imp_Vend)
            rs!ImportoLordoDaReg = rs!ImponibileDaReg + rs!ImpostaDaReg
            
            fncCalcoloTrattenute idliquidazione, rs
            
            rs.Update
            
        End If
    End If
    
    rsRigheElaTMP.Filter = vbNullString
        
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    NumeroElaborazione = NumeroElaborazione + 1
    DoEvents
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Me.GrigliaCorpo.UpdatePosition = True

CONFERMA_RIC_LIQ_VEND = 1

Unload Me

Exit Sub
ERR_ELABORAZIONE_LIQUIDAZIONE:
    MsgBox Err.Description, vbCritical, "ELABORAZIONE_LIQUIDAZIONE"

End Sub
Private Sub fncCalcoloTrattenute(IDLiqdazioneOLD As Long, rsRiga As ADODB.Recordset)
On Error GoTo ERR_fncCalcoloTrattenute
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsTratt As DmtOleDbLib.adoResultset
Dim Quantita_da_prendere_in_considerazione As Double
Dim Valore_TrattenutaPerLavorazione As Double
Dim Valore_TrattenutaGenerale As Double
Dim TrattValoreGen1 As Double
Dim TrattValoreGen2 As Double
Dim TrattPercGen1 As Double
Dim TrattPercGen2 As Double
Dim TrattValoreLav1 As Double
Dim TrattValoreLav2 As Double
Dim TrattPercLav1 As Double
Dim TrattPercLav2 As Double
Dim NumeroDecimali As Double
Dim TipoArrotondamento As Long

Quantita_da_prendere_in_considerazione = fnNotNullN(rsRiga!QuantitaVenduta)

sSQL = "SELECT * FROM RV_POIELiquidazioneRigheTratt "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiqdazioneOLD
sSQL = sSQL & " AND IDArticoloVendita=" & rsRiga("IDArticolo").Value
sSQL = sSQL & " AND IDTipoOggetto=" & rsRiga("IDTipoOggetto").Value
sSQL = sSQL & " AND IDOggetto=" & rsRiga("IDOggetto").Value
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & rsRiga("IDValoriOggettoDettaglioArticolo").Value
sSQL = sSQL & " ORDER BY TipoTrattenuta"

Set rs = Cn.OpenResultset(sSQL)

Valore_TrattenutaPerLavorazione = 0
Valore_TrattenutaGenerale = 0
TrattValoreGen1 = 0
TrattValoreGen2 = 0
TrattPercGen1 = 0
TrattPercGen2 = 0
TrattValoreLav1 = 0
TrattValoreLav2 = 0
TrattPercLav1 = 0
TrattPercLav2 = 0

While Not rs.EOF
    If ((fnNotNullN(rs!IDRV_POTipoTrattenuta) = 4) Or (fnNotNullN(rs!IDRV_POTipoTrattenuta) = 7) Or (fnNotNullN(rs!IDRV_POTipoTrattenuta) = 9) Or (fnNotNullN(rs!IDRV_POTipoTrattenuta) = 10)) Then
        'Trattetenute per lavorazione
        TrattValoreLav1 = TrattValoreLav1 + (Quantita_da_prendere_in_considerazione * fnNotNullN(rs!ValoreTrattenuta1))
        TrattValoreLav2 = TrattValoreLav2 + (Quantita_da_prendere_in_considerazione * fnNotNullN(rs!ValoreTrattenuta2))
        TrattPercLav1 = TrattPercLav1 + ((fnNotNullN(rsRiga!ImponibileDaReg) / 100) * fnNotNullN(rs!PercTrattenuta1))
        TrattPercLav2 = TrattPercLav2 + ((fnNotNullN(rsRiga!ImponibileDaReg) / 100) * fnNotNullN(rs!PercTrattenuta2))
    Else
        'Trattenute generali
        TrattValoreGen1 = TrattValoreGen1 + (Quantita_da_prendere_in_considerazione * fnNotNullN(rs!ValoreTrattenuta1))
        TrattValoreGen2 = TrattValoreGen2 + (Quantita_da_prendere_in_considerazione * fnNotNullN(rs!ValoreTrattenuta2))
        TrattPercGen1 = TrattPercGen1 + ((fnNotNullN(rsRiga!ImponibileDaReg) / 100) * fnNotNullN(rs!PercTrattenuta1))
        TrattPercGen2 = TrattPercGen2 + ((fnNotNullN(rsRiga!ImponibileDaReg) / 100) * fnNotNullN(rs!PercTrattenuta2))
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

'TrattValoreLav1 = fnRoundChange(TrattValoreLav1, NumeroDecimali, TipoArrotondamento)
'TrattValoreLav2 = fnRoundChange(TrattValoreLav2, NumeroDecimali, TipoArrotondamento)
'TrattPercLav1 = fnRoundChange(TrattPercLav1, NumeroDecimali, TipoArrotondamento)
'TrattPercLav2 = fnRoundChange(TrattPercLav2, NumeroDecimali, TipoArrotondamento)
'
'TrattValoreGen1 = fnRoundChange(TrattValoreGen1, NumeroDecimali, TipoArrotondamento)
'TrattValoreGen2 = fnRoundChange(TrattValoreGen2, NumeroDecimali, TipoArrotondamento)
'TrattPercGen1 = fnRoundChange(TrattPercGen1, NumeroDecimali, TipoArrotondamento)
'TrattPercGen2 = fnRoundChange(TrattPercGen2, NumeroDecimali, TipoArrotondamento)



'Totali trattenute
Valore_TrattenutaPerLavorazione = TrattValoreLav1 + TrattValoreLav2 + TrattPercLav1 + TrattPercLav2
Valore_TrattenutaGenerale = TrattValoreGen1 + TrattValoreGen2 + TrattPercGen1 + TrattPercGen2

rsRiga!TrattenutaValoreGen1 = TrattValoreGen1
rsRiga!TrattenutaValoreGen2 = TrattValoreGen2
rsRiga!TrattenutaPercGen1 = TrattPercGen1
rsRiga!TrattenutaPercGen2 = TrattPercGen2
rsRiga!TrattenutaValoreLav1 = TrattValoreLav1
rsRiga!TrattenutaValoreLav2 = TrattValoreLav2
rsRiga!TrattenutaPercLav1 = TrattPercLav1
rsRiga!TrattenutaPercLav2 = TrattPercLav2

rsRiga!TrattenuteGenerali = Valore_TrattenutaGenerale
rsRiga!TrattenutePerLavorazione = Valore_TrattenutaPerLavorazione
rsRiga!TrattenuteTotali = rsRiga!TrattenuteGenerali + rsRiga!TrattenutePerLavorazione

Exit Sub
ERR_fncCalcoloTrattenute:
    MsgBox Err.Description, vbCritical, "fncCalcoloTrattenute"
End Sub

