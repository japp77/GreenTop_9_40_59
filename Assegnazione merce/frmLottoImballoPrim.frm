VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmLottoImballoPrim 
   Caption         =   "LOTTO IMBALLO PRIMARIO"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLottoImballoPrim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   17385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiine 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   17055
      Begin VB.CommandButton cmdConferma 
         Caption         =   "CONFERMA"
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
         Left            =   15600
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "ANNULLA"
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
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmLottoImballoPrim.frx":4781A
         Top             =   0
         Width           =   13815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   17175
      Begin VB.CommandButton cmdSelAut 
         Caption         =   "SELEZIONE AUTOMATICA"
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "RESET"
         Height          =   495
         Left            =   15240
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
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
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   17175
      Begin DMTEDITNUMLib.dmtNumber txtQtaSel 
         Height          =   375
         Left            =   15120
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaDaEvadere 
         Height          =   375
         Left            =   15120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
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
      Begin DMTEDITNUMLib.dmtNumber txtQtaDiff 
         Height          =   375
         Left            =   15120
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità selezionata"
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
         Left            =   6840
         TabIndex        =   6
         Top             =   195
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantità da evadere"
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
         Left            =   6840
         TabIndex        =   5
         Top             =   660
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Differenza"
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
         Left            =   6840
         TabIndex        =   4
         Top             =   1125
         Width           =   8175
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   6165
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
Attribute VB_Name = "frmLottoImballoPrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsGrigliaSomma As ADODB.Recordset

Private CONFERMA_RIPRISTINA As Long


Private Sub cmdConferma_Click()
Dim Testo As String

    If Me.txtQtaDiff.Value > 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La quantità selezionata è maggiore della quantità da evadere" & vbCrLf
        Testo = Testo & "Controllare i valori"
        MsgBox Testo, vbCritical, "Controllo valori"
        Exit Sub
    End If
   
    If Me.txtQtaDiff.Value < 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La quantità selezionata è minore della quantità da evadere" & vbCrLf
        Testo = Testo & "Continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo valori") = vbNo Then Exit Sub
    End If
    
    If CONTROLLO_RIMANENZA_NEGATIVA = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Uno o più lotti selezionati hanno una rimanenza negativa" & vbCrLf
        Testo = Testo & "Controllare i valori"
        MsgBox Testo, vbCritical, "Controllo valori"
        Exit Sub
    End If

    CONFERMA
    
    If CONFERMA_RIPRISTINA = 1 Then
        CONFERMA_LOTTO_IMBALLO_DA_UTENTE_PRIM = 0
    Else
        CONFERMA_LOTTO_IMBALLO_DA_UTENTE_PRIM = 1
    End If
    
    
    Unload Me
    
End Sub

Private Sub cmdReset_Click()
    Dim Testo As String
    
    Testo = "Con questo comando le quantità selezionate verranno valorizzate a zero" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "RESET") = vbNo Then Exit Sub
    
    
    
    RIPRISTINA_RESET
End Sub

Private Sub cmdRipristina_Click()
    Dim Testo As String
    
    Testo = "Con questo comando verrà ripristinata la situazione iniziale" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "RIPRISTINO") = vbNo Then Exit Sub

    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI
End Sub

Private Sub cmdSelAut_Click()
Dim Testo As String
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double



Testo = "Con questo comando verrà ripristinata la situazione iniziale e poi le quantità dei lotti verranno ricalcolate automaticatimente " & vbCrLf
Testo = Testo & "Vuoi continuare?"

If MsgBox(Testo, vbQuestion + vbYesNo, "SELEZIONE AUTOMATICA") = vbNo Then Exit Sub


RIPRISTINA_RESET

QuantitaRimasta = Me.txtQtaDaEvadere.Value

rsGriglia.Filter = "Giacenza>0"
rsGriglia.Sort = "NumeroProgressivo"

While Not rsGriglia.EOF
    If QuantitaRimasta > 0 Then
        If (QuantitaRimasta - fnNotNullN(rsGriglia!Giacenza)) <= 0 Then
            QuantitaUtilizzata = QuantitaRimasta
        Else
            QuantitaUtilizzata = fnNotNullN(rsGriglia!Giacenza)
        End If
        
        rsGriglia!QuantitaSelezionata = QuantitaUtilizzata
        rsGriglia!Rimanenza = rsGriglia!Giacenza - rsGriglia!QuantitaSelezionata
        rsGriglia!Registra = 1
        
        QuantitaRimasta = QuantitaRimasta - QuantitaUtilizzata

    End If
    
rsGriglia.MoveNext
Wend

rsGriglia.Filter = vbNullString


If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    rsGriglia.MoveFirst
    While Not rsGriglia.EOF
        rsGrigliaSomma.Filter = "IDRV_POLottoImballo=" & fnNotNullN(rsGriglia!IDRV_POLottoImballo)
        For I = 0 To rsGriglia.Fields.Count - 1
            rsGrigliaSomma.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
        Next
        rsGrigliaSomma.Update
        rsGrigliaSomma.Filter = vbNullString
        
    rsGriglia.MoveNext
    Wend
    
    rsGriglia.MoveFirst

End If

GET_GRIGLIA

GET_TOTALI

CONFERMA_RIPRISTINA = 1


End Sub
Private Sub RIPRISTINA_RESET()
Dim I As Integer
    
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        rsGriglia.MoveFirst
        
        While Not rsGriglia.EOF
            rsGriglia!Registra = 0
            rsGriglia!QuantitaSelezionata = 0
            rsGriglia!Rimanenza = rsGriglia!Giacenza - rsGriglia!QuantitaSelezionata
            rsGriglia.Update
        rsGriglia.MoveNext
        Wend
        
        rsGriglia.MoveFirst
        While Not rsGriglia.EOF
            rsGrigliaSomma.Filter = "IDRV_POLottoImballo=" & fnNotNullN(rsGriglia!IDRV_POLottoImballo)
            For I = 0 To rsGriglia.Fields.Count - 1
                rsGrigliaSomma.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
            Next
            rsGrigliaSomma.Update
            rsGrigliaSomma.Filter = vbNullString
            
        rsGriglia.MoveNext
        Wend
        
        rsGriglia.MoveFirst
    End If
    
    


    GET_GRIGLIA
    
    GET_TOTALI
End Sub
Private Sub RIPRISTINA()
    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    CONFERMA_RIPRISTINA = 0
    
    
    
    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI
End Sub
Private Sub GET_GRIGLIA()
'On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader


rsGriglia.Filter = "Giacenza>0"


OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        Set cl = .ColumnsHeader.Add("Registra", "Seleziona", dgBoolean, True, 1800, dgAligncenter)
            cl.Editable = True

        .ColumnsHeader.Add "IDRV_POLottoImballo", "IDRV_POLottoImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDArticoloImballo", "IDArticoloImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "LottoImballo", "Lotto", dgchar, True, 3500, dgAlignleft
        .ColumnsHeader.Add "RiferimentoEsterno", "Riferimento esterno", dgchar, False, 3500, dgAlignleft
        Set cl = .ColumnsHeader.Add("QuantitaCaricata", "Q.tà Caricata", dgDouble, True, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Giacenza", "Disponibilità", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("QuantitaSelezionata", "Q.tà selezionata", dgDouble, True, 2000, dgAlignRight)
            cl.Editable = True
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
            cl.BackColor = vbYellow
            
        Set cl = .ColumnsHeader.Add("Rimanenza", "Rimanenza", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor

End Sub
Private Sub CONFERMA()
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        rsLottoImballoPrim.MoveFirst
        
        While Not rsLottoImballoPrim.EOF
            rsLottoImballoPrim.Delete
        rsLottoImballoPrim.MoveNext
        Wend

       rsGriglia.MoveFirst
       While Not rsGriglia.EOF
           rsLottoImballoPrim.AddNew
                For I = 0 To rsGriglia.Fields.Count - 1
                    rsLottoImballoPrim.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
                Next
                
           rsLottoImballoPrim.Update
           
       rsGriglia.MoveNext
       Wend
       
       rsGriglia.MoveFirst
    
    End If
    


End Sub
Private Sub GET_RECORDSET_GRIGLIA()
Dim I As Integer

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    
    Set rsGriglia = Nothing
    
End If

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

If Not (rsGrigliaSomma Is Nothing) Then
    If rsGrigliaSomma.State > 0 Then
        rsGrigliaSomma.Close
    End If
    
    Set rsGrigliaSomma = Nothing
    
End If

Set rsGrigliaSomma = New ADODB.Recordset
rsGrigliaSomma.CursorLocation = adUseClient

For I = 0 To rsLottoImballoPrim.Fields.Count - 1
    rsGriglia.Fields.Append rsLottoImballoPrim.Fields(I).Name, rsLottoImballoPrim.Fields(I).Type, rsLottoImballoPrim.Fields(I).DefinedSize, adFldIsNullable
    rsGrigliaSomma.Fields.Append rsLottoImballoPrim.Fields(I).Name, rsLottoImballoPrim.Fields(I).Type, rsLottoImballoPrim.Fields(I).DefinedSize, adFldIsNullable
    
Next

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
rsGrigliaSomma.Open , , adOpenKeyset, adLockBatchOptimistic

If Not ((rsLottoImballoPrim.EOF) And (rsLottoImballoPrim.BOF)) Then
    rsLottoImballoPrim.MoveFirst
    
    While Not rsLottoImballoPrim.EOF
        rsGriglia.AddNew
        rsGrigliaSomma.AddNew
            For I = 0 To rsLottoImballoPrim.Fields.Count - 1
                rsGriglia.Fields(rsLottoImballoPrim.Fields(I).Name).Value = rsLottoImballoPrim.Fields(I).Value
                rsGrigliaSomma.Fields(rsLottoImballoPrim.Fields(I).Name).Value = rsLottoImballoPrim.Fields(I).Value
                
            Next
        rsGriglia.Update
        rsGrigliaSomma.Update
    rsLottoImballoPrim.MoveNext
    Wend
    
    
End If


End Sub
Private Sub GET_TOTALI()
Me.txtQtaSel.Value = 0
Me.txtQtaDaEvadere.Value = frmMain.txtColli.Value * frmMain.txtNumeroConfImballo.Value
Me.txtQtaDiff.Value = 0

rsGrigliaSomma.Filter = "Registra=1 AND QuantitaSelezionata>0"


While Not rsGrigliaSomma.EOF
    Me.txtQtaSel.Value = Me.txtQtaSel.Value + fnNotNullN(rsGrigliaSomma!QuantitaSelezionata)
    
rsGrigliaSomma.MoveNext
Wend

rsGrigliaSomma.Filter = vbNullString

Me.txtQtaDiff.Value = Me.txtQtaSel.Value - Me.txtQtaDaEvadere.Value

If Me.txtQtaDiff.Value <> 0 Then
    Me.txtQtaDiff.BackColor = vbRed
    Me.Label1(2).ForeColor = vbRed
Else
    Me.txtQtaDiff.BackColor = vbWhite
    Me.Label1(2).ForeColor = vbBlack
End If
End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
Dim I As Integer

    If rsGriglia!QuantitaSelezionata = 0 Then
        rsGriglia!Registra = 0
    Else
        rsGriglia!Registra = 1
    End If

    rsGriglia!Rimanenza = rsGriglia!Giacenza - rsGriglia!QuantitaSelezionata
    
    Me.Griglia.Refresh
    rsGriglia.Update
    
    rsGrigliaSomma.Filter = "IDRV_POLottoImballo=" & fnNotNullN(rsGriglia!IDRV_POLottoImballo)
    If Not rsGrigliaSomma.EOF Then
        For I = 0 To rsGriglia.Fields.Count - 1
            rsGrigliaSomma.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
        Next
        
        rsGrigliaSomma.Update
    End If
    rsGrigliaSomma.Filter = vbNullString
    
    
    
    
    GET_TOTALI

End Sub

Private Function CONTROLLO_RIMANENZA_NEGATIVA() As Boolean
Dim ris As Boolean
ris = False

If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        If ris = False Then
            If rsGriglia!Rimanenza < 0 Then
                ris = True
            End If
        End If
    rsGriglia.MoveNext
    Wend
    
    rsGriglia.MoveFirst

End If

CONTROLLO_RIMANENZA_NEGATIVA = ris

End Function


