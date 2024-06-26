VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmLottoImballo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOTTO IMBALLO"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
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
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frmLottoImballo.frx":0000
      Top             =   6000
      Width           =   9975
   End
   Begin VB.CommandButton cmdRipristina 
      Caption         =   "RIPRISTINA"
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   1935
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
      TabIndex        =   7
      Top             =   4200
      Width           =   13695
      Begin DMTEDITNUMLib.dmtNumber txtQtaSel 
         Height          =   375
         Left            =   11640
         TabIndex        =   8
         Top             =   240
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
         Left            =   11640
         TabIndex        =   9
         Top             =   720
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
         Left            =   11640
         TabIndex        =   10
         Top             =   1200
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
         Left            =   3360
         TabIndex        =   13
         Top             =   1245
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantit� da evadere"
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
         Left            =   3360
         TabIndex        =   12
         Top             =   780
         Width           =   8175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantit� selezionata"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   315
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton cmdReset 
         Caption         =   "RESET"
         Height          =   495
         Left            =   11640
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelAut 
         Caption         =   "SELEZIONE AUTOMATICA"
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame fraFiine 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   13695
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
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
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
         Left            =   12120
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13695
      _ExtentX        =   24156
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
Attribute VB_Name = "frmLottoImballo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsGrigliaSomma As ADODB.Recordset

Private CONFERMA_RIPRISTINA As Long


Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdConferma_Click()
Dim Testo As String

    If Me.txtQtaDiff.Value > 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La quantit� selezionata � maggiore della quantit� da evadere" & vbCrLf
        Testo = Testo & "Controllare i valori"
        MsgBox Testo, vbCritical, "Controllo valori"
        Exit Sub
    End If
   
    If Me.txtQtaDiff.Value < 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La quantit� selezionata � minore della quantit� da evadere" & vbCrLf
        Testo = Testo & "Continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo valori") = vbNo Then Exit Sub
    End If
    
    If CONTROLLO_RIMANENZA_NEGATIVA = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Uno o pi� lotti selezionati hanno una rimanenza negativa" & vbCrLf
        Testo = Testo & "Controllare i valori"
        MsgBox Testo, vbCritical, "Controllo valori"
        Exit Sub
    End If

    Conferma
    
    If CONFERMA_RIPRISTINA = 1 Then
        CONFERMA_LOTTO_IMBALLO_DA_UTENTE = 0
    Else
        CONFERMA_LOTTO_IMBALLO_DA_UTENTE = 1
    End If
    
    
    Unload Me
    
End Sub

Private Sub cmdReset_Click()
    Dim Testo As String
    
    Testo = "Con questo comando le quantit� selezionate verranno valorizzate a zero" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "RESET") = vbNo Then Exit Sub
    
    
    
    RIPRISTINA_RESET
End Sub

Private Sub cmdRipristina_Click()
    Dim Testo As String
    
    Testo = "Con questo comando verr� ripristinata la situazione iniziale" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "RIPRISTINO") = vbNo Then Exit Sub

    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI
End Sub

Private Sub cmdSelAut_Click()
On Error GoTo ERR_cmdSelAut_Click
Dim Testo As String
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double



Testo = "Con questo comando verr� ripristinata la situazione iniziale e poi le quantit� dei lotti verranno ricalcolate automaticatimente " & vbCrLf
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

Exit Sub
ERR_cmdSelAut_Click:
    MsgBox Err.Description, vbCritical, "cmdSelAut_Click"
    
End Sub
Private Sub RIPRISTINA_RESET()
On Error GoTo ERR_RIPRISTINA_RESET
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
Exit Sub
ERR_RIPRISTINA_RESET:
    MsgBox Err.Description, vbCritical, "RIPRISTINA_RESET"
    
End Sub
Private Sub RIPRISTINA()
    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    CONFERMA_RIPRISTINA = 0
    
    
    
    GET_RECORDSET_GRIGLIA

    GET_GRIGLIA
    
    GET_TOTALI

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
    
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
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
        Set cl = .ColumnsHeader.Add("QuantitaCaricata", "Q.t� Caricata", dgDouble, True, 1800, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Giacenza", "Disponibilit�", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("QuantitaSelezionata", "Q.t� selezionata", dgDouble, True, 2000, dgAlignRight)
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
        .ColumnsHeader.Add "RiferimentoEsterno", "Rif. esterno", dgchar, True, 3500, dgAlignleft
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With
Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub
Private Sub Conferma()
On Error GoTo ERR_CONFERMA
    If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
        rsLottoImballo.MoveFirst
        
        While Not rsLottoImballo.EOF
            rsLottoImballo.Delete
        rsLottoImballo.MoveNext
        Wend

       rsGriglia.MoveFirst
       While Not rsGriglia.EOF
           rsLottoImballo.AddNew
                For I = 0 To rsGriglia.Fields.Count - 1
                    rsLottoImballo.Fields(rsGriglia.Fields(I).Name).Value = rsGriglia.Fields(I).Value
                Next
                
           rsLottoImballo.Update
           
       rsGriglia.MoveNext
       Wend
       
       rsGriglia.MoveFirst
    
    End If
    

Exit Sub
ERR_CONFERMA:
    MsgBox Err.Description, vbCritical, "CONFERMA"
    
End Sub
Private Sub GET_RECORDSET_GRIGLIA()
On Error GoTo ERR_GET_RECORDSET_GRIGLIA
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






For I = 0 To rsLottoImballo.Fields.Count - 1
    rsGriglia.Fields.Append rsLottoImballo.Fields(I).Name, rsLottoImballo.Fields(I).Type, rsLottoImballo.Fields(I).DefinedSize, adFldIsNullable
    rsGrigliaSomma.Fields.Append rsLottoImballo.Fields(I).Name, rsLottoImballo.Fields(I).Type, rsLottoImballo.Fields(I).DefinedSize, adFldIsNullable
    
Next

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
rsGrigliaSomma.Open , , adOpenKeyset, adLockBatchOptimistic

If Not ((rsLottoImballo.EOF) And (rsLottoImballo.BOF)) Then
    rsLottoImballo.MoveFirst
    
    While Not rsLottoImballo.EOF
        rsGriglia.AddNew
        rsGrigliaSomma.AddNew
            For I = 0 To rsLottoImballo.Fields.Count - 1
                rsGriglia.Fields(rsLottoImballo.Fields(I).Name).Value = rsLottoImballo.Fields(I).Value
                rsGrigliaSomma.Fields(rsLottoImballo.Fields(I).Name).Value = rsLottoImballo.Fields(I).Value
                
            Next
        rsGriglia.Update
        rsGrigliaSomma.Update
    rsLottoImballo.MoveNext
    Wend
    
    
End If

Exit Sub
ERR_GET_RECORDSET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_RECORDSET_GRIGLIA"
    
End Sub
Private Sub GET_TOTALI()
On Error GoTo ERR_GET_TOTALI
Me.txtQtaSel.Value = 0
Me.txtQtaDaEvadere.Value = frmMain.txtQuantitaImballo.Value
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
Exit Sub
ERR_GET_TOTALI:
    MsgBox Err.Description, vbCritical, "GET_TOTALI"
    
End Sub

Private Sub Griglia_AfterChangeFieldValue(ByVal Column As dmtgridctl.dgColumnHeader, ByVal Value As Variant)
On Error GoTo ERR_Griglia_AfterChangeFieldValue
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
Exit Sub
ERR_Griglia_AfterChangeFieldValue:
    MsgBox Err.Description, vbCritical, "Griglia_AfterChangeFieldValue"
    
End Sub

Private Function CONTROLLO_RIMANENZA_NEGATIVA() As Boolean
On Error GoTo ERR_CONTROLLO_RIMANENZA_NEGATIVA
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
Exit Function
ERR_CONTROLLO_RIMANENZA_NEGATIVA:
    MsgBox Err.Description, vbCritical, "CONTROLLO_RIMANENZA_NEGATIVA"
    
End Function

