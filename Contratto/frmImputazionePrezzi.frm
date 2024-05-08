VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmImputazionePrezzi 
   Caption         =   "IMPUTAZIONE PREZZI"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   20715
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
   ScaleHeight     =   10545
   ScaleWidth      =   20715
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   7440
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      ScaleHeight     =   10425
      ScaleWidth      =   20490
      TabIndex        =   20
      Top             =   0
      Width           =   20520
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "RICERCA"
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
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   8580
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Frame Frame1 
         Caption         =   "Parametri di selezione delle righe ordine"
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
         Height          =   975
         Left            =   9360
         TabIndex        =   34
         Top             =   120
         Width           =   11040
         Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
            Height          =   615
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
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
            Left            =   5280
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
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
      End
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
         ForeColor       =   &H00FF0000&
         Height          =   9615
         Left            =   6120
         TabIndex        =   22
         Top             =   120
         Width           =   3135
         Begin VB.CheckBox chkAggiornaPrezziAZero 
            Caption         =   "AGGIORNA IMPORTI A ZERO"
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   3360
            Width           =   2895
         End
         Begin VB.CheckBox chkAggiornaPrezzoImballoDaListino 
            Caption         =   "Aggiorna l'importo unitario imballo da listino quando è 0 (zero)"
            Height          =   615
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   4320
            Width           =   2775
         End
         Begin VB.CommandButton cmdAggiorna 
            Caption         =   "AGGIORNA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   9000
            Width           =   2895
         End
         Begin VB.CheckBox chkMerceInclusoImballo 
            Caption         =   "MERCE INCLUSO IMBALLO"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   3840
            Width           =   2775
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   2880
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
            TabIndex        =   5
            TabStop         =   0   'False
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
         Begin DMTEDITNUMLib.dmtNumber txtSconto1 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   2280
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
            Height          =   315
            Left            =   1680
            TabIndex        =   9
            Top             =   2280
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
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
            TabIndex        =   6
            TabStop         =   0   'False
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
         Begin VB.Label Label4 
            Caption         =   "Listino merce"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 2 %"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 1 %"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Listino imballi"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario imballo"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   2640
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario merce"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   2775
         End
      End
      Begin VB.Frame fraRicerca 
         Caption         =   "PARAMETRI DI SELEZIONE ORDINI"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   5895
         Begin VB.CommandButton Command3 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4320
            TabIndex        =   37
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   36
            Top             =   480
            Width           =   375
         End
         Begin DmtCodDescCtl.DmtCodDesc CDSocio 
            Height          =   615
            Left            =   15360
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
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
         Begin DMTDATETIMELib.dmtDate txtDataInizio 
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   480
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDataFine 
            Height          =   315
            Left            =   2880
            TabIndex        =   1
            Top             =   480
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboDestDiversa 
            Height          =   315
            Left            =   3840
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
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
         Begin VB.Label Label1 
            Caption         =   "Destinazione diversa"
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Data fine"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Data inizio"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
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
         Height          =   495
         Left            =   18120
         TabIndex        =   17
         Top             =   9840
         Width           =   2295
      End
      Begin DmtGridCtl.DmtGrid GrigliaCorpo 
         Height          =   8535
         Left            =   9360
         TabIndex        =   21
         Top             =   1200
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   15055
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9360
         TabIndex        =   35
         Top             =   9960
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmImputazionePrezzi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsImp As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim rsClone As ADODB.Recordset

Private Sub CDArticolo_ChangeElement()
    SetGrigliaRigheOrdine

End Sub

Private Sub CDImballo_ChangeElement()
    SetGrigliaRigheOrdine
End Sub

Private Sub cmdAggiorna_Click()
On Error GoTo ERR_cmdAggiorna_Click
Dim ImportoImballo As Double

    If Not (rsNew.BOF And rsNew.EOF) Then
        rsNew.MoveFirst
        
        While Not rsNew.EOF
            If Me.txtImportoUnitarioArticolo.Value > 0 Then
                If Me.chkAggiornaPrezziAZero.Value = vbChecked Then
                    If fnNotNullN(rsNew("Art_prezzo_unitario_neutro").Value) = 0 Then
                        If (Me.cboListinoMerce.CurrentID = 0) Then
                            rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                        Else
                            rsNew("Art_prezzo_unitario_neutro").Value = GET_PREZZO_ARTICOLO(fnNotNullN(rsNew!Link_Art_articolo), Me.cboListino.CurrentID)
                        End If
                    End If
                Else
                    If (Me.cboListinoMerce.CurrentID = 0) Then
                        rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                    Else
                        rsNew("Art_prezzo_unitario_neutro").Value = GET_PREZZO_ARTICOLO(fnNotNullN(rsNew!Link_Art_articolo), Me.cboListino.CurrentID)
                    End If
                End If
            End If
            
            If Me.txtSconto1.Value > 0 Then
                rsNew("Art_sco_in_percentuale_1").Value = Me.txtSconto1.Value
            End If
            If Me.txtSconto2.Value > 0 Then
                rsNew("Art_sco_in_percentuale_2").Value = Me.txtSconto2.Value
            End If
            
            rsNew("RV_POImportoImballoInArticolo").Value = Me.chkMerceInclusoImballo.Value
            
            If chkAggiornaPrezzoImballoDaListino.Value = vbChecked Then
                If fnNotNullN(rsNew.Fields("ImportoUnitarioImballo").Value) = 0 Then
                    If Me.cboListino.CurrentID > 0 Then
                        rsNew.Fields("ImportoUnitarioImballo").Value = GET_PREZZO_ARTICOLO(fnNotNullN(rsNew!RV_POIDImballo), Me.cboListino.CurrentID)
                    Else
                        rsNew.Fields("ImportoUnitarioImballo").Value = Me.txtImportoUnitarioImballo.Value
                    End If
                End If
            Else
                If Me.cboListino.CurrentID > 0 Then
                    rsNew.Fields("ImportoUnitarioImballo").Value = GET_PREZZO_ARTICOLO(fnNotNullN(rsNew!RV_POIDImballo), Me.cboListino.CurrentID)
                Else
                    rsNew.Fields("ImportoUnitarioImballo").Value = Me.txtImportoUnitarioImballo.Value
                End If
            End If
            
            rsNew.UpdateBatch
            DoEvents
        rsNew.MoveNext
        Wend

    End If
Exit Sub
ERR_cmdAggiorna_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdAggiorna_Click"
End Sub

Private Sub cmdRicerca_Click()
    AVVIA_RICERCA_ORDINI
End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
Dim I As Integer
Dim rsOrdini As ADODB.Recordset
Dim ObjectDoc As DmtDocs.cDocument
Dim rsRigaOrdineSel As ADODB.Recordset
Dim Filtro As String

'Impostazione dell'oggetto
Set ObjectDoc = New DmtDocs.cDocument
Set ObjectDoc.Connection = Cn
ObjectDoc.UseAutomation = True 'Dico di effettuare il ricalcolo dei campi automaticamente
ObjectDoc.SetTipoOggetto 15
ObjectDoc.IDFunzione = 128
ObjectDoc.IDAzienda = TheApp.IDFirm
ObjectDoc.IDFiliale = TheApp.Branch
ObjectDoc.InitAmbientVariables Nothing
 



rsNew.UpdateBatch


Me.Command1.Enabled = False

If Not (rsNew.BOF And rsNew.EOF) Then
    Me.lblInfo.Caption = "RECUPERO DATI IN CORSO..."
    DoEvents
    
    Set rsOrdini = New ADODB.Recordset
    rsOrdini.CursorLocation = adUseClient
    rsOrdini.Fields.Append "IDOrdine", adInteger, , adFldIsNullable
    rsOrdini.Open , , adOpenKeyset, adLockBatchOptimistic
    
    rsNew.MoveFirst
    
    While Not rsNew.EOF
        rsOrdini.Filter = "IDOrdine=" & rsNew!IDOggetto
        If rsOrdini.EOF Then
            rsOrdini.AddNew
                rsOrdini!IDOrdine = rsNew!IDOggetto
            rsOrdini.Update
        End If
        rsOrdini.Filter = ""
        DoEvents
    rsNew.MoveNext
    Wend

    rsNew.MoveFirst
    Filtro = rsNew.Filter
    
    If Not (rsClone.BOF And rsClone.EOF) Then
        rsClone.MoveFirst
        While Not rsClone.EOF
            rsClone.Delete
        rsClone.MoveNext
        Wend
    End If
        
    While Not rsNew.EOF
        rsClone.AddNew
            For I = 0 To rsNew.Fields.Count - 1
                rsClone.Fields(rsNew.Fields(I).Name).Value = rsNew.Fields(I).Value
            Next
        rsClone.Update
    rsNew.MoveNext
    Wend
    
    If Not (rsOrdini.BOF And rsOrdini.EOF) Then
        rsOrdini.MoveFirst
        While Not rsOrdini.EOF
            rsClone.Filter = "IDOggetto = " & rsOrdini!IDOrdine
            
            While Not rsClone.EOF
                Me.lblInfo.Caption = "AGGIORNAMENTO DATI IN CORSO ORDINE NUMERO " & fnNotNull(rsClone!Doc_numero) & " del " & fnNotNull(rsClone!Doc_data)
                DoEvents
                
                'AGGIORNAMENTO PREZZO MERCE
                sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
                sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & rsClone!IDValoriOggettoDettaglio
                
                Set rsRigaOrdineSel = New ADODB.Recordset
                rsRigaOrdineSel.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
                
                If Not rsRigaOrdineSel.EOF Then
                    rsRigaOrdineSel!Art_prezzo_unitario_neutro = rsClone!Art_prezzo_unitario_neutro
                    rsRigaOrdineSel!Art_sco_in_percentuale_1 = rsClone!Art_sco_in_percentuale_1
                    rsRigaOrdineSel!Art_sco_in_percentuale_2 = rsClone!Art_sco_in_percentuale_2
                    rsRigaOrdineSel!RV_POImportoImballoInArticolo = rsClone!RV_POImportoImballoInArticolo
                    rsRigaOrdineSel.Update
                End If
                
                rsRigaOrdineSel.Close
                Set rsRigaOrdineSel = Nothing
                                
                                
                'AGGIORNAMENTO PREZZO IMBALLO
                sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
                sSQL = sSQL & "WHERE IDOggetto=" & rsOrdini!IDOrdine
                sSQL = sSQL & " AND RV_POTipoRiga=2 "
                sSQL = sSQL & " AND RV_POLinkRiga=" & rsClone!RV_POLinkRiga
                
                Set rsRigaOrdineSel = New ADODB.Recordset
                rsRigaOrdineSel.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
                
                If Not rsRigaOrdineSel.EOF Then
                    rsRigaOrdineSel!Art_prezzo_unitario_neutro = rsClone!ImportoUnitarioImballo
                    rsRigaOrdineSel!RV_POImportoImballoInArticolo = rsClone!RV_POImportoImballoInArticolo
                    rsRigaOrdineSel.Update
                End If
                
                rsRigaOrdineSel.Close
                Set rsRigaOrdineSel = Nothing
                
                'AGGIORNAMENTO LAVORAZIONI COLLEGATE ALLA RIGA ORDINE
                sSQL = "UPDATE RV_POAssegnazioneMerce SET "
                sSQL = sSQL & "ImportoUnitarioArticolo=" & fnNormNumber(rsClone!Art_prezzo_unitario_neutro) & ", "
                sSQL = sSQL & "ImportoUnitarioImballo=" & fnNormNumber(rsClone!ImportoUnitarioImballo) & ", "
                sSQL = sSQL & "MerceInclusoImballo=" & fnNormBoolean(rsClone!RV_POImportoImballoInArticolo) & ", "
                sSQL = sSQL & "Sconto1=" & fnNormNumber(rsClone!Art_sco_in_percentuale_1) & ", "
                sSQL = sSQL & "Sconto2=" & fnNormNumber(rsClone!Art_sco_in_percentuale_2)
                sSQL = sSQL & " WHERE IDValoriOggettoDettaglioRigaOrd=" & rsClone!IDValoriOggettoDettaglio
                Cn.Execute sSQL
                
                DoEvents
                rsClone.MoveNext
            Wend
            rsClone.Filter = ""
            
            ObjectDoc.ReadWithTO rsOrdini!IDOrdine, 15
            ObjectDoc.PerformDocument Nothing, True
            ObjectDoc.Update
        rsOrdini.MoveNext
        Wend
    End If
    Set ObjectDoc = Nothing
End If


Me.lblInfo.Caption = "OPERAZIONE AVVENUTA CON SUCCESSO!"
DoEvents

GrigliaCorpo.Refresh
Me.Command1.Enabled = True
Me.lblInfo.Caption = ""
DoEvents


MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO!", vbInformation, "Aggiornamento prezzi"
Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Aggiornamento prezzi"
    Me.Command1.Enabled = True

End Sub
Private Sub SCRIVI_RIGA_IMBALLO()
Dim sSQL As String

rsClone.Filter = "RV_POTipoRiga=1 AND RV_POLinkRiga=" & rsNew!RV_POLinkRiga

If Not rsClone.EOF Then



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

Private Sub Command2_Click()
    Me.txtDataInizio.Value = Me.txtDataInizio.Value - 7
    Me.txtDataFine.Value = Me.txtDataInizio.Value + 6
End Sub

Private Sub Command3_Click()
    Me.txtDataInizio.Value = Me.txtDataInizio.Value + 7
    Me.txtDataFine.Value = Me.txtDataInizio.Value + 6
End Sub

Private Sub Form_Load()
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
    
    

    Me.txtDataInizio.Text = GetWeekStartDate(DatePart("ww", Date), DatePart("yyyy", Date))
    Me.txtDataFine.Value = Me.txtDataInizio.Value + 6

    INIT_CONTROLLI
    CREA_TABELLA_TEMPORANEA
    
    AVVIA_RICERCA_ORDINI
    
'    GET_GRIGLIA
    Me.cboListino.WriteOn GET_LISTINO_DEFAULT(frmMain.cdAnagrafica.KeyFieldID)
    If Me.cboListino.CurrentID > 0 Then
        Me.chkAggiornaPrezzoImballoDaListino.Value = vbChecked
    End If
    Me.chkAggiornaPrezziAZero.Value = vbChecked
    Me.chkMerceInclusoImballo.Value = GET_MERCE_INCLUSO_IMBALLO_CLIENTE(frmMain.cdAnagrafica.KeyFieldID)

End Sub
Private Sub CREA_TABELLA_TEMPORANEA()
Dim sSQL As String
Dim I As Long

sSQL = "SELECT * FROM RV_POIEOrdineRigheDaContratto "
sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=0"

Set rsImp = New ADODB.Recordset
Set rsNew = New ADODB.Recordset
Set rsClone = New ADODB.Recordset

rsImp.Open sSQL, Cn.InternalConnection

rsNew.CursorLocation = adUseClient
rsClone.CursorLocation = adUseClient

''''CREA TABELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For I = 0 To rsImp.Fields.Count - 1
    Select Case rsImp.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar
            rsNew.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, rsImp.Fields(I).DefinedSize, rsImp.Fields(I).Attributes
            rsClone.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, rsImp.Fields(I).DefinedSize, rsImp.Fields(I).Attributes
        
        Case adInteger
            rsNew.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
            rsClone.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
        
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsNew.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
            rsClone.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
        
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsNew.Fields.Append rsImp.Fields(I).Name, adBoolean, , rsImp.Fields(I).Attributes
            rsClone.Fields.Append rsImp.Fields(I).Name, adBoolean, , rsImp.Fields(I).Attributes
        
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsNew.Fields.Append rsImp.Fields(I).Name, adDouble, , rsImp.Fields(I).Attributes
            rsClone.Fields.Append rsImp.Fields(I).Name, adDouble, , rsImp.Fields(I).Attributes
    End Select
Next

rsImp.Close
Set rsImp = Nothing

rsNew.Fields.Append "Registra", adBoolean, , adFldIsNullable
rsNew.Fields.Append "ImportoUnitarioImballo", adDouble, , adFldIsNullable

rsClone.Fields.Append "Registra", adBoolean, , adFldIsNullable
rsClone.Fields.Append "ImportoUnitarioImballo", adDouble, , adFldIsNullable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rsNew.Open , , adOpenKeyset, adLockBatchOptimistic
rsClone.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_CURSOR As Long

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

rsNew.Filter = "RV_POTipoRiga=1"

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
        .ColumnsHeader.Add "RV_POLinkRiga", "RV_POLinkRiga", dgInteger, False, 500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDOggettoContratto", "RV_POIDOggettoContratto", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Link_Nom_anagrafica", "IDAnagraficaCliente", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Nom_codice", "Codice cliente", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Nom_ragione_sociale_o_cognome", "Anagrafica cliente", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Nom_nome", "Nome cliente", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Link_Nom_ult_sito", "IDDestinazioneDiversa", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagrafica", "Distinazione diversa", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Doc_ordine_chiuso", "Ordine chiuso", dgBoolean, False, 1100, dgAligncenter
        .ColumnsHeader.Add "RV_POIDOrdinePadre", "IDOggettoOrdinePadre", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_PONumeroOrdinePadre", "N° ordine padre", dgNumeric, False, 1100, dgAlignRight
        .ColumnsHeader.Add "RV_PODataOrdinePadre", "Data ordine padre", dgDate, False, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PONumeroListaPrelievo", "N° lista prelievo", dgInteger, False, 1100, dgAlignRight
        .ColumnsHeader.Add "Doc_data_prevista_evasione", "Data consegna", dgDate, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Doc_numero_presso_nom", "Numero ordine cliente", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Doc_data_presso_nom", "Data ordine cliente", dgDate, False, 1100, dgAlignleft
        .ColumnsHeader.Add "Doc_numero", "N° ordine", dgNumeric, False, 1100, dgAlignRight
        .ColumnsHeader.Add "Doc_data", "Data ordine", dgDate, False, 1100, dgAlignleft
        
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
        Set cl = .ColumnsHeader.Add("Art_numero_colli", "Colli", dgDouble, True, 1300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "Link_Art_unita_di_misura", "Link_Art_unita_di_misura", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "UnitaDiMisura", "U.M.", dgchar, True, 1100, dgAlignleft
        Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà", dgDouble, True, 1300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        
        Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
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
            
        .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, True, 3000, dgAlignleft

        Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Importo imballo", dgDouble, True, 1300, dgAlignRight)
            cl.Editable = True
            cl.BackColor = vbRed
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POImportoImballoInArticolo", "Incluso Imballo", dgBoolean, True, 1300, dgAligncenter)
            cl.Editable = True
            cl.BackColor = vbYellow
            
        .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Sub lotto", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Codice tipo pedana", dgchar, False, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione tipo pedana", dgchar, False, 1100, dgAlignleft
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedanaEffettiva", "N° ped. effettive", dgDouble, False, 1300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POColliSfusi", "Colli sfusi", dgDouble, False, 1300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POColliPerPedana", "Colli per pedana", dgDouble, False, 1300, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

        
        'RIFERIMENTO ORDINE

        Set .Recordset = rsNew
        .Refresh
        .LoadUserSettings
    End With

Cn.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Form_Resize()
On Error GoTo ERR_Form_Resize
  If Me.WindowState <> 1 Then
        
        If Me.Width > 19200 Then
            'GESTIONE LARGHEZZA
            Me.Pic1.Width = Me.Width - 240
            Me.GrigliaCorpo.Width = Me.Pic1.Width - 240 - (Me.Frame2.Left + Me.Frame2.Width + 120)
            Me.Frame1.Width = Me.GrigliaCorpo.Width
            Me.Command1.Left = (Me.Pic1.Width - 240) - Me.Command1.Width
            
        End If
        If Me.Height > 11000 Then
            'GESTIONE LUNGHEZZA
            Me.Pic1.Height = Me.Height - 840
            Me.Command1.Top = (Me.Pic1.Height - 120) - Me.Command1.Height
            Me.lblInfo.Top = Me.Command1.Top
            Me.GrigliaCorpo.Height = (Me.Command1.Top - 120) - Me.GrigliaCorpo.Top
            Me.List1.Height = (Me.Command1.Top - 120) - Me.List1.Top
            Me.Frame2.Height = (Me.Command1.Top - 120) - Me.Frame2.Top
            Me.cmdAggiorna.Top = Me.Frame2.Height - 120 - Me.cmdAggiorna.Height
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
Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"
End Sub

Private Sub GrigliaCorpo_AfterChangeFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, ByVal Value As Variant)
'    If rsNew!RV_POTipoRiga = 1 Then
'        Select Case Column.FieldName
'            Case "Art_prezzo_unitario_neutro"
'                rsNew!Art_pre_uni_net_sco_net_IVA = Value - ((Value / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
'                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'
'                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
'                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'
'                rsNew!Art_prezzo_unitario_netto_IVA = Value
'                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
'                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
'                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
'
'            Case "Art_sco_in_percentuale_1"
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_prezzo_unitario_neutro - ((rsNew!Art_prezzo_unitario_neutro / 100) * Value)
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
'                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'
'                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (Value + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
'                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'
'                rsNew!Art_prezzo_unitario_netto_IVA = rsNew!Art_prezzo_unitario_neutro
'                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
'                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
'                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
'
'            Case "Art_sco_in_percentuale_2"
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_prezzo_unitario_neutro - ((rsNew!Art_prezzo_unitario_neutro / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * Value)
'                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'
'                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + Value))
'                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'
'                rsNew!Art_prezzo_unitario_netto_IVA = rsNew!Art_prezzo_unitario_neutro
'                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
'                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
'                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
'            Case "RV_POImportoImballoInArticolo"
'
'        End Select
'
'    End If
'
'
'    If rsNew!RV_POTipoRiga = 2 Then
'        Select Case Column.FieldName
'            Case "Art_prezzo_unitario_neutro"
'                rsNew!Art_pre_uni_net_sco_net_IVA = Value - ((Value / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_1))
'                rsNew!Art_pre_uni_net_sco_net_IVA = rsNew!Art_pre_uni_net_sco_net_IVA - ((fnNotNullN(rsNew!Art_pre_uni_net_sco_net_IVA) / 100) * fnNotNullN(rsNew!Art_sco_in_percentuale_2))
'                rsNew!Art_pre_uni_net_sco_lor_IVA = rsNew!Art_pre_uni_net_sco_net_IVA + ((rsNew!Art_pre_uni_net_sco_net_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_lordo_IVA = rsNew!Art_pre_uni_net_sco_lor_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_importo_totale_netto_IVA = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'
'                rsNew!Art_importo_sconto_netto_IVA = ((rsNew!Art_pre_uni_net_sco_lor_IVA / 100) * (fnNotNullN(rsNew!Art_sco_in_percentuale_1) + fnNotNullN(rsNew!Art_sco_in_percentuale_2)))
'                rsNew!Art_importo_sconto_lordo_IVA = rsNew!Art_importo_sconto_netto_IVA + ((rsNew!Art_importo_sconto_netto_IVA / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'
'                rsNew!Art_prezzo_unitario_netto_IVA = Value
'                rsNew!Art_prezzo_unitario_lordo_IVA = fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) + ((fnNotNullN(rsNew!Art_prezzo_unitario_netto_IVA) / 100) * fnNotNullN(rsNew!Art_aliquota_iva))
'                rsNew!Art_importo_totale_neutro = rsNew!Art_pre_uni_net_sco_net_IVA * fnNotNullN(rsNew!Art_quantita_totale)
'                rsNew!Art_Importo_netto_IVA = rsNew!Art_importo_totale_netto_IVA
'                rsNew!Art_importo_net_sconto_lor_IVA = rsNew!Art_importo_totale_lordo_IVA
'                rsNew!Art_importo_net_sconto_net_IVA = rsNew!Art_importo_totale_netto_IVA
'        End Select
'
'    End If
'
'
'    rsNew.UpdateBatch
'
'    Me.GrigliaCorpo.Refresh

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
        

        
        rsNew.UpdateBatch
                
        Me.GrigliaCorpo.Refresh

    End If

End Sub
Private Function GET_PREZZO_ARTICOLO(IDArticolo As Long, IDListino As Long) As Double
On Error GoTo ERR_GET_PREZZO_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & " WHERE IDListino=" & IDListino
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_PREZZO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIVA)
Else
    GET_PREZZO_ARTICOLO = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PREZZO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_PREZZO_ARTICOLO"
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

    'Listino
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

    'Destinazione diversa
    With Me.cboDestDiversa
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica "
        .SQL = .SQL & "FROM SitoPerAnagrafica "
        .SQL = .SQL & "WHERE IDAnagrafica=" & frmMain.cdAnagrafica.KeyFieldID
        .Fill
    End With
    


End Sub

Private Sub GrigliaOrdine_Click()

End Sub



Private Sub List1_ItemCheck(Item As Integer)
    SetGrigliaRigheOrdine
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

sSQL = "SELECT IDListinoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    
    sSQL = "SELECT IDListinoDiBase "
    sSQL = sSQL & "FROM ConfigurazioneVendite "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_LISTINO_DEFAULT = 0
    Else
        GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
    End If
Else
    If fnNotNullN(rs!IDListinoDefault) = 0 Then
        rs.CloseResultset
        Set rs = Nothing
        
        sSQL = "SELECT IDListinoDiBase "
        sSQL = sSQL & "FROM ConfigurazioneVendite "
        sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF Then
            GET_LISTINO_DEFAULT = 0
        Else
            GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
        End If
    Else
        GET_LISTINO_DEFAULT = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
    End If
End If

rs.CloseResultset
Set rs = Nothing

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


Private Function GetWeekStartDate(weekNumber As Integer, year As Integer) As String

    Dim startDate As String
    Dim day As Integer

    startDate = DateSerial(year, 1, 1)
    day = Weekday(startDate, vbMonday)
    startDate = DateAdd("d", DaysToAdd(day), startDate)

    GetWeekStartDate = DateAdd("ww", weekNumber - 1, startDate)

End Function

Private Function DaysToAdd(day As Integer) As Integer

    DaysToAdd = 0
    If day > 1 Then DaysToAdd = 7 - day + 1

End Function
Private Sub AVVIA_RICERCA_ORDINI()
On Error GoTo ERR_AVVIA_RICERCA_ORDINI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer
Dim DescrizioneOrdine As String

Me.GrigliaCorpo.UpdatePosition = False

If Not (rsNew.BOF And rsNew.EOF) Then
    rsNew.MoveFirst
    While Not rsNew.EOF
        rsNew.Delete
    rsNew.MoveNext
    Wend
End If

Me.GrigliaCorpo.UpdatePosition = True

Me.List1.Clear

sSQL = "SELECT * FROM RV_POIEOrdineDaContratto "
sSQL = sSQL & " WHERE Link_nom_anagrafica=" & frmMain.cdAnagrafica.KeyFieldID
sSQL = sSQL & " AND IDOggetto=RV_POIDOrdinePadre "
sSQL = sSQL & " AND RV_POIDOggettoContratto=" & oDoc.IDOggetto
sSQL = sSQL & " AND Doc_ordine_chiuso=0"
If Me.txtDataInizio.Value > 0 Then
    sSQL = sSQL & " AND RV_PODataOrdinePadre>=" & fnNormDate(Me.txtDataInizio.Text)
End If
If Me.txtDataFine.Value > 0 Then
    sSQL = sSQL & " AND RV_PODataOrdinePadre<=" & fnNormDate(Me.txtDataFine.Text)
End If
If Me.cboDestDiversa.CurrentID > 0 Then
    sSQL = sSQL & " AND Link_Nom_ult_sito=" & Me.cboDestDiversa.CurrentID
End If

Set rs = Cn.OpenResultset(sSQL)

Screen.MousePointer = 11
I = 0
While Not rs.EOF
    DescrizioneOrdine = GET_NUMERO_ORDINE(fnNotNull(rs!RV_PONumeroOrdinePadre), fnNotNull(Trim(rs!prefisso))) & " del " & GET_DATA_FORMATTATA(rs!RV_PODataOrdinePadre)
    If Len(Trim(fnNotNull(rs!Doc_numero_presso_nom))) > 0 Then
        DescrizioneOrdine = DescrizioneOrdine & " - " & rs!Doc_numero_presso_nom
    End If
    If fnNotNullN(rs!Link_Nom_ult_sito) > 0 Then
        DescrizioneOrdine = DescrizioneOrdine & " (" & rs!SitoPerAnagrafica & ")"
    End If
    Me.List1.AddItem DescrizioneOrdine
    Me.List1.ItemData(I) = fnNotNullN(rs!IDOggetto)
    
    GET_RIGHE_ORDINE fnNotNullN(rs!IDOggetto)
    
    DoEvents
I = I + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Screen.MousePointer = 0

GET_GRIGLIA
Exit Sub
ERR_AVVIA_RICERCA_ORDINI:
    MsgBox Err.Description, vbCritical, "AVVIA_RICERCA_ORDINI"
    Screen.MousePointer = 0
End Sub
Private Function GET_NUMERO_ORDINE(numero As String, prefisso As String) As String
Dim I As Integer
Dim StringaNumeroZeri As String
StringaNumeroZeri = ""
For I = Len(numero) To 6
    StringaNumeroZeri = StringaNumeroZeri + "0"
Next

GET_NUMERO_ORDINE = StringaNumeroZeri & numero

If Len(Trim(prefisso)) > 0 Then
    GET_NUMERO_ORDINE = Trim(prefisso) & GET_NUMERO_ORDINE
End If

End Function
Private Function GET_DATA_FORMATTATA(data As String) As String
Dim Anno As String
Dim mese As String
Dim giorno As String

Anno = DatePart("yyyy", data)
mese = Month(data)
giorno = day(data)

If Len(mese) = 1 Then mese = "0" & mese
If Len(giorno) = 1 Then giorno = "0" & giorno

GET_DATA_FORMATTATA = giorno & "/" & mese & "/" & Anno
End Function
Private Sub GET_RIGHE_ORDINE(IDOggettoOrdine As Long)
Dim rsImp As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT * FROM RV_POIEOrdineRigheDaContratto "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1 "
Set rsImp = New ADODB.Recordset
rsImp.Open sSQL, Cn.InternalConnection

While Not rsImp.EOF
    rsNew.AddNew
    rsClone.AddNew
        For I = 0 To rsImp.Fields.Count - 1
            rsNew.Fields(rsImp.Fields(I).Name).Value = rsImp.Fields(I).Value
        Next
        rsNew!Registra = 0
        rsNew!ImportoUnitarioImballo = GET_IMPORTO_IMBALLO_RIGA_ORDINE(IDOggettoOrdine, fnNotNullN(rsNew!RV_POLinkRiga))
    rsNew.Update
    rsClone.Update
rsImp.MoveNext
Wend

rsImp.Close
Set rsImp = Nothing
End Sub
Private Function GET_IMPORTO_IMBALLO_RIGA_ORDINE(IDOggetto As Long, linkriga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IMPORTO_IMBALLO_RIGA_ORDINE = 0

sSQL = "SELECT IDValoriOggettoDettaglio, Art_prezzo_unitario_neutro "
sSQL = sSQL & "FROM RV_POIEOrdineRigheDaContratto "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & linkriga
sSQL = sSQL & " AND RV_POTipoRiga=2"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_IMPORTO_IMBALLO_RIGA_ORDINE = fnNotNullN(rs!Art_prezzo_unitario_neutro)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub SetGrigliaRigheOrdine()
On Error GoTo ERR_SetGrigliaRigheOrdine
Dim sSQL As String
Dim sSQL_OR As String
Dim sSQL_std As String
Dim I As Integer

sSQL_std = "(RV_POTipoRiga=1)"

If Me.CDArticolo.KeyFieldID > 0 Then
    sSQL_std = sSQL_std & " AND (Link_art_articolo=" & Me.CDArticolo.KeyFieldID & ")"
End If
If Me.CDImballo.KeyFieldID > 0 Then
    sSQL_std = sSQL_std & " AND (RV_POIDImballo=" & Me.CDImballo.KeyFieldID & ")"
End If

sSQL = ""

For I = 0 To List1.ListCount - 1
    If (List1.Selected(I) = True) Then
        If Len(sSQL) > 0 Then
            sSQL_OR = sSQL_OR & " OR "
        End If
        sSQL = "(" & sSQL_std & " AND (RV_POIDOrdinePadre=" & List1.ItemData(I) & "))"
        sSQL_OR = sSQL_OR & sSQL
    End If
Next
If Len(sSQL_OR) > 0 Then
    rsNew.Filter = sSQL_OR
Else
    rsNew.Filter = sSQL_std
End If
If Not (rsNew.EOF And rsNew.BOF) Then
    Me.GrigliaCorpo.Requery
End If

Me.GrigliaCorpo.LoadUserSettings
Exit Sub
ERR_SetGrigliaRigheOrdine:
    MsgBox Err.Description, vbCritical, "SetGrigliaRigheOrdine"
End Sub

