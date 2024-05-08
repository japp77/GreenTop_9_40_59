VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmAltriDatilavorazione 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTRI DATI"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAltriDatilavorazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Riferimenti produzione"
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
      TabIndex        =   20
      Top             =   3840
      Width           =   5655
      Begin VB.TextBox txtRifProcessoProd 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   5415
      End
      Begin DMTDataCmb.DMTCombo cboLineaProduzione 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Riferimento processo produzione"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label11 
         Caption         =   "Linea di produzione"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   5415
      End
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
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   "ALTRI DATI"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtImballoPrimInLinguaPred 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox txtImballoPrimInLinguaCliente 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox txtImballoInLinguaPred 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox txtDescrCodBarreImpPrimAz 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtDescrCodBarreImbPrimCli 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtCodBarreImbPrimAz 
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtCodBarreImbPrimCli 
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3480
         Width           =   2655
      End
      Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerCollo 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   465
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   0
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoPerCollo 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   465
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtMoltiplicatore 
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   465
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà pezzi per collo"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Peso per collo"
         Height          =   255
         Index           =   24
         Left            =   1920
         TabIndex        =   18
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Moltiplicatore"
         Height          =   255
         Index           =   31
         Left            =   3360
         TabIndex        =   17
         ToolTipText     =   "Numero di confezioni in un imballo"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Descrizione imballo primario in lingua predefinita"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   5415
      End
      Begin VB.Label Label11 
         Caption         =   "Descrizione imballo primario in lingua cliente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label11 
         Caption         =   "Descrizione imballo in lingua predefinita"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "Codice a barre azienda dell'imballo primario"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "Codice a barre cliente dell'imballo primario"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmAltriDatilavorazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub PRELEVA_DATI()
    
    Me.txtQuantitaPerCollo.Value = QUANTITA_PER_COLLO
    Me.txtPesoPerCollo.Value = PESO_LORDO_ARTICOLO
    Me.txtMoltiplicatore.Value = MOLTIPLICATORE
    
    Me.txtImballoInLinguaPred.Text = DESCR_IMB_LINGUA_PRED
    Me.txtImballoPrimInLinguaCliente.Text = DESCR_IMB_PRIM_LINGUA_CLIENTE
    Me.txtImballoPrimInLinguaPred.Text = DESCR_IMB_PRIM_LINGUA_PRED
    
    Me.txtDescrCodBarreImbPrimCli.Text = DESCR_A_BARRE_IMB_PRIM_CLI
    Me.txtCodBarreImbPrimCli.Text = COD_A_BARRE_IMB_PRIM_CLI
    
    Me.txtDescrCodBarreImpPrimAz.Text = DESCR_A_BARRE_IMB_PRIM_PRED
    Me.txtCodBarreImbPrimAz.Text = COD_A_BARRE_IMB_PRIM_PRED
    Me.cboLineaProduzione.WriteOn IDLineaProduzione
    Me.txtRifProcessoProd.Text = GET_DESCRIZIONE_PROCESSO_PROD(IDProcessoProd)
End Sub

Private Sub CONFERMA_DATI()

    DESCR_IMB_LINGUA_PRED = Me.txtImballoInLinguaPred.Text
    DESCR_IMB_PRIM_LINGUA_CLIENTE = Me.txtImballoPrimInLinguaCliente.Text
    DESCR_IMB_PRIM_LINGUA_PRED = Me.txtImballoPrimInLinguaPred.Text
    
    DESCR_A_BARRE_IMB_PRIM_CLI = Me.txtDescrCodBarreImbPrimCli.Text
    DESCR_A_BARRE_IMB_PRIM_PRED = Me.txtDescrCodBarreImpPrimAz.Text
    
    COD_A_BARRE_IMB_PRIM_CLI = GET_CODICEABARRE(DESCR_A_BARRE_IMB_PRIM_CLI)
    COD_A_BARRE_IMB_PRIM_PRED = GET_CODICEABARRE(DESCR_A_BARRE_IMB_PRIM_PRED)

    QUANTITA_PER_COLLO = Me.txtQuantitaPerCollo.Value
    PESO_LORDO_ARTICOLO = Me.txtPesoPerCollo.Value
    MOLTIPLICATORE = Me.txtMoltiplicatore.Value
    IDLineaProduzione = Me.cboLineaProduzione.CurrentID
    
    Unload Me
    
End Sub
Private Sub cmdConferma_Click()
    CONFERMA_DATI
    
End Sub



Private Sub Form_Load()
     With Me.cboLineaProduzione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "ID"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POLineaProduzione ORDER BY Descrizione"
        .Fill
    End With
    

    PRELEVA_DATI
End Sub

Private Function GET_CODICEABARRE(CodiceABarre As String) As String
Dim CodiceABarreDaPassare As String

If Len(Trim(CodiceABarre)) = 0 Then
    GET_CODICEABARRE = ""
Else
    If Len((fnNotNull(Trim(CodiceABarre)))) <= 11 Then
        GET_CODICEABARRE = ""
    End If
    If Len((fnNotNull(Trim(CodiceABarre)))) > 13 Then
        GET_CODICEABARRE = ""
    End If
    If Len((fnNotNull(Trim(CodiceABarre)))) = 13 Then
        CodiceABarreDaPassare = Mid(fnNotNull(Trim(CodiceABarre)), 1, Len(fnNotNull(Trim(CodiceABarre))) - 1)
        GET_CODICEABARRE = ean13$(CodiceABarreDaPassare)
    End If
    If Len((fnNotNull(Trim(CodiceABarre)))) = 12 Then
        GET_CODICEABARRE = ean13$((fnNotNull(Trim(CodiceABarre))))
    End If
End If


End Function


Public Function ean13$(chaine$)

  Dim I%, checksum%, first%, CodeBarre$, tableA As Boolean
  ean13$ = ""
  If Len(chaine$) = 12 Then
    For I% = 1 To 12
      If Asc(Mid$(chaine$, I%, 1)) < 48 Or Asc(Mid$(chaine$, I%, 1)) > 57 Then
        I% = 0
        Exit For
      End If
    Next
    If I% = 13 Then
      For I% = 12 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, I%, 1))
      Next
      checksum% = checksum% * 3
      For I% = 11 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, I%, 1))
      Next
      chaine$ = chaine$ & (10 - checksum% Mod 10) Mod 10
      CodeBarre$ = Left$(chaine$, 1) & Chr$(65 + Val(Mid$(chaine$, 2, 1)))
      first% = Val(Left$(chaine$, 1))
      For I% = 3 To 7
        tableA = False
         Select Case I%
         Case 3
           Select Case first%
           Case 0 To 3
             tableA = True
           End Select
         Case 4
           Select Case first%
           Case 0, 4, 7, 8
             tableA = True
           End Select
         Case 5
           Select Case first%
           Case 0, 1, 4, 5, 9
             tableA = True
           End Select
         Case 6
           Select Case first%
           Case 0, 2, 5, 6, 7
             tableA = True
           End Select
         Case 7
           Select Case first%
           Case 0, 3, 6, 8, 9
             tableA = True
           End Select
         End Select
       If tableA Then
         CodeBarre$ = CodeBarre$ & Chr$(65 + Val(Mid$(chaine$, I%, 1)))
       Else
         CodeBarre$ = CodeBarre$ & Chr$(75 + Val(Mid$(chaine$, I%, 1)))
       End If
     Next
      CodeBarre$ = CodeBarre$ & "*"   'Ajout séparateur central / Add middle separator
      For I% = 8 To 13
        CodeBarre$ = CodeBarre$ & Chr$(97 + Val(Mid$(chaine$, I%, 1)))
      Next
      CodeBarre$ = CodeBarre$ & "+"   'Ajout de la marque de fin / Add end mark
      ean13$ = CodeBarre$
    End If
  End If
End Function
Private Function GET_DESCRIZIONE_PROCESSO_PROD(ID As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_DESCRIZIONE_PROCESSO_PROD = ""
sSQL = "SELECT * FROM RV_POProcessoLavorazione "
sSQL = sSQL & "WHERE ID=" & ID

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESCRIZIONE_PROCESSO_PROD = "Processo di lavorazione n° " & fnNotNullN(rs!Anno) & "-" & fnNotNullN(rs!ID) & " del " & fnNotNull(rs!DataOra)
End If


rs.CloseResultset
Set rs = Nothing
End Function

