VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form frmImputazionePrezzi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMPUTAZIONE PREZZI"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13650
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
   ScaleHeight     =   8895
   ScaleWidth      =   13650
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
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8835
      ScaleWidth      =   13515
      TabIndex        =   2
      Top             =   0
      Width           =   13575
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
         Height          =   1575
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   13455
         Begin VB.CheckBox chkAggiornaPrezziAZero 
            Caption         =   "AGGIORNA IMPORTI A ZERO"
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
            Left            =   3600
            TabIndex        =   9
            Top             =   840
            Width           =   2055
         End
         Begin VB.CheckBox chkAggiornaPrezzoImballoDaListino 
            Caption         =   "Aggiorna l'importo unitario imballo da listino quando è 0 (zero)"
            Height          =   495
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   840
            Width           =   3375
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
            Height          =   375
            Left            =   9480
            TabIndex        =   7
            Top             =   1080
            Width           =   3255
         End
         Begin VB.CheckBox chkMerceInclusoImballo 
            Caption         =   "MERCE INCLUSO IMBALLO"
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
            Left            =   5880
            TabIndex        =   6
            Top             =   840
            Width           =   2295
         End
         Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
            Height          =   375
            Left            =   3600
            TabIndex        =   10
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
            Left            =   5880
            TabIndex        =   11
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
            Left            =   8160
            TabIndex        =   13
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
            Left            =   10440
            TabIndex        =   14
            Top             =   480
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
         Begin VB.Label Label2 
            Caption         =   "Sconto 2 %"
            Height          =   255
            Index           =   3
            Left            =   10440
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Sconto 1 %"
            Height          =   255
            Index           =   2
            Left            =   8160
            TabIndex        =   18
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Listino imballi"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario imballo"
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Importo unitario articolo"
            Height          =   255
            Index           =   0
            Left            =   3600
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
         Height          =   975
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   13455
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
         Left            =   11160
         TabIndex        =   3
         Top             =   8280
         Width           =   2295
      End
      Begin DmtGridCtl.DmtGrid GrigliaCorpo 
         Height          =   5535
         Left            =   0
         TabIndex        =   4
         Top             =   2520
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   9763
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
         Left            =   120
         TabIndex        =   24
         Top             =   8400
         Width           =   10575
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

    rsNew.Filter = sSQL
    
    'rsNew.Requery
    
    If Not (rsNew.EOF And rsNew.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings

End Sub

Private Sub CDImballo_ChangeElement()
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

    rsNew.Filter = sSQL

    
    'rsNew.Requery
    
    If Not (rsNew.EOF And rsNew.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings
End Sub

Private Sub CDSocio_ChangeElement()
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

    rsNew.Filter = sSQL

    
    
    
    If Not (rsNew.EOF And rsNew.BOF) Then
        Me.GrigliaCorpo.Requery
    End If
    Me.GrigliaCorpo.LoadUserSettings
End Sub

Private Sub cmdAggiorna_Click()
Dim ImportoImballo As Double

    If Not (rsNew.BOF And rsNew.EOF) Then
        rsNew.MoveFirst
        
        While Not rsNew.EOF
            If Me.txtImportoUnitarioArticolo.Value > 0 Then
                If Me.chkAggiornaPrezziAZero.Value = vbChecked Then
                    If fnNotNullN(rsNew("Art_prezzo_unitario_neutro").Value) = 0 Then
                        rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                    End If
                Else
                    rsNew("Art_prezzo_unitario_neutro").Value = Me.txtImportoUnitarioArticolo.Value
                End If
            End If
            
            If Me.txtSconto1.Value > 0 Then
                rsNew("Art_sco_in_percentuale_1").Value = Me.txtSconto1.Value
            End If
            If Me.txtSconto2.Value > 0 Then
                rsNew("Art_sco_in_percentuale_2").Value = Me.txtSconto2.Value
            End If
            
            rsNew("RV_POImportoImballoInArticolo").Value = Me.chkMerceInclusoImballo.Value
           
           
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
            


            ImportoImballo = fnNotNullN(rsNew!RV_POImportoUnitarioImballo)
        
            If rsNew.Fields("RV_POImportoImballoInArticolo").Value = False Then
                rsNew.Fields("RV_POImportoUnitarioImballo").Value = GET_PREZZO_IMBALLO(ImportoImballo)
            Else
                rsNew.Fields("RV_POImportoUnitarioImballo").Value = 0
            End If
        
            rsNew!RV_POImportoDaLiq = sbCalcolaImportoVariazioneLiquidazione(ImportoImballo)
        

            rsNew.UpdateBatch
            DoEvents
        rsNew.MoveNext
        Wend

    End If
    
End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
Dim I As Integer

rsNew.UpdateBatch

rsNew.Filter = adFilterNone
Me.Command1.Enabled = False
Me.lblInfo.Caption = "AGGIORNAMENTO IN CORSO........................."
DoEvents

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

While Not rsNew.EOF
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
DoEvents
'oDoc.PerformTable sTabellaDettaglio, True
DoEvents
rsNew.MoveNext
Wend

rsNew.Close

Set rsNew = Nothing
DoEvents
oDoc.PerformTable sTabellaDettaglio, True
Me.lblInfo.Caption = ""
Unload Me
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
    INIT_CONTROLLI
    CREA_TABELLA_TEMPORANEA
    GET_GRIGLIA
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

sSQL = "SELECT * FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
'sSQL = sSQL & " AND RV_PORigaCompleta=1"

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rsNew.Open , , adOpenKeyset, adLockBatchOptimistic
rsClone.Open , , adOpenKeyset, adLockBatchOptimistic
While Not rsImp.EOF
    rsNew.AddNew
    rsClone.AddNew
        For I = 0 To rsImp.Fields.Count - 1
            rsNew.Fields(rsImp.Fields(I).Name).Value = rsImp.Fields(I).Value
        Next
    rsNew.Update
    rsClone.Update
rsImp.MoveNext
Wend

rsImp.Close
Set rsImp = Nothing
End Sub
Private Sub GET_GRIGLIA()
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
                .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
                .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
                
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


                Set cl = .ColumnsHeader.Add("Art_pre_uni_net_sco_net_IVA", "Importo netto IVA", dgDouble, True, 1300, dgAlignRight)
                    'cl.Editable = True
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
                Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioImballo", "Importo unitario", dgDouble, True, 1300, dgAlignRight)
                    cl.Editable = True
                    cl.BackColor = vbYellow
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericCurSymbol = ""
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
       
                    
        Set .Recordset = rsNew
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
    LINK_UM_COOP_ARTICOLO = 0
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
       sbCalcolaImportoVariazioneLiquidazione = sbCalcolaImportoVariazioneLiquidazione / MOLTIPLICATORE_ARTICOLO
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
    sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
    sSQL = sSQL & "WHERE ("
    sSQL = sSQL & "(IDListino=" & frmMain.cboListino.CurrentID & ") "
    sSQL = sSQL & "AND (IDArticolo=" & rsNew!RV_POIDImballo & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIVA)
    Else
        GET_PREZZO_IMBALLO = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
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
