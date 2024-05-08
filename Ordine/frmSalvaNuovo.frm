VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Begin VB.Form frmSalvaNuovo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salva come nuovo"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalvaNuovo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdSalva 
         Caption         =   "SALVA"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   3120
         Width           =   1575
      End
      Begin DmtCodDescCtl.DmtCodDesc cdAnagrafica 
         Height          =   585
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   1032
         PropCodice      =   $"frmSalvaNuovo.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmSalvaNuovo.frx":4787E
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmSalvaNuovo.frx":478CF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DMTDataCmb.DMTCombo cboListino 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2520
         _ExtentX        =   4445
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
      Begin DMTDataCmb.DMTCombo cboPagamento 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   1680
         Width           =   3240
         _ExtentX        =   5715
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
      Begin DMTDataCmb.DMTCombo cboAltroSito 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDATETIMELib.dmtDate dtData 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber lngNumero 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1680
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboMagazzino 
         Height          =   315
         Left            =   2760
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Magazzino"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Doc."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "Altra destinazione"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "Pagamento"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblDocument 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listino cliente"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmSalvaNuovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    INIT_CONTROLLI
    
End Sub

Private Sub INIT_CONTROLLI()
    'Inizializza la combo dei pagamenti
    With cboPagamento
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDPagamento"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Pagamento"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDPagamento, Pagamento FROM Pagamento"
        .SQL = .SQL & " ORDER BY Pagamento"
    End With

    With Me.cboMagazzino
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDMagazzino"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Magazzino"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDMagazzino, Magazzino FROM Magazzino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " ORDER BY Magazzino"
    End With
    'Inizializza la combo dei listini
    With cboListino
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDListino"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Listino"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDListino, Listino FROM Listino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino = 0"
        .SQL = .SQL & " ORDER BY Listino"
    End With

    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With cdAnagrafica
        'Imposta l'oggetto Application del modello ad oggetti
        Set .Application = TheApp
        'Imposta la connessione corrente al database DMT
        Set .Database = TheApp.Database  '<-- Notare la connessione DMTDBLib.Database
        'Imposta l'Handle del form che contiene il controllo
        .HwndContainer = Me.hwnd
        'Imposta la descrizione dell'intestazione di colonna per il campo codice
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .CodeCaption4Find = "Cognome / Ragione sociale"
        'Imposta il campo da utilizzare come campo codice
        .CodeField = "Anagrafica"
        'indica se il codice è un campo numerico
        .CodeIsNumeric = False
        'Imposta la descrizione dell'intestazione di colonna per il campo descrizione
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .DescriptionCaption4Find = "Nome"
        'Imposta il campo da utilizzare come campo descrizione
        .DescriptionField = "Nome"
        'Indica il campo chiave univoco di accesso al record
        .KeyField = "IDAnagrafica"
        'Indica il nome della tabella o View da utilizzare per il reperimento dei dati
        .TableName = "IERepCliente"
        'Indica eventuali filtri fissi da utilizzare per l'estrazione dei record
        .Filter = "IDAzienda = " & TheApp.IDFirm
        'Abilita la voce del menu popup per l'esegui gestione
        .MenuFunctions("EseguiGestione").Enabled = True
        'Indica la funzione DMT da eseguire quando viene lanciata l'esegui gestione
        .IDExecuteFunction = 29 'Anagrafica
    End With

End Sub
