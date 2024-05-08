VERSION 5.00
Begin VB.Form frmCausaliXML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIPORTA IN XML"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCausaliXML.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check11 
      Caption         =   "Non riportare riferimento Vs ordine cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   4695
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   10
      Top             =   5640
      Width           =   3375
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Targa automezzo"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   4695
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Agenzia trasporto"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   4695
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Vettore successivo"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Istruzioni del mittente"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Annotazione generale del documento"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   4695
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Annotazione 3 del documento"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Annotazione 2 del documento"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Annotazione 1 del documento"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Note I.V.A."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Lettera d'intento"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmCausaliXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
    Rip_InXMLRifLetteraIntento = Me.Check1.Value
    Rip_InXMLRifNoteIva = Me.Check2.Value
    Rip_InXMLRifNota01Doc = Me.Check3.Value
    Rip_InXMLRifNota02Doc = Me.Check4.Value
    Rip_InXMLRifNota03Doc = Me.Check5.Value
    Rip_InXMLRifNotaDoc = Me.Check6.Value
    Rip_InXMLRifIstrMitt = Me.Check7.Value
    Rip_InXMLRifVettSucc = Me.Check8.Value
    Rip_InXMLRifAgenziaTrasp = Me.Check9.Value
    Rip_InXMLRifTargaAutoMezzo = Me.Check10.Value
    NonRiportaInXMLRifVsNumOrd = Me.Check11.Value
    
    CONFERMA_PARAMETRI_XML = True
    Unload Me
Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    CONFERMA_PARAMETRI_XML = False
    Me.Check1.Value = Rip_InXMLRifLetteraIntento
    Me.Check2.Value = Rip_InXMLRifNoteIva
    Me.Check3.Value = Rip_InXMLRifNota01Doc
    Me.Check4.Value = Rip_InXMLRifNota02Doc
    Me.Check5.Value = Rip_InXMLRifNota03Doc
    Me.Check6.Value = Rip_InXMLRifNotaDoc
    Me.Check7.Value = Rip_InXMLRifIstrMitt
    Me.Check8.Value = Rip_InXMLRifVettSucc
    Me.Check9.Value = Rip_InXMLRifAgenziaTrasp
    Me.Check10.Value = Rip_InXMLRifTargaAutoMezzo
    Me.Check11.Value = NonRiportaInXMLRifVsNumOrd
    
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
