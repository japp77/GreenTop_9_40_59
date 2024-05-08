VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocChangeNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Questa classe viene utilizzata per ottenere la notifica sul cambiamento
'dei valori dei singoli campi della struttura tables dell'oggetto cDocument
'Per effettuare questo meccanismo bisogna implementare l'interfaccia
'IChangeNotify esposta dalla DmtDocs.
Implements DmtDocs.IChangeNotify

'Dichiariamo un evento per gestirlo all'interno dell'applicazione
Public Event ChangeValue(ByVal Table As DmtTables.cTable, ByVal Field As DmtTables.cField, ByVal Value As Variant)

'Questo metodo viene scatenato ogni volta che un valore viene modificato
Private Sub IChangeNotify_ChangeValue(ByVal Table As DmtTables.cTable, ByVal Field As DmtTables.cField, ByVal Value As Variant)
    
    If bSaving = 0 Then
        RaiseEvent ChangeValue(Table, Field, Value)
    End If
End Sub
