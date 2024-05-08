VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaintNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DmtGridCtl.IPaintNotify

Private Sub IPaintNotify_PaintValue(ByVal ColumnsValue As DmtGridCtl.dgColumns, ByVal CurrentColumn As DmtGridCtl.dgColumnHeader, ByVal Row As Integer, Value As Variant)
On Error Resume Next
    Select Case CurrentColumn.FieldName
    
            
        Case "PesoNetto"
            Value = fnNotNullN(ColumnsValue("Art_peso").Value) - fnNotNullN(ColumnsValue("Art_tara").Value)
    End Select

End Sub

