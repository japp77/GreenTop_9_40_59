Attribute VB_Name = "ModFF_EAN13"
Public X As String

Function FF_EAN13Digit(CodeString As String)

Dim V1(9, 2) As String
V1(0, 0) = "a"
V1(0, 1) = "b"
V1(0, 2) = "c"
V1(1, 0) = "d"
V1(1, 1) = "e"
V1(1, 2) = "f"
V1(2, 0) = "g"
V1(2, 1) = "h"
V1(2, 2) = "i"
V1(3, 0) = "j"
V1(3, 1) = "k"
V1(3, 2) = "l"
V1(4, 0) = "m"
V1(4, 1) = "n"
V1(4, 2) = "o"
V1(5, 0) = "p"
V1(5, 1) = "q"
V1(5, 2) = "r"
V1(6, 0) = "s"
V1(6, 1) = "t"
V1(6, 2) = "u"
V1(7, 0) = "v"
V1(7, 1) = "w"
V1(7, 2) = "x"
V1(8, 0) = "y"
V1(8, 1) = "z"
V1(8, 2) = "A"
V1(9, 0) = "B"
V1(9, 1) = "C"
V1(9, 2) = "D"

Dim V2(9) As String
V2(0) = "000000"
V2(1) = "001011"
V2(2) = "001101"
V2(3) = "001110"
V2(4) = "010011"
V2(5) = "011001"
V2(6) = "011100"
V2(7) = "010101"
V2(8) = "010110"
V2(9) = "011010"


Dim Risultato As String
Dim Codifica As Integer
Dim CheckDigit As Integer

X = Trim(CodeString)
If Not IsNumeric(X) Then 'Or Len(CodeString) < 12 Then
    FF_EAN13Digit = ""
    Exit Function
End If
X = Left(CodeString, 13)

'Aggiunta del check-digit
CheckDigit = 0
For I = 1 To 11 Step 2
    CheckDigit = CheckDigit + Val(Mid(X, I, 1))
    CheckDigit = CheckDigit + Val(Mid(X, I + 1, 1)) * 3
Next I
CheckDigit = (10 - CheckDigit Mod 10) Mod 10
'X = X '& Trim(Str(CheckDigit))

'Trasformazione del 13. carattere (codificato come start/stop)
'Codifica = Val(Left(X, 1))
'Risultato = Left(X, 1)

'Trasformazione dei caratteri da 12 a 7
'For I = 2 To 7
'    Risultato = Risultato & V1(Val(Mid(X, I, 1)), Val(Mid(V2(Codifica), I - 1, 1)))
'Next I

'Aggiunta del carattere di controllo centrale
'Risultato = Risultato & "G"

'Trasformazione dei caratteri da 6 a 1
'For I = 8 To 13
'    Risultato = Risultato & V1(Val(Mid(X, I, 1)), 2)
'Next I

'Aggiunta del carattere di start/stop finale
'Risultato = Risultato & "F"

FF_EAN13Digit = CheckDigit
'FF_EAN13 = X & Trim(Str(CheckDigit))
End Function
Function FF_EAN13(CodeString As String)

Dim V1(9, 2) As String
V1(0, 0) = "a"
V1(0, 1) = "b"
V1(0, 2) = "c"
V1(1, 0) = "d"
V1(1, 1) = "e"
V1(1, 2) = "f"
V1(2, 0) = "g"
V1(2, 1) = "h"
V1(2, 2) = "i"
V1(3, 0) = "j"
V1(3, 1) = "k"
V1(3, 2) = "l"
V1(4, 0) = "m"
V1(4, 1) = "n"
V1(4, 2) = "o"
V1(5, 0) = "p"
V1(5, 1) = "q"
V1(5, 2) = "r"
V1(6, 0) = "s"
V1(6, 1) = "t"
V1(6, 2) = "u"
V1(7, 0) = "v"
V1(7, 1) = "w"
V1(7, 2) = "x"
V1(8, 0) = "y"
V1(8, 1) = "z"
V1(8, 2) = "A"
V1(9, 0) = "B"
V1(9, 1) = "C"
V1(9, 2) = "D"

Dim V2(9) As String
V2(0) = "000000"
V2(1) = "001011"
V2(2) = "001101"
V2(3) = "001110"
V2(4) = "010011"
V2(5) = "011001"
V2(6) = "011100"
V2(7) = "010101"
V2(8) = "010110"
V2(9) = "011010"


Dim Risultato As String
Dim Codifica As Integer
Dim CheckDigit As Integer

X = Trim(CodeString)
If Not IsNumeric(X) Then 'Or Len(CodeString) < 12 Then
    FF_EAN13 = ""
    Exit Function
End If
X = Left(CodeString, 13)

'Aggiunta del check-digit
'CheckDigit = 0
'For I = 1 To 11 Step 2
'    CheckDigit = CheckDigit + Val(Mid(X, I, 1))
'    CheckDigit = CheckDigit + Val(Mid(X, I + 1, 1)) * 3
'Next I
'CheckDigit = (10 - CheckDigit Mod 10) Mod 10
X = X '& Trim(Str(CheckDigit))

'Trasformazione del 13. carattere (codificato come start/stop)
Codifica = Val(Left(X, 1))
Risultato = Left(X, 1)

'Trasformazione dei caratteri da 12 a 7
For I = 2 To 7
    Risultato = Risultato & V1(Val(Mid(X, I, 1)), Val(Mid(V2(Codifica), I - 1, 1)))
Next I

'Aggiunta del carattere di controllo centrale
Risultato = Risultato & "G"

'Trasformazione dei caratteri da 6 a 1
For I = 8 To 13
    Risultato = Risultato & V1(Val(Mid(X, I, 1)), 2)
Next I

'Aggiunta del carattere di start/stop finale
Risultato = Risultato & "F"

FF_EAN13 = Risultato
'FF_EAN13 = X & Trim(Str(CheckDigit))
End Function

