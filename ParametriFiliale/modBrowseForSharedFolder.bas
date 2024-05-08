Attribute VB_Name = "modBrowseForSharedFolder"
Option Explicit

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_STATUSTEXT = &H4
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_COMPUTERSNEARME As Long = &H3D
Private Const CSIDL_DRIVES = &H11
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Declare Function WNetGetConnection32 Lib "MPR.DLL" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, lSize As Long) As Long

Private lpszRemoteName As String
Private lSize As Long

Private Const NO_ERROR As Long = 0

Private Const lBUFFER_SIZE As Long = 255

Function GetNetPath(Lettera As String) As String
Dim cbRemoteName
Dim lStatus&
   ' Prompt the user to type the mapped drive letter.
   'DriveLetter = Lettera

   ' Add a colon to the drive letter entered.
   'DriveLetter = DriveLetter & ":"

   ' Specifies the size in characters of the buffer.
   cbRemoteName = lBUFFER_SIZE

   ' Prepare a string variable by padding spaces.
   lpszRemoteName = lpszRemoteName & Space(lBUFFER_SIZE)

   ' Return the UNC path (\\Server\Share).
   lStatus& = WNetGetConnection32(Lettera, lpszRemoteName, cbRemoteName)

   ' Verify that the WNetGetConnection() succeeded. WNetGetConnection()
   ' returns 0 (NO_ERROR) if it successfully retrieves the UNC path.
   If lStatus& = NO_ERROR Then

      ' Display the UNC path.
        GetNetPath = lpszRemoteName ', vbInformation

   Else
      ' Unable to obtain the UNC path.
        GetNetPath = Lettera
        
      'MsgBox "Unable to obtain the UNC path.", vbInformation
   End If

End Function
Public Function SfogliaCondivisione(ByVal hWnd As Long, ByRef strCondivisione As String, TipoRisorsa As Long, Optional ByVal strErrorStatus As String = "") As Boolean
Dim BInfo As BROWSEINFO
Dim lngPercorsoAllocato As Long
Dim strPercorsoScelto As String
Dim ARRAY_DRIVES() As String
Dim PercorsoUNC As String
Dim Percorso As String
Dim i As Long

    With BInfo
        .hOwner = hWnd
        .lpfn = FunPtr(AddressOf BrowseCallbackProc)
        .lpszTitle = "Scegli una condivisione:"
        If TipoRisorsa = 1 Then
            .pidlRoot = CSIDL_NETWORK
        End If
        If TipoRisorsa = 2 Then
            .pidlRoot = CSIDL_DRIVES
        End If
        .ulFlags = BIF_RETURNONLYFSDIRS
        If strErrorStatus <> "" Then
            .ulFlags = .ulFlags + BIF_STATUSTEXT
            strErrorStatus = StrConv(strErrorStatus, vbFromUnicode)
            .lParam = StrPtr(strErrorStatus)
        End If
    End With
    
    lngPercorsoAllocato = SHBrowseForFolder(BInfo)
    
    strPercorsoScelto = Space$(MAX_PATH)
    
    SfogliaCondivisione = SHGetPathFromIDList(lngPercorsoAllocato, strPercorsoScelto)
    
    If SfogliaCondivisione = True Then
        CoTaskMemFree lngPercorsoAllocato
        strCondivisione = Left$(strPercorsoScelto, InStr(strPercorsoScelto, Chr$(0)) - 1)
        
        If TipoRisorsa = 2 Then
            PercorsoUNC = ""
            ARRAY_DRIVES = Split(strCondivisione, "\")
            PercorsoUNC = GetNetPath(ARRAY_DRIVES(0))
            Percorso = ""
            For i = LBound(ARRAY_DRIVES) + 1 To UBound(ARRAY_DRIVES)
                Percorso = Percorso & "\" & ARRAY_DRIVES(i)
            Next
            
            strCondivisione = Trim(PercorsoUNC) & Percorso
        End If
    End If
    
    
End Function

Private Function FunPtr(ByVal lngFn As Long) As Long
    FunPtr = lngFn
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim strBuf As String
    Dim intPos As Integer
    
    If uMsg = BFFM_SELCHANGED Then
        strBuf = Space$(MAX_PATH)
        If SHGetPathFromIDList(lParam, strBuf) Then
            intPos = InStr(3, strBuf, "\")
            If InStr(intPos + 1, strBuf, "\") > 0 Then
                'SendMessage hWnd, BFFM_ENABLEOK, ByVal 0&, ByVal 0&
                'SendMessage hWnd, BFFM_SETSTATUSTEXTA, ByVal 0&, ByVal lpData
                SendMessage hWnd, BFFM_SETSTATUSTEXTA, ByVal 0&, ByVal vbNullChar
            Else
                SendMessage hWnd, BFFM_SETSTATUSTEXTA, ByVal 0&, ByVal vbNullChar
            End If
        Else
            SendMessage hWnd, BFFM_SETSTATUSTEXTA, ByVal 0&, ByVal lpData
            SendMessage hWnd, BFFM_ENABLEOK, ByVal 0&, ByVal 0&
        End If
    End If
End Function
