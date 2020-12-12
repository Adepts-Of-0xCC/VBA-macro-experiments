Private Declare PtrSafe Function GetModuleHandleA Lib "KERNEL32" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "KERNEL32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

'VBA Macro that detects hooks made by EDRs
'PoC By Juan Manuel Fernandez (@TheXC3LL) based on a post from SpecterOps (https://posts.specterops.io/adventures-in-dynamic-evasion-1fe0bac57aa)


Public Function checkHook(ByVal target As String, hModule As LongPtr) As Integer
    Dim address As LongPtr
    Dim safeCheck As LongLong
    Dim tmpCheck As LongLong
    
    'Opcodes turned into numeric value (mov r10, rcx; mov eax, ??)
    safeCheck = 3100740428#

    address = GetProcAddress(hModule, target)
    Call CopyMemory(VarPtr(tmpCheck), address, 4)
    If tmpCheck <> safeCheck Then
            checkHook = 1
        Else
            checkHook = 0
    End If
End Function


Sub hookdetector()
    Dim functionList() As String
    Dim element As Variant
    Dim hModule As LongPtr
    Dim result As Integer
    Dim row As Integer
    
    ' Set as needed, this is just a PoC :)
    functionList = Split("NtAllocateVirtualMemory,NtAllocateVirtualMemoryEx,NtCreateThread,NtCreateThreadEx,NtCreateUserProcess,NtFreeVirtualMemory,NtLoadDriver,NtMapViewOfSection,NtOpenProcess,NtProtectVirtualMemory,NtQueueApcThread,NtQueueApcThreadEx,NtResumeThread,NtSetContextThread,NtSetInformationProcess,NtSuspendThread,NtUnloadDriver,NtWriteVirtualMemory", ",")
    
    hModule = GetModuleHandleA("ntdll.dll")
    row = 1
    For Each element In functionList
        result = checkHook(element, hModule)
        Cells(row, 1) = element
        If result <> 0 Then
                Cells(row, 2) = "Hooked"
                Cells(row, 2).Interior.Color = RGB(255, 0, 0)
            Else
                Cells(row, 2) = "Clear"
                Cells(row, 2).Interior.Color = RGB(0, 255, 0)
        End If
        row = row + 1
    Next element
End Sub
