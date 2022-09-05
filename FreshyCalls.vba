' Proof of Concept: retrieving SSN for syscalling in VBA
' Author: Juan Manuel Fernandez (@TheXC3LL)


'Based on:
'https://www.mdsec.co.uk/2020/12/bypassing-user-mode-hooks-and-direct-invocation-of-system-calls-for-red-teams/
'https://www.crummie5.club/freshycalls/


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    Reserved0                       As Long
    PEBBaseAddress                  As LongPtr
    AffinityMask                    As LARGE_INTEGER
    BasePriority                    As Long
    Reserved1                       As Long
    uUniqueProcessId                As LARGE_INTEGER
    uInheritedFromUniqueProcessId   As LARGE_INTEGER
End Type
Private Type PEB
    Reserved1(1) As Byte
    BeingDebugged As Byte
    Reserved2(20) As Byte
    Ldr As LongPtr
    ProcessParameters As LongPtr
    Reserved3(519) As Byte
    PostProcessInitRoutine As Long
    Reserved4(135) As Byte
    SessionId As Long
End Type
Private Type IMAGE_DOS_HEADER
     e_magic As Integer
     e_cblp As Integer
     e_cp As Integer
     e_crlc As Integer
     e_cparhdr As Integer
     e_minalloc As Integer
     e_maxalloc As Integer
     e_ss As Integer
     e_sp As Integer
     e_csum As Integer
     e_ip As Integer
     e_cs As Integer
     e_lfarlc As Integer
     e_ovno As Integer
     e_res(4 - 1) As Integer
     e_oemid As Integer
     e_oeminfo As Integer
     e_res2(10 - 1) As Integer
     e_lfanew As Long
End Type
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Private Type IMAGE_OPTIONAL_HEADER
        Magic As Integer
        MajorLinkerVersion As Byte
        MinorLinkerVersion As Byte
        SizeOfCode As Long
        SizeOfInitializedData As Long
        SizeOfUninitializedData As Long
        AddressOfEntryPoint As Long
        BaseOfCode As Long
        ImageBase As LongLong
        SectionAlignment As Long
        FileAlignment As Long
        MajorOperatingSystemVersion As Integer
        MinorOperatingSystemVersion As Integer
        MajorImageVersion As Integer
        MinorImageVersion As Integer
        MajorSubsystemVersion As Integer
        MinorSubsystemVersion As Integer
        Win32VersionValue As Long
        SizeOfImage As Long
        SizeOfHeaders As Long
        CheckSum As Long
        Subsystem As Integer
        DllCharacteristics As Integer
        SizeOfStackReserve As LongLong
        SizeOfStackCommit As LongLong
        SizeOfHeapReserve As LongLong
        SizeOfHeapCommit As LongLong
        LoaderFlags As Long
        NumberOfRvaAndSizes As Long
        DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type
Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type
Private Type IMAGE_NT_HEADERS
    Signature As Long                         'DWORD Signature;
    FileHeader As IMAGE_FILE_HEADER           'IMAGE_FILE_HEADER FileHeader;
    OptionalHeader As IMAGE_OPTIONAL_HEADER   'IMAGE_OPTIONAL_HEADER OptionalHeader;
End Type




Private Declare PtrSafe Function NtQueryInformationProcess Lib "NTDLL" ( _
                        ByVal hProcess As LongPtr, _
                        ByVal processInformationClass As Long, _
                        ByRef pProcessInformation As Any, _
                        ByVal uProcessInformationLength As Long, _
                        ByRef puReturnLength As LongPtr) As Long
                         
Private Declare PtrSafe Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
                        ByVal Destination As LongPtr, _
                        ByVal Source As LongPtr, _
                        ByVal Length As Long)
                        
Private Declare PtrSafe Function lstrlenW Lib "KERNEL32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function lstrlenA Lib "KERNEL32" (ByVal lpString As LongPtr) As Long
Private Function StringFromPointerW(ByVal pointerToString As LongPtr) As String
    Const BYTES_PER_CHAR As Integer = 2
    Dim tmpBuffer()    As Byte
    Dim byteCount      As Long
    ' determine size of source string in bytes
    byteCount = lstrlenW(pointerToString) * BYTES_PER_CHAR
    If byteCount > 0 Then
        'Resize the buffer as required
        ReDim tmpBuffer(0 To byteCount - 1) As Byte
        ' Copy the bytes from pointerToString to tmpBuffer
        Call CopyMemory(VarPtr(tmpBuffer(0)), pointerToString, byteCount)
    End If
    'Straigth assigment Byte() to String possible - Both are Unicode!
    StringFromPointerW = tmpBuffer
End Function
Public Function StringFromPointerA(ByVal pointerToString As LongPtr) As String

    Dim tmpBuffer()    As Byte
    Dim byteCount      As Long
    Dim retVal         As String

    ' determine size of source string in bytes
    byteCount = lstrlenA(pointerToString)

    If byteCount > 0 Then
        ' Resize the buffer as required
        ReDim tmpBuffer(0 To byteCount - 1) As Byte

        ' Copy the bytes from pointerToString to tmpBuffer
        Call CopyMemory(VarPtr(tmpBuffer(0)), pointerToString, byteCount)
    End If

    ' Convert (ANSI) buffer to VBA string
    retVal = StrConv(tmpBuffer, vbUnicode)

    StringFromPointerA = retVal

End Function
Private Function LdrAddress()
    Dim ret As Long
    Dim size As LongPtr
    Dim pbi As PROCESS_BASIC_INFORMATION
    
    'Get PROCESS_BASIC_INFORMATION
    ret = NtQueryInformationProcess(-1, 0, pbi, LenB(pbi), size)
    
    'Copy PEB to a buffer
    Dim cPEB As PEB
    Call CopyMemory(VarPtr(cPEB), pbi.PEBBaseAddress, LenB(cPEB))
    
    'Return PPEB_LDR_DATA
    LdrAddress = cPEB.Ldr
End Function

Private Function FindNtdll()
    'https://www.vergiliusproject.com/kernels/x64/Windows%2010%20|%202016/2110%2021H2%20(November%202021%20Update)/_PEB_LDR_DATA
    'struct _LIST_ENTRY InLoadOrderModuleList;                               //0x10
    Dim InLoadOrderModuleList As LongPtr
    Dim currentEntry As LongPtr
    Dim nextEntry As LongPtr
    Dim dllbase As LongPtr
    Dim DllNamePtr As LongPtr
    Dim DllName As String
    Dim currentDllBase As LongPtr
    Dim Ldr As LongPtr
    Dim row As Integer
    'Ldr Address
    Ldr = LdrAddress
    
    'First entry
    Call CopyMemory(VarPtr(InLoadOrderModuleList), LdrAddress + &H18, LenB(InLoadOrderModuleList))
    Call CopyMemory(VarPtr(dllbase), InLoadOrderModuleList + &H30, LenB(dllbase))
    
    'Walk the list
    currentEntry = InLoadOrderModuleList
    Do Until nextEntry = InLoadOrderModuleList
        Call CopyMemory(VarPtr(nextEntry), currentEntry, LenB(nextEntry))
        Call CopyMemory(VarPtr(dllbase), currentEntry + &H30, LenB(dllbase))
        Call CopyMemory(VarPtr(DllNamePtr), currentEntry + &H58 + 8, LenB(DllNamePtr)) 'UNICODE_STRING USHORT + USHORT = 8
        DllName = StringFromPointerW(DllNamePtr)
        ' This should be done using a hash, but it's just a PoC
        If StrComp("ntdll.dll", DllName, 0) = 0 Then
            Exit Do
        End If
        currentEntry = nextEntry
    Loop
    FindNtdll = dllbase
End Function

Sub FreshyCalls()
    Dim dllbase As LongPtr
    Dim DosHeader As IMAGE_DOS_HEADER
    Dim pNtHeaders As LongPtr
    Dim ntHeader As IMAGE_NT_HEADERS
    Dim DataDirectory As IMAGE_DATA_DIRECTORY
    Dim IMAGE_EXPORT_DIRECTORY As LongPtr 'http://pinvoke.net/default.aspx/Structures.IMAGE_EXPORT_DIRECTORY
    Dim NumberOfFunctions As Long
    Dim NumberOfNames As Long
    Dim FunctionsPtr As LongPtr
    Dim NamesPtr As LongPtr
    Dim OrdinalsPtr As LongPtr
    Dim FunctionsOffset As Long
    Dim NamesOffset As Long
    Dim OrdinalsOffset As Long
    Dim OrdinalBase As Long
    
    ' Get ntdll.dll base
    dllbase = FindNtdll
    ' Get DOS Header
    Call CopyMemory(VarPtr(DosHeader), dllbase, LenB(DosHeader))
    ' Get NtHeader
    pNtHeaders = dllbase + DosHeader.e_lfanew
    Call CopyMemory(VarPtr(ntHeader), pNtHeaders, LenB(ntHeader))
    
    IMAGE_EXPORT_DIRECTORY = ntHeader.OptionalHeader.DataDirectory(0).VirtualAddress + dllbase
    
    'Number of Functions pIMAGE_EXPORT_DIRECTORY + 0x14
    Call CopyMemory(VarPtr(NumberOfFunctions), IMAGE_EXPORT_DIRECTORY + &H14, LenB(NumberOfFunctions))
    
    'Number of Names pIMAGE_EXPORT_DIRECTORY + 0x18
    Call CopyMemory(VarPtr(NumberOfNames), IMAGE_EXPORT_DIRECTORY + &H18, LenB(NumberOfNames))
    
    'AddressOfFunctions pIMAGE_EXPORT_DIRECTORY + 0x1C
    Call CopyMemory(VarPtr(FunctionsOffset), IMAGE_EXPORT_DIRECTORY + &H1C, LenB(FunctionsOffset))
    FunctionsPtr = dllbase + FunctionsOffset

    'AddressOfNames pIMAGE_EXPORT_DIRECTORY + 0x20
    Call CopyMemory(VarPtr(NamesOffset), IMAGE_EXPORT_DIRECTORY + &H20, LenB(NamesOffset))
    NamesPtr = dllbase + NamesOffset
    
    'AddressOfNameOrdianls pIMAGE_EXPORT_DIRECTORY + 0x24
    Call CopyMemory(VarPtr(OrdinalsOffset), IMAGE_EXPORT_DIRECTORY + &H24, LenB(OrdinalsOffset))
    OrdinalsPtr = dllbase + OrdinalsOffset
    
    'Ordinal Base pIMAGE_EXPORT_DIRECTORY + 0x10
    Call CopyMemory(VarPtr(OrdinalBase), IMAGE_EXPORT_DIRECTORY + &H10, LenB(OrdinalBase))
    
    Dim j As Long
    Dim i As Long
    j = 0
    For i = 0 To NumberOfNames - 1
        Dim tmpOffset As Long
        Dim tmpName As String
        Dim tmpOrd As Integer
        ' Get name
        Call CopyMemory(VarPtr(tmpOffset), NamesPtr + (LenB(tmpOffset) * i), LenB(tmpOffset))
        tmpName = StringFromPointerA(tmpOffset + dllbase)
        If InStr(1, tmpName, "Zw") = 1 Then
            Cells(j + 1, 1) = Replace(tmpName, "Zw", "Nt")
        'Get Ordinal
            Call CopyMemory(VarPtr(tmpOrd), OrdinalsPtr + (LenB(tmpOrd) * i), LenB(tmpOrd))
            Cells(j + 1, 2) = tmpOrd + OrdinalBase
        'Get Address
            tmpOffset = 0
            Call CopyMemory(VarPtr(tmpOffset), FunctionsPtr + (LenB(tmpOffset) * tmpOrd), LenB(tmpOffset))
            Cells(j + 1, 3) = tmpOffset
            j = j + 1
        End If
    Next i
    'Sort by Address
    Range("A1:C" & j).Sort , Key1:=Range("C1"), Order1:=xlAscending
    'Set number
    For k = 0 To j - 1
        Cells(k + 1, 2) = k
        Cells(k + 1, 3) = ""
    Next k
End Sub
