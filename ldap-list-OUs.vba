Public Function toStr(pVar_In As Variant) As String
    On Error Resume Next
    toStr = CStr(pVar_In)
End Function

Sub LDAPSenum() 'https://www.remkoweijnen.nl/blog/2007/11/01/query-active-directory-from-excel/
    'Get the domain string ("dc=domain, dc=local")
    Dim strDomain As String
    strDomain = GetObject("LDAP://rootDSE").Get("defaultNamingContext")

    'ADODB Connection to AD
    Dim objConnection As Object
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"

    'Connection
    'Dim objCommand As ADODB.Command
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection

    'Search the AD recursively, starting at root of the domain
    objCommand.CommandText = _
        "<LDAP://" & strDomain & ">;(&(objectClass=organizationalUnit));,distinguishedName;subtree"
    'Dim objRecordSet As ADODB.Recordset
    Set objRecordSet = objCommand.Execute

    Dim i As Long

    If objRecordSet.EOF And objRecordSet.BOF Then
    Else
        Dim c As Integer
        c = 1
        Do While Not objRecordSet.EOF
            For i = 0 To objRecordSet.Fields.Count - 1
                Cells(c, 1) = toStr(objRecordSet!distinguishedName)
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
