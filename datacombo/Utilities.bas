Attribute VB_Name = "Utilities"
' Function Used to open connection and establish a recordset
Public Function OpenADO(rs As ADODB.Recordset, conn As ADODB.Connection, sql As String)
    Dim strConnect As String
    Dim strProvider As String
    Dim strDataSource As String
    Dim strDatabaseName As String

    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDataSource = PATH & "\"
    strDatabaseName = "addresses.mdb"
    strDataSource = "Data Source=" & strDataSource & _
        strDatabaseName
    
    strConnect = strProvider & strDataSource



    conn.CursorLocation = adUseClient
    conn.Open strConnect


    rs.CursorType = adOpenStatic

    rs.LockType = adLockPessimistic

    rs.Source = sql
    
    rs.ActiveConnection = conn

    rs.Open
End Function


' Function for Deleting a record!
Public Function DeleteRecord(rs As ADODB.Recordset)
    Dim X As Integer
    
    If USER_TYPE = "administrator" Then
        X = MsgBox("Are you sure you want to delete this record?", vbYesNo)
        If X = vbYes Then
            rs.Delete
            MsgBox ("Record has been Deleted!")
            rs.MoveNext
        End If
    Else
        MsgBox "You don't have permission to delete records!", vbExclamation + vbOKOnly, "Record Deletion"
    End If
    

End Function

' Function Used to open Excel connection and establish a recordset
' Ok, I got tired of writing this code over & over just to open a recordset...
' This function is customize to open the CheckRequest database with a little tweaking
' it could be a universal function for any connection!
'**********************************************************************
' Function Revised by: Adam Lankford on July 11 2001.
' Company: RMIC Corp (Policy Servicing Department)
'**********************************************************************
Public Function OpenExcelADO(rs As ADODB.Recordset, conn As ADODB.Connection, strPath As String, strSheetname As String)
    Dim strConnect As String
    Dim strProvider As String
    Dim strAdditionalData As String
    Dim strDataSource As String


    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDataSource = "Data Source=" & strPath & ";"
    strAdditionalData = "Extended Properties=""Excel 8.0;HDR=YES;"""
    strConnect = strProvider & strDataSource & strAdditionalData

    conn.CursorLocation = adUseClient
    Debug.Print strConnect
    conn.Open strConnect
    

    rs.CursorType = adOpenStatic

    rs.LockType = adLockPessimistic


    rs.ActiveConnection = conn

    rs.Open "Select * from [" & strSheetname & "$]"""
End Function


