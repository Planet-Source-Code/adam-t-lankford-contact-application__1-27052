Attribute VB_Name = "Module1"
Option Explicit
' various methods used to populate data into combo boxes
' constructed by ~ATL~ [BoMbSnAtCh]
Public Function PopulateAM(cbo As ComboBox)
 Dim connAM As ADODB.Connection
    
    Dim rsAM As ADODB.Recordset
    Dim strSQL As String
    
    
    Set connAM = New ADODB.Connection
    Set rsAM = New ADODB.Recordset
    
    strSQL = "SELECT AccountManager FROM AccountManagers ORDER BY AccountManager"
    
    Call Utilities.OpenADO(rsAM, connAM, strSQL)
    
    ReDim AccountManagers(1 To rsAM.RecordCount)
    Dim i As Integer
    Dim AMcnt As Integer
    
    With rsAM
        .MoveFirst
        AMcnt = 1
        For i = 1 To .RecordCount
            AccountManagers(AMcnt) = !AccountManager
            AMcnt = AMcnt + 1
            .MoveNext
        Next i
    End With
    
    connAM.Close
    
    Set connAM = Nothing
    Set rsAM = Nothing
    
    For i = 1 To AMcnt - 1
        cbo.AddItem (AccountManagers(i))
    Next i
      
End Function

Public Function PopulateEmployees(cbo As ComboBox)
    Dim conn            As ADODB.Connection
    Dim rs              As ADODB.Recordset
    Dim i               As Long
    Dim sql             As String
    Dim strConnect      As String
    Dim strProvider     As String
    Dim strDataSource   As String

    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDataSource = "\\Poseidon\common\Employee Database\PS_Employee.mdb"
    strDataSource = "Data Source=" & strDataSource
    strConnect = strProvider & strDataSource
    sql = "SELECT name FROM employee ORDER BY name"
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    conn.CursorLocation = adUseClient
    conn.Open strConnect
    

    rs.CursorType = adOpenStatic

    rs.LockType = adLockPessimistic

    rs.Source = sql

    rs.ActiveConnection = conn

    rs.Open
     
    For i = 1 To rs.RecordCount
        cbo.AddItem (rs.Fields("name"))
        rs.MoveNext
    Next i
    
End Function
