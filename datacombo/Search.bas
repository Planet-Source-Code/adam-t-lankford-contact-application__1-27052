Attribute VB_Name = "Search"
'********************************************************************
'This module contains a function that will allow record
'set search procedures to take place.
'~Created by Adam Lankford 4/18/2001
' Company:       RMIC Corp.  (Policy Servicing Department)
'-----------------------------------------------------------------
'********************************************************************

Public Function Search(parameter As String, rs As adodb.Recordset, x As field) As Boolean
    Dim foundFlag As Boolean
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
                For i = 1 To .RecordCount
                    If x = parameter Then
                        foundFlag = True
                        i = .RecordCount
                    End If
                    If foundFlag = False Then
                        .MoveNext
                    End If
                Next i
                If foundFlag = True Then
                   'MsgBox ("Record has been location!")
                Else
                    MsgBox ("No Match in Database!")
                    foundFlag = False
                    .MoveFirst
                End If
        Else
            MsgBox ("There are no records to search!")
            foundFlag = False
        End If
    End With
    Search = foundFlag
End Function


