Attribute VB_Name = "Module1"
Option Compare Database

Sub Compare()
    Dim db As Database
    Dim rs As Recordset
    Dim strSQL As String
    Dim FieldName As String
    Dim FieldValue As String
    Dim recOut As Recordset
    Dim delDataSQL As String
    Dim f2Index As Integer 'calculate index of matching column

    
    'SQL Queries
    strSQL = "SELECT [Combined SCAF].*, [Snapshot Combined SCAF].* FROM [Combined SCAF], [Snapshot Combined SCAF] Where [Combined SCAF].[Customer ID]=[Snapshot Combined SCAF].[Customer ID]"
    delDataSQL = "Delete * FROM [Snapshot Combined SCAF]"
    appendDataSQL = "INSERT Into [Snapshot Combined SCAF] Select * From [Combined SCAF]"
    
    Set db = CurrentDb()
    Set recOut = db.OpenRecordset("SCAF Change Logs", dbOpenDynaset, dbEditAdd)
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    Dim j As Integer
    j = 0
    With rs
        .MoveFirst
        Do
            For col = 1 To .Fields.Count \ 2 - 1
            f2Index = (.Fields.Count \ 2) + col + 1
                If .Fields(col) <> .Fields(f2Index) And col <> 7 Then
                    recOut.AddNew
                        recOut.Fields("Project") = .Fields("Combined SCAF.Project")
                        recOut.Fields("Cluster/Area") = .Fields("Snapshot Combined SCAF.Cluster/Area")
                        recOut.Fields("Crown Node ID") = .Fields("Snapshot Combined SCAF.Crown Node ID")
                        recOut.Fields("Crown Node SCU") = .Fields("Snapshot Combined SCAF.Crown Node SCU")
                        recOut.Fields("Lat") = .Fields("Snapshot Combined SCAF.Latitude")
                        recOut.Fields("Lon") = .Fields("Snapshot Combined SCAF.Longitude")
                        recOut.Fields("Column Edited") = .Fields(col).Name
                        recOut.Fields("Original value") = .Fields(f2Index)
                        recOut.Fields("New Value") = .Fields(col)
                        recOut.Fields("Date Changed") = Date
                    recOut.Update
                End If
            Next col
            
            .MoveNext
        Loop Until .EOF
    End With
    db.Execute (delDataSQL)
    db.Execute (appendDataSQL)

    Set rs = Nothing
    Set db = Nothing
End Sub


