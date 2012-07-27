Sub mySQl_Export()
'
' mySQl_Export Macro
' Exports sheet as mySQL for easy import
'
Dim query
Dim tmp As String
Dim columns() As String
Dim Output As String
Dim RowTmp As String
Dim RowArr() As String
Dim OutputFile As String

OutputFile = "R:\Clients A - F\Anglo American\03 people dev way\ANGA20-006 Revised PDW\02 Content\learningDirectory\ALAN\test.sql"


' Start of query
query = "INSERT INTO coursedata ("

' Get cell boundaries
LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
LastCol = ActiveSheet.UsedRange.columns(ActiveSheet.UsedRange.columns.Count).Column



ReDim Preserve columns(1 To LastCol)
ReDim Preserve RowArr(1 To LastCol)

'   Get cell headers
r = 1
For c = 1 To LastCol
    Cells(r, c).Select
    If ActiveCell <> "" Then
        columns(c) = "`" + ActiveCell + "`"
    End If
Next c
query = query + Join(columns, ",") + ") VALUES "
        
' Cycle all rows and get data
For r = 2 To LastRow
    RowTmp = "("
    ReDim RowArr(1 To LastCol)
    For c = 1 To LastCol
        Cells(r, c).Select
        If ActiveCell = "" Then
            RowArr(c) = "''"
        Else
            If IsNumeric(ActiveCell) Then
                RowArr(c) = ActiveCell
            Else
                tmp = Active
                RowArr(c) = "'" + Replace(ActiveCell, "'", "") + "'"
            End If
        End If
    Next c
    RowTmp = "(" + Join(RowArr, ",") + ");"
    Output = Output & query & RowTmp & vbNewLine
Next r

tmp = WriteToATextFile(Output, OutputFile)


End Sub

Function WriteToATextFile(Content As String, MyFile As String)

'set and open file for output
fnum = FreeFile()
Open MyFile For Output As fnum
'write project info and then a blank line. Note the comma is required
Print #fnum, Content
Close #fnum

MsgBox "SQL file written to " & MyFile

End Function
