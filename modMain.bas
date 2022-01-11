Attribute VB_Name = "modMain"
'@Folder("Project")
Option Explicit

Private Const SOURCE_FILENAME As String = "C:\Users\User\Documents\excel-table-to-word\Source.xlsx"
Private Const DESTINATION_FILENAME As String = "C:\Users\User\Documents\excel-table-to-word\Destination.docx"
Private Const LIST_LEVEL As Integer = 3
Private Const COLUMN_OFFSET As Integer = 1
Private Const COLUMN_COUNT As Integer = 5 ' Number of columns to import into Word
Private Const COLUMN_GROUPNAME As Integer = 1 ' Index of the Excel column with groupName

Private SourceWorkbook As Excel.Workbook
Private DestinationDocument As Document
Private ExcelTableData As Variant ' Excel Range.Value2 as Variant

Public Sub Main()
    Set SourceWorkbook = GetSourceExcelSourceWorkbook
    ExcelTableData = GetTableFromExcel(SourceWorkbook)
    
    Set DestinationDocument = GetDestinationWordDocument
    LoopThroughTables DestinationDocument
End Sub

Public Sub LoopThroughTables(ByRef doc As Document)
    Dim par As Paragraph
    Dim tbl As Table
    Dim groupName As String
    
    For Each par In doc.ListParagraphs
        If par.Range.ListFormat.ListLevelNumber = LIST_LEVEL Then
            If (Not par.Next Is Nothing) Then
                If (par.Next.Range.Tables.Count > 0) Then
                    groupName = Replace(par.Range.Text, vbCr, vbNullString)
                    Set tbl = par.Next.Range.Tables(1)
                    ClearTable tbl
                    AddRows groupName, tbl
                End If
            End If
        End If
    Next par
End Sub

Private Sub AddRows(ByVal groupName As String, ByRef tbl As Table)
    Dim r As Row
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To UBound(ExcelTableData, 1)
        If StrComp(ExcelTableData(i, COLUMN_GROUPNAME), groupName, vbTextCompare) = 0 Then
            Set r = tbl.Rows.Add
            r.HeightRule = wdRowHeightAuto
            r.HeadingFormat = False
            For j = 1 To COLUMN_COUNT
                r.Cells(j).Range.Text = ExcelTableData(i, COLUMN_OFFSET + j)
            Next j
        End If
    Next i
End Sub

Private Sub ClearTable(ByRef tbl As Table)
    Dim i As Integer
    
    For i = tbl.Rows.Count To 2 Step -1
        tbl.Rows(i).Delete
    Next i
End Sub

Private Function GetTableFromExcel(ByRef exWb As Excel.Workbook) As Variant
    Dim lo As Excel.ListObject
    
    Set lo = exWb.Worksheets(1).ListObjects(1)
    GetTableFromExcel = lo.DataBodyRange.Value2
End Function

Private Function GetSourceExcelSourceWorkbook() As Excel.Workbook
    Dim objExcel As Object
    
    Set objExcel = New Excel.Application
    Set GetSourceExcelSourceWorkbook = objExcel.Workbooks.Open(SOURCE_FILENAME)
End Function

Private Function GetDestinationWordDocument() As Document
    Set GetDestinationWordDocument = Application.Documents(DESTINATION_FILENAME)
End Function
