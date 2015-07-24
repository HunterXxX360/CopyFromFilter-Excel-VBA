Option Explicit

Sub DoSplit()
Dim xlXML As Object
Dim GenRst As Object      'GeneralRecordset
Dim GenAr() As Variant    'GeneralArray
Dim UnAr As Variant       'UniqueArray
Dim GenArRows As Double   'Row-Count of GeneralArray
Dim UnArRows As Double    'Row-Count of UniqueArray
Dim GenArRow As Double    'Row-Value of GeneralArray
Dim UnArRow As Double     'Row-Value of UniqueArray
Dim FiltVal As Variant    'FilterValue
Dim WBName As String      'Workbook-Name


'   create recordset
Set GenRst = CreateObject("ADODB.Recordset")
Set xlXML = CreateObject("MSXML2.DOMDocument")
    
    '   define recordset from Functions.RangeContent()
    xlXML.LoadXML RangeContent.Value(xlRangeValueMSPersistXML)
    
    With GenRst
        .Open xlXML
        .MoveFirst
        
        '   ReDim GeneralArray on length of recordset
        GenArRows = .RecordCount
        ReDim GenAr(GenArRows) As Variant
        GenArRow = 0
        
        '   fill GeneralArray with FilterValue
        Do While Not .EOF
            GenAr(GenArRow) = .Fields(#).Value  '###use the desired field for filter###
            GenArRow = GenArRow + 1
            .MoveNext
        Loop
        
        '   create UniqueArray by identifying unique items with functions.UniqueItems()
        UnAr = UniqueItems(GenAr)
        UnArRows = UBound(UnAr)
        
        Application.ScreenUpdating = False
        On Error GoTo proceed:
        
        '   for every unique item ...
        For UnArRow = 0 To UnArRows
            VO = UnAr(UnArRow)
            WBName = VO & ".xls"
            '   ... create a workbook with MakeWB()
            MakeWB WBName
            
            With Workbooks(WBName)
                With .Worksheets("Tabelle1")
                    '   filter GeneralRecordset on the unique item
                    GenRst.Filter = "[] = '" & VO & "'"
                    '   copy the filtered GeneralRecordset into new workbook
                    .Range("A3").CopyFromRecordset GenRst
                    GenRst.Filter = 0
                End With
                .Save
                .Close
            End With
        Next UnArRow
proceed:
        Application.ScreenUpdating = True
    End With
    
Set xlXML = Nothing
Set GenRst = Nothing
    
End Sub

Private Sub MakeWB(WBName As String)
Dim WBPath As String
Dim WB As Workbook
    '   create a new workbook using WBName
    WBPath = ThisWorkbook.Path & "\" & WBName
    If Dir(WBPath) <> "" Then
        Kill WBPath
    End If
    Set WB = Workbooks.Add
    WB.SaveAs Filename:=WBPath, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Set WB = Nothing
End Sub
