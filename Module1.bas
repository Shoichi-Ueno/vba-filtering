Attribute VB_Name = "Module1"
Sub Filtering()
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim pass As String
    pass = ThisWorkbook.Path & "\sample"    '/Folder for data to be filtered.
    Dim mybook As Workbook
    Set mybook = Workbooks("list.xlsm")     '/File to list up.
    Dim F As File
    Dim maxRow As Long
    Dim maRow As Long
    Dim listmRow As Long
    Dim F As File
    For Each F In fso.GetFolder(pass).Files
        With Workbooks.Open(F)
        
            With .Worksheets(1)
            
                Worksheets(1).Copy Before: mybook.Worksheets ("LIST")                                                               '/Sheet of a sigle file in the folder, "\sample," is copied.
                
                maxRow = Cells(Rows.Count, 1).End(xlUp).Row                                                                         '/Count a number of rows and make variable.
                mybook.Worksheets(1).Range("E1:E" & maxRow).CurrentRegion.AdvancedFilter _
                Action:=xlFilterCopy, _
                CopyToRange:=Worksheets("Temporary List").Range("A1:E1"), _
                Unique:=True
                
                maRow = mybook.Sheets("Temporary List").Cells(Rows.Count, 1).End(xlUp).Row                                          '/ a number of filterd rows.
                
                listmRow = mybook.Sheets("LIST").Cells(Rows.Count, 1).End(xlUp).Row                                                 '/ a number ofadded rows in "LIST"'s sheet.
                
                mybook.Sheets("Temporary List").Range("A2:E" & maRow).Copy Destination:=Sheets("LIST").Range("A" & listmRow + 1)    '/Filterd data is temporary listed up in "Temporary List," and copy to "LIST"'s sheet.
                
                mybook.Sheets("Temporary List").Range("A2:E" & maRow).Clear                                                         '/After copied, the list in "Temporary list" is deleted.
                
                Application.DisplayAlerts = False
                mybook.Worksheets(1).Delete                                                                                         '/The copied sheet from the file in the folder is deleted.
                Application.DisplayAlerts = True
                
            End With
            .Close
        End With
    Next F
End Sub


