Attribute VB_Name = "InterfaceManagement"
Private Function collectData()
'This function collects all the data that is not in the Validated State,
'and returns and Array with all the info
    Dim data_sheet As Worksheet
    Dim dataArray() As Variant
    
    'Initalize the public Variables
    Call Variables
    Call data_headers
    
    'The shortcut
    Set data_sheet = ThisWorkbook.Worksheets(dataName)
    
    'The end of Data
    end_data = data_sheet.Cells(Rows.Count, ID_data.Column).End(xlUp).Row
    
    'Collect al data that hasn't been validated
    counter = 0
    For i = ID_data.Row + 1 To end_data
        ID = data_sheet.Cells(i, ID_data.Column).Value
        file = data_sheet.Cells(i, file_data.Column).Value
        requestor = data_sheet.Cells(i, requestor_data.Column).Value
        Comment = data_sheet.Cells(i, comm_data.Column).Value
        state = data_sheet.Cells(i, state_data.Column).Value
        If state <> "Validated" Then
            ReDim Preserve dataArray(4, counter)
            dataArray(0, counter) = ID
            dataArray(1, counter) = file
            dataArray(2, counter) = requestor
            dataArray(3, counter) = Comment
            dataArray(4, counter) = state
            counter = counter + 1
        End If
    Next i
    
    'Return the array
    collectData = dataArray

End Function

Private Function getLastColumn(headerRow)
'This function get the last column for a state of the interface Sheet.
    Dim inter_sheet As Worksheet
    
    'Initialize the variables
    Call Variables
    'The shortcut
    Set inter_sheet = ThisWorkbook.Worksheets(interfaceName)
    
    'We obtain the largest column of the section
    For i = headerRow + 1 To headerRow + 4
        endRow = inter_sheet.Cells(i, Columns.Count).End(xlToLeft).Column
        dev = 1
        If endRow > result Then
            result = endRow
        End If
    Next i
    
    getLastColumn = result
    
End Function

Private Sub cleanInterface()
'This sub will clean all the Post-Its from the interface
'color: 15652797
    Dim inter_sheet As Worksheet
    
    'Initialize the variables
    Call Variables
    bgColor = 15652797
    
    'The shortcut
    Set inter_sheet = ThisWorkbook.Worksheets(interfaceName)
        
    'we go through all the states cleaning the post it
    For k = 0 To UBound(states)
        Set Search = inter_sheet.Cells.Find(states(k), LookAt:=xlWhole)
        If Not (Search Is Nothing) Then
            headerRow = Search.Row
            headerCol = Search.Column
            lastCol = getLastColumn(headerRow)
    
            For i = headerCol To lastCol
                For j = headerRow + 1 To headerRow + 4
                    With inter_sheet.Cells(j, i)
                        .Value = Null
                        .Interior.Color = bgColor
                        .Font.Color = vbBlack
                        .VerticalAlignment = xlVAlignCenter
                        .HorizontalAlignment = xlLeft
                    End With
                    dev = 1
                Next j
            Next i
        End If
    Next k
    
End Sub

Private Sub CreatePostIt(state, ID, file, Comment, requestor)
'This sub create a Post-It in the line we give it
'color:10086143
    Dim inter_sheet As Worksheet
    
    'Initialize the variables
    Call Variables
    bgColor = 10086143
    
    'The shortcut
    Set inter_sheet = ThisWorkbook.Worksheets(interfaceName)
    
    'The position of the state
    Set Search = inter_sheet.Cells.Find(state, LookAt:=xlWhole)
    
    If Not (Search Is Nothing) Then
        'the position of the post-it
        lastCol = getLastColumn(Search.Row)
        If lastCol = 1 Then
            col = 2
        Else
            col = lastCol + 2
        End If
        
        'The format of the post-it and its content
        With inter_sheet.Cells(Search.Row + 1, col)
            .Value = file
            .Interior.Color = bgColor
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlLeft
        End With
        With inter_sheet.Cells(Search.Row + 2, col)
            .Value = ID
            .Interior.Color = bgColor
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlLeft
        End With
        With inter_sheet.Cells(Search.Row + 3, col)
            .Value = Comment
            .Font.Color = vbRed
            .Interior.Color = bgColor
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlVAlignCenter
        End With
        With inter_sheet.Cells(Search.Row + 4, col)
            .Value = requestor
            .Interior.Color = bgColor
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlLeft
        End With
    End If
    
    
End Sub

Sub renderInterface()
'this sub render the screen of the Sheet with al the post its
    
    Application.ScreenUpdating = False
    
    'Get all the data from Data Sheet
    allData = collectData()
    
    'Clean the interface
    Call cleanInterface
    
    'Add all the Post-Its
    For i = 0 To UBound(allData, 2)
        state = allData(4, i)
        ID = allData(0, i)
        file = allData(1, i)
        Comment = allData(3, i)
        requestor = allData(2, i)
        Call CreatePostIt(state, ID, file, Comment, requestor)
    Next i
    
    Application.ScreenUpdating = True
    
End Sub
