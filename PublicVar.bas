Attribute VB_Name = "PublicVar"
'The general variables
Public dataName, interfaceName, listName As String
Public states(5) As String

'Headers of the Data Sheet
Public ID_data, file_data, requestor_data, comm_data, state_data As Variant

'Headers of Interface Sheet
Public pend_int, prog_int, revi_int, sign_int, block_int As Variant

'Headers of List Sheet
Public state_list As Variant

Public Sub Variables()
    dataName = "Data"
    interfaceName = "Interface"
    listName = "List"
    
    'states
    states(0) = "Pending"
    states(1) = "In progress"
    states(2) = "In review"
    states(3) = "To be sign"
    states(4) = "Validated"
    states(5) = "Blocked"
    
End Sub

Public Sub data_headers()
    Dim data_sheet As Worksheet
    Call Variables
    Set data_sheet = ThisWorkbook.Worksheets(dataName)
    
    Set ID_data = data_sheet.Cells.Find("ID", LookAt:=xlWhole)
    Set file_data = data_sheet.Cells.Find("File", LookAt:=xlWhole)
    Set requestor_data = data_sheet.Cells.Find("Requestor", LookAt:=xlWhole)
    Set comm_data = data_sheet.Cells.Find("Comment", LookAt:=xlWhole)
    Set state_data = data_sheet.Cells.Find("State", LookAt:=xlWhole)
    
End Sub


Public Sub interface_headers()
    Dim int_sheet As Worksheet
    Call Variables
    Set int_sheet = ThisWorkbook.Worksheets(interfaceName)
    
    Set pend_int = int_sheet.Cells.Find("Pending", LookAt:=xlWhole)
    Set prog_int = int_sheet.Cells.Find("In progress", LookAt:=xlWhole)
    Set revi_int = int_sheet.Cells.Find("In review", LookAt:=xlWhole)
    Set sign_int = int_sheet.Cells.Find("To be sign", LookAt:=xlWhole)
    Set block_int = int_sheet.Cells.Find("Blocked", LookAt:=xlWhole)
    
End Sub


Public Sub list_headers()
    Dim list_sheet As Worksheet
    Call Variables
    Set list_sheet = ThisWorkbook.Worksheets(listName)
    
    Set state_list = list_sheet.Cells.Find("State", LookAt:=xlWhole)
    
End Sub
