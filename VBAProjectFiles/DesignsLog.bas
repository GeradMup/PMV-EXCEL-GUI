Attribute VB_Name = "DesignsLog"
'Describes what happens when the Select Design button is pressed
Function selectDesign()
    'First let's check if the selected row is valid for designs
    If selectedRowValid() = False Then
        reportError "NO DESIGN SELECTED!"
        Exit Function
    End If
    
    'Now extract the file name from the selected row
    readSelectedFileName
    
    'Now open the file and read its contents
    readInputFile
    
End Function

'Validates if the user has selected a row which is within the required ranges or not
'Valid rows must be between row number 9 and the last row for design logs
'Returns True if selected row is valid, otherwise False
Function selectedRowValid() As Boolean
    Const MIN_ROW As Integer = 9
    Dim MAX_ROW As Integer
    Dim selectedRow As Integer
    
    MAX_ROW = Sheets("Designs Log").Range("AA3").Value
    selectedRow = ActiveCell.Row

    If selectedRow >= 9 And selectedRow < MAX_ROW Then
        selectedRowValid = True
    Else
        selectedRowValid = False
    End If
End Function

'Reads the selected file name from the selected row and saves it on the Designs Log page
Sub readSelectedFileName()
    Dim selectedRow As Integer
    Dim fileName As String
    
    selectedRow = ActiveCell.Row
    'The file name is found in column N in the log
    fileName = Sheets("Designs Log").Range("N" & selectedRow).Value
    
    fileName = Trim(fileName)
    fileName = Right(fileName, Len(fileName) - 1)
    fileName = Left(fileName, Len(fileName) - 1)
    fileName = Trim(fileName)
    
    Sheets("Designs Log").Range("S3").Value = fileName
    
End Sub

'Reads a csv file given the file name from the designs folder
'After reading, contents are saved in the Edit Design page
Sub readInputFile()
    Const fileNumber As Integer = 1
    Dim fileName As String
    Dim filePath As String
    Dim rowNumber As Integer
    Dim lineFromFile As String
    Dim lineItems As Variant
    
    filePath = Sheets("Designs Log").Range("AA7").Value
    
    Open filePath For Input As #fileNumber
    rowNumber = 0
    
    'Activate the first line where the input must be saved
    Sheets("Editing Page").Activate
    Sheets("Editing Page").Range("A8").Select
    
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineFromFile
            
        'Split the comma delimeted line into an array
        lineItems = Split(lineFromFile, ",")
        ActiveCell.Offset(rowNumber, 0).Value = lineItems(0)              'Item Code
        ActiveCell.Offset(rowNumber, 1).Value = lineItems(1)                'Item Value
        ActiveCell.Offset(rowNumber, 2).Value = lineItems(2)                'Item Units
        'ActiveCell.Offset(rowNumber, 3).Value = lineItems(3)                'Item Description

        rowNumber = rowNumber + 1
    Loop
    Close #fileNumber
    
End Sub

'Check who owns the current file to be edited
'If the current user owns the file, prompt to overwrite the file or create a copy of the file
'If the current user does not own the file, force use to create a new file
Sub checkOwnership()
    Dim currentUser As String
    Dim fileOwnder As String
    
    
End Sub

'Creates a copy of the selected file and adds the newly created file to the logs
Sub createCopy()

End Sub

'Creates and displays a message box with error messages
Sub reportError(errorMessage As String)
    MsgBox errorMessage, vbCritical
End Sub
