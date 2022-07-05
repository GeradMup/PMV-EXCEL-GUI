Attribute VB_Name = "UsersPage"
'THIS MODULE IS FOR USER HANDLING AND GUIDES THEM TO THE NEXT PAGE REQUIRED FOR THE DESIGN PROCESS TO COMMENCE

'=============================================================================================================
' PAGE NAVIGATION
'=============================================================================================================
Sub add_remove_user()
    'This turns off screen updating while code runs to improve efficiency of execution
    Application.ScreenUpdating = False
    Sheets("Users").Select
    Range("DA100:EH150").Select
    Application.ScreenUpdating = True
    ActiveWindow.Zoom = True
    Range("DA100").Select
    
End Sub

Sub back_to_user_page()

    'This turns off screen updating while code runs to improve efficiency of execution
    Application.ScreenUpdating = False
    Sheets("Users").Select
    Range("A1:AC55").Select
    Application.ScreenUpdating = True
    ActiveWindow.Zoom = True
    Range("A1").Select

End Sub

'============================================================================================================
' USER SELECTIONS
'============================================================================================================

 Sub user_select_dw()
 
 'The designs log needs to be updated to contain the most recent designs that has been added.
 update_designs_log
 
 'After updating the list the required data must be read in
 read_in_designs_log_data
 
 
 ' The active user credentials must be written to the appropriate field on the program data page
    Dim userOne As String
 'Users details are contained on the program data page so that if employees change their details can be updated
    userOne = UserPage.Range("DP119").Value
 
 'Active user is set to user 1
    UserPage.Range("DG139").Value = userOne
 
 'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select
 
 End Sub
 
Sub user_select_pmv()

 'The designs log needs to be updated to contain the most recent designs that has been added.
    update_designs_log
 
'After updating the list the required data must be read in
     read_in_designs_log_data
 
' The active user credentials must be written to the appropriate field on the program data page
    Dim userTwo As String
'Users details are contained on the program data page so that if employees change their details can be updated
    userTwo = UserPage.Range("DP120").Value
 
 'Active user is set to user 2
    UserPage.Range("DG139").Value = userTwo
 
'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select

End Sub

Sub user_select_pr()
    
 'The designs log needs to be updated to contain the most recent designs that has been added.
    update_designs_log
 
'After updating the list the required data must be read in
    read_in_designs_log_data
 
' The active user credentials must be written to the appropriate field on the program data page
    Dim userThree As String
'Users details are contained on the program data page so that if employees change their details can be updated
    userThree = UserPage.Range("DP121").Value
 
'Active user is set to user 3
    UserPage.Range("DG139").Value = userThree
 
'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select
        
End Sub

Sub user_select_gm()

 'The designs log needs to be updated to contain the most recent designs that has been added.
    update_designs_log
 
'After updating the list the required data must be read in
    read_in_designs_log_data
 
' The active user credentials must be written to the appropriate field on the program data page
    Dim userFour As String
'Users details are contained on the program data page so that if employees change their details can be updated
    userFour = UserPage.Range("DP122").Value
    
'Active user is set to user 4
    UserPage.Range("DG139").Value = userFour
 
'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select
End Sub

Sub user_select_dm()

'The designs log needs to be updated to contain the most recent designs that has been added.
    update_designs_log
 
'After updating the list the required data must be read in
    read_in_designs_log_data
  
' The active user credentials must be written to the appropriate field on the program data page
    Dim userFive As String
 'Users details are contained on the program data page so that if employees change their details can be updated
    userFive = UserPage.Range("DP123").Value
 
 'Active user is set to user 5
    UserPage.Range("DG139").Value = userFive
  
'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select
End Sub
 
Sub user_select_other()

'It is essential to acquire the initials of the user that will be using the program since it is a required field within the program design files
    Dim userInitials
    userInitials = InputBox("Please insert your initials")
    
    UserPage.Range("DP124").Value = userInitials
'The designs log needs to be updated to contain the most recent designs that has been added.
    update_designs_log
 
'After updating the list the required data must be read in
    read_in_designs_log_data
 
' The active user credentials must be written to the appropriate field on the program data page
    Dim userSix As String
 
'Users details are contained on the program data page so that if employees change their details can be updated
    userSix = UserPage.Range("DP124").Value
 
'Active user is set to user 6
    UserPage.Range("DG139").Value = userSix
 
'Direct User To Design Logs Sheet
    'This turns off screen updating while code runs to improve efficiency of execution
        Application.ScreenUpdating = False
        Sheets("Designs Log").Activate
        Range("A1:X40").Select
        Application.ScreenUpdating = True
        ActiveWindow.Zoom = True
        Range("A1").Select
End Sub

'============================================================================================================
' UPDATING DESIGNS LOG
'============================================================================================================
Sub update_designs_log()
'This module updates the designs log so that it has the latest designs added

    Dim objShell As Object

'These variables are used to store the path to the python executable file as well as the script file required to run with the quotation marks for spaces added
    Dim pythonExePath, MainModelPath As String

    Set objShell = VBA.CreateObject("Wscript.Shell")
    
'For Sheets it works with single quotation marks when reading from the cell
    
    pythonExePath = Sheets("Users").Range("DG140").Value
    MainModelPath = Sheets("Users").Range("DG144").Value
  
    'pythonExePath = """C:\Python310\python.exe"""
    'MainModelPath = """T:\Shared\ACTOM DESP and MOTOR Proto Programme\05 Actom-PMV-GUI\src\Models\MainModel.py"""
    
'Pass in the zero as the second parameter to hide the script
    objShell.Run pythonExePath & MainModelPath, 0
     
End Sub

'============================================================================================================
' READ IN LIST OF DESIGNS AFTER UPDATING DESIGNS LOG
'============================================================================================================

Sub read_in_designs_log_data()
'This module reads in the required data
    
    Const fileNumber As Integer = 1
    Dim fileName, filePath, lineFromFile As String
    
    Dim rowOffset As Integer
    
'Since we are reading from a csv file the comma will be used as the delimiting factor and the items are stored in an array
    Dim lineContents As Variant

'The filepath is stored on the program data page
    filePath = UserPage.Range("DG148").Value
    
    Open filePath For Input As #fileNumber
    rowOffset = 0
    
'Activate the first line where the input must be saved
    DesignsLogPage.Activate
    DesignsLogPage.Range("A8").Select
    
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineFromFile
            
'Split the comma delimeted line into an array
        lineContents = Split(lineFromFile, ",")
        
        ActiveCell.Offset(rowOffset, 0).Value = lineContents(0)                'Engineer that created the design
        ActiveCell.Offset(rowOffset, 1).Value = lineContents(1)                'Design Job Number
        ActiveCell.Offset(rowOffset, 2).Value = lineContents(2)                'Frame Designation
        ActiveCell.Offset(rowOffset, 3).Value = lineContents(3)                'kW Rating
        ActiveCell.Offset(rowOffset, 4).Value = lineContents(4)                'Poles
        ActiveCell.Offset(rowOffset, 5).Value = lineContents(5)                'Line Voltage
        ActiveCell.Offset(rowOffset, 6).Value = lineContents(6)                'Frequency
        ActiveCell.Offset(rowOffset, 7).Value = lineContents(7)                'VSD
        ActiveCell.Offset(rowOffset, 8).Value = lineContents(8)                'Insulation Class
        ActiveCell.Offset(rowOffset, 9).Value = lineContents(9)                'Number of Stator Slots
        ActiveCell.Offset(rowOffset, 10).Value = lineContents(10)              'Number of Rotor Slots
        ActiveCell.Offset(rowOffset, 11).Value = lineContents(11)              'Skew
        ActiveCell.Offset(rowOffset, 12).Value = lineContents(12)              'MATBAR
        ActiveCell.Offset(rowOffset, 13).Value = lineContents(13)              'Filename
        
        rowOffset = rowOffset + 1
    Loop
    Close #fileNumber
    
End Sub
