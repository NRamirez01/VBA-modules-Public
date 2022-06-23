Sub UpdateOOR()

Dim parentFile As Workbook
Dim childFile As Workbook
Dim parentPath As String
Dim childPath As String

Dim FilePicker As FileDialog

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
       .Title = "Select Parent file"
       .AllowMultiSelect = False
       If .Show <> -1 Then GoTo NextCode
       parentPath = .SelectedItems(1)
    
    End With
    




    With FilePicker
       .Title = "Select file to update"
       .AllowMultiSelect = False
       If .Show <> -1 Then GoTo NextCode
       childPath = .SelectedItems(1)
    
    End With
    

    
NextCode:
 parentPath = parentPath
 childPath = childPath
 If childPath = "" Or parentPath = "" Then GoTo ResetSettings
 
 
Set childFile = Workbooks.Open(childPath)
Set parentFile = Workbooks.Open(parentPath)

DoEvents

parentFile.Activate

numRowsParent = Range("C3", Range("C3").End(xlDown)).Rows.Count


childFile.Activate

Range("P2").Select
Selection.AutoFill Destination:=Range("P2:R2"), Type:=xlFillDefault


Range("Q2").Select
ActiveCell.FormulaR1C1 = "Comment"




Range("R2").Select
ActiveCell.FormulaR1C1 = "Work Center"



Dim Destination As String
Dim SKU As String
Dim PO As String
Dim rowChild As Integer
Dim rowCheck As Integer
Dim rowParent As Integer
Dim parentIndex As Integer
Dim workCenter As String
Dim comment As String


NumRowsChild = Range("C3", Range("C3").End(xlDown)).Rows.Count


Range("C3").Select

rowChild = 3

While rowChild <= NumRowsChild
  childFile.Activate
  SKU = Cells(rowChild, 3).Text
  PO = Cells(rowChild, 5).Text
  Debug.Print "Row number is" + " " + CStr(rowChild) + "SKU is " + CStr(SKU) + "PO is " + CStr(PO)
      For rowParent = 3 To numRowsParent
         parentFile.Activate
         If Cells(rowParent, 3).Text = SKU Then
            If Cells(rowParent, 5).Text = PO Then
               Location = Cells(rowParent, 18)
               comment = Cells(rowParent, 17)
               childFile.Activate
               Cells(rowChild, 17) = comment
               Cells(rowChild, 18) = Location
            End If
         End If
     If rowParent = numRowsParent + 1 Then Exit For
        Next rowParent

rowChild = 1 + rowChild
Wend

childFile.Activate

Columns("Q:Q").Select
Selection.Columns.AutoFit
Columns("R:R").Select
Selection.Columns.AutoFit


ResetSettings:
    Application.EnableEvents = True
    Application.Calculation = xlcalculatioautomatic
    Application.ScreenUpdating = True
    



End Sub



