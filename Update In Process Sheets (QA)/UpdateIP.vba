Sub updateInProcess()


Dim parentFile As Workbook
Dim parentPath As String
Dim parentLayout As Range
Dim childCount As Integer
Dim frm As UserForm1



Dim FilePicker As FileDialog

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
       .Title = "Select parent file"
       .AllowMultiSelect = False
       If .Show <> -1 Then GoTo NextCode
       parentPath = .SelectedItems(1)

    End With

NextCode:
 parentPath = parentPath

Set parentFile = Workbooks.Open(parentPath)
numRowsParent = Range("A9", Range("A9").End(xlDown)).Rows.Count

Set frm = New UserForm1

parentFile.Activate

rowSpec = 9


While rowSpec <= numRowsParent + 9
    With frm.ListBox1
        .AddItem (Cells(rowSpec, 1).Text)
    End With
    rowSpec = rowSpec + 1
Wend

Dim specs As Object
Set specs = CreateObject("System.Collections.ArrayList")

With frm
    .Show
    For i = 0 To .ListBox1.ListCount - 1
        If .ListBox1.Selected(i) Then
            specs.Add i
        End If
        Next i


End With

Dim sourceSheet As Worksheet
Dim rev As String
Dim newPart As String
Dim newSpec As String
Dim newDesc As String
Dim newUpper As String
Dim newBook As Workbook
Dim specEdit As Integer

childCount = InputBox("New IP sheet quantity", "Please enter how many parts to make")

parentFile.Activate
Set sourceSheet = ActiveWorkbook.Worksheets("sheet1")

rev = Cells(2, 6).Text

    For i = 1 To childCount
        Set newBook = Workbooks.Add
        sourceSheet.Copy before:=newBook.Sheets(1)
            newPart = InputBox("What is the new part number?", "New part: " & i)
            newDesc = InputBox(ActiveSheet.Cells(2, 9).Text, "What is the new description?")
            ActiveSheet.Cells(2, 2).Value = newPart
            ActiveSheet.Cells(2, 9).Value = newDesc
            For specEdit = 0 To specs.Count - 1
                newSpec = InputBox(ActiveSheet.Cells(specs(specEdit) + 9, 1).Text, "What is the new value for the following dimension:")
                ActiveSheet.Cells(specs(specEdit) + 9, 1).Value = newSpec
                newUpper = InputBox(ActiveSheet.Cells(specs(specEdit) + 9, 5).Text, "What is the new upper tolerance for dim: " & newSpec)
                ActiveSheet.Cells(specs(specEdit) + 9, 5).Value = newUpper
                newLower = InputBox(ActiveSheet.Cells(specs(specEdit) + 9, 6).Text, "What is the new lower tolerance for dim: " & newSpec)
                ActiveSheet.Cells(specs(specEdit) + 9, 6).Value = newLower
            Next specEdit
        newBook.SaveAs Filename:=newPart & "_r" & LCase(rev) & "-MDR Inspection IP QI Sheet" & ".xlsx"
        newBook.Close
    Next i


ResetSettings:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


End Sub