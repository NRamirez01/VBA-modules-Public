Sub updateInProcess()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False


Dim parentFile As Workbook
Dim parentPath As String
Dim specFrm As specVariationDialogue


Dim FilePicker As FileDialog

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

parentFile.Sheets("Raw Data").Activate
numRowsWithSpecs = parentFile.Worksheets("Raw Data").Range("A2", Range("A2").End(xlDown)).Rows.Count + 1

Set specFrm = New specVariationDialogue

parentFile.Activate

rowDimension = 2

While rowDimension <= numRowsWithSpecs
    With specFrm.specBox
        .AddItem (Cells(rowDimension, 1).Text)
    End With
    rowDimension = rowDimension + 1
Wend

Dim newBook As Workbook

Dim specsVary As Object
Set specsVary = CreateObject("System.Collections.ArrayList")

With specFrm
 .Show
 For i = 0 To .specBox.ListCount - 1
   If .specBox.Selected(i) Then
   specsVary.Add i
   End If
   Next i
End With

childCount = InputBox("Enter quantity of new checksheets", "How many new checksheets?")

Dim specI As Integer
Dim sh As Worksheet
Dim rev As String
Dim issued As String
Dim revDate As String
Dim approved As String

rev = parentFile.Worksheets("Sheet1").Cells(2, 6).Text
issued = parentFile.Worksheets("Sheet1").Cells(2, 24).Text
revDate = parentFile.Worksheets("Sheet1").Cells(3, 24).Text
approved = parentFile.Worksheets("Sheet1").Cells(4, 24).Text

Dim updateFrm As specUpdate

parentFile.AutoSaveOn = False

For i = 1 To childCount
With parentFile
    Set main = Workbooks.Open(parentPath)
    newPart = InputBox("What is the new part number?", "For new part " & i)
    newDesc = InputBox(main.Sheets("sheet1").Cells(2, 9).Text, "What is the new description for part " & newPart)
    main.Sheets("Raw Data").Activate
    For Each sh In main.Sheets
        If sh.Name = "Raw Data" Then
            For specI = 0 To specsVary.Count - 1
                Set updateFrm = New specUpdate
                With updateFrm
                .Caption = "Update dimensions for part " & newPart
                .Row = specsVary(specI) + 2
                .Balloon = sh.Cells(specsVary(specI) + 2, 1).Text
                .Dimension = sh.Cells(specsVary(specI) + 2, 4).Text
                .Method = sh.Cells(specsVary(specI) + 2, 2).Text
                .Upper = sh.Cells(specsVary(specI) + 2, 6).Text
                .Lower = sh.Cells(specsVary(specI) + 2, 5).Text
                .Show
                End With

            Next specI
        End If
        If sh.Name <> "Raw Data" Then
            sh.Cells(2, 2).Value = newPart
            sh.Cells(2, 6).Value = rev
            sh.Cells(2, 9).Value = newDesc
            sh.Cells(2, 24).Value = issued
            sh.Cells(3, 24).Value = revDate
            sh.Cells(3, 24).ShrinkToFit = True
            sh.Cells(4, 24).Value = approved
        End If
    Next sh
    main.SaveAs Filename:=newPart & "_r" & LCase(rev) & "-CHECKSHEET.xlsx"
    main.Close
End With

Next i




End Sub

