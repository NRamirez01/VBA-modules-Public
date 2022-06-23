Attribute VB_Name = "Module11"
Sub newRevAppend()
    Dim toPath As String
    Dim fso As Object
    Dim targetFolder As Object
    Dim revision As Integer
    Dim starting_ws As Worksheet
    Dim Current As Worksheet
    
    Set starting_ws = ActiveSheet
    Set fso = CreateObject("Scripting.FileSystemObject")

    'fso.CopyFolder "c:\mydocuments\letters\*", "c:\tempfolder\"

    'Set targetFolder = Application.FileDialog(msoFileDialogFolderPicker)

    'With targetFolder
        '.Title = "Target for Archive"
        '.AllowMultiSelect = False
        '.InitialFileName = initialPath
        'If .Show <> -1 Then
           ' MsgBox "You didn't select anything"
           ' Exit Sub
       ' End If
        'toPath = .SelectedItems(1) & "\"
   ' End With
        
    revision = 65
    For Each Current In ThisWorkbook.Worksheets
        Current.Activate
        If Current.Name = "Raw Data" Then
            Current.Cells(23, 4) = "5"
            Current.Cells(23, 5) = "4.8"
            Current.Cells(23, 6) = "5.2"
            Current.Cells(24, 4) = "5"
            Current.Cells(24, 5) = "4.8"
            Current.Cells(24, 6) = "5.2"
            Current.Cells(28, 4) = "3"
            Current.Cells(28, 5) = "2.8"
            Current.Cells(28, 6) = "3.2"
        ElseIf Current.Name = "Sheet1" Then
            Do Until (Chr(revision)) = Current.Cells(2, 6).Value
                revision = revision + 1
            Loop
            Current.Cells(2, 6) = Chr(revision + 1)
            Current.Cells(3, 24).NumberFormat = "@"
            Current.Cells(3, 24) = ("11" + "/" + "11" + "/" + "2021") 'Annoying way to set date
            For i = 65 To 48 Step -1
                If Range(Range("A" & i), Range("F" & i)).Text = "" Then
                Range(Range("A" & i), Range("G" & i)).Select
                    With Selection
                        .EntireRow.Delete
                    End With
                End If
            Next i
            Current.Range("X2:Y2").Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                 Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        ElseIf Current.Name <> "Raw Data" And _
        Current.Name <> "Sheet1" And _
        Current.Range("X2") <> "" Then
            Current.Range("X2:Y2").Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                 Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            Current.Cells(2, 6) = Worksheets("Sheet1").Cells(2, 6)
            Current.Cells(3, 24) = Worksheets("Sheet1").Cells(3, 24)
        End If
        If Current.Range("X2") = "" _
        And Current.Range("D10") = 0 Then
            Application.DisplayAlerts = False
            Current.Delete
        End If
    Next Current
End Sub


