Sub Appear()
    Dim name As String
    Dim depart As String
    Dim used As String
    Dim supply As String
    Dim status As String
    Dim fix_num As String
    Dim Destination1 As String
    Dim Destination2 As String
    Dim imageFile As String
    Dim img_name As String

    fix_num = Sheets("Input").Cells(3, 3)

    If Trim(fix_num) = "" Then
        MsgBox "Please enter a fixture number in cell C3.", vbExclamation, "No Input"
        Exit Sub
    End If

    On Error Resume Next
    name = Application.WorksheetFunction.VLookup(fix_num, Sheets("Fixtures").Range("B1:I1000"), 3, False)
    On Error GoTo 0

    If name = "" Then
        MsgBox "Fixture '" & fix_num & "' was not found in the database.", vbExclamation, "Not Found"
        Exit Sub
    End If

    depart = Application.WorksheetFunction.VLookup(fix_num, Sheets("Fixtures").Range("B1:I1000"), 4, False)
    used = Application.WorksheetFunction.VLookup(fix_num, Sheets("Fixtures").Range("B1:I1000"), 5, False)
    supply = Application.WorksheetFunction.VLookup(fix_num, Sheets("Fixtures").Range("B1:I1000"), 6, False)
    status = Application.WorksheetFunction.VLookup(fix_num, Sheets("Fixtures").Range("B1:I1000"), 7, False)

    Sheets("Input").Cells(11, 5) = name
    Sheets("Input").Cells(13, 5) = depart
    Sheets("Input").Cells(15, 5) = used
    Sheets("Input").Cells(17, 5) = supply
    Sheets("Input").Cells(19, 5) = status

    img_name = fix_num & ".HEIC"
    Destination1 = "\\clcrs010\CLCOvensGroup\Manufacturing\Pub\BTP 2020\ME Projects\Jigs and Fixtures\Photo Archive\"
    Destination2 = "\Pictures\"
    imageFile = Destination1 & fix_num & Destination2 & img_name

    If Dir(imageFile) = "" Then
        MsgBox "Record loaded. Photo not found for fixture '" & fix_num & "'.", vbInformation, "Photo Missing"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sheets("Input")

    Dim pic As Object
    Set pic = ws.Shapes.AddPicture(imageFile, msoFalse, msoTrue, 0, 0, -1, -1)
    pic.Top = ws.Range("E26").Top
    pic.Left = ws.Range("E26").Left
    pic.Width = ws.Range("E26").Width
    pic.Height = ws.Range("E26").Height

End Sub
