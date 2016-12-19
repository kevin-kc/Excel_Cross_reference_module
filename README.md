# Excel_Cross_reference_module
summary: small script adding an excel auto execute function allowing for cross linking between sheets, sample excel file in directory as an example.

/* Function LookupMatches - for cross linking two worksheets. specifications: see the column “Mappings if there is more than one", function should be generic, allow user to specify the required table names and field names as properties.

E.g. when the function is run on the ECAR tab LookupMaches({ECAR Table Name},“Full Name”,{TBS Table Name},”ECAR Full Name”,”Full Name”)

P1. and P2. First Table and Field to match P3. And P4. Second Table and Field to match, P5. Column containing the values to return as a comma separated list.

*/

  Public Function LookupMatches(startTableName As Variant, startFullName As String, targetTableName As Variant, targetFullName As String, _ targetColumnName As String) As String Dim TableOne As Range Dim TableTwo As Range Set TableOne = startTableName Set TableTwo = targetTableName

    Dim startColNum As Integer
    startColNum = TableOne.Cells.Find(startFullName, , xlValues, xlWhole).Column
    Dim TargetColNum As Integer
    TargetColNum = TableTwo.Cells.Find(targetFullName, , xlValues, xlWhole).Column
    Dim MatchColNum As Integer
    MatchColNum = TableTwo.Cells.Find(targetColumnName, , xlValues, xlWhole).Column

    Dim Matches As String
    Dim j As Integer

    Matches = ""

    For j = 1 To TableTwo.rows.Count
        If TableOne.Cells(CInt(Application.Caller.row - 3), startColNum) = TableTwo.Cells(j, TargetColNum) Then
            Matches = Matches & TableTwo.Cells(j, MatchColNum).value & ", "
        End If
    Next

    If Right(Matches, 2) = ", " Then
        Matches = Left(Matches, Len(Matches) - 2)
    End If

    LookupMatches = Matches
  End Function
