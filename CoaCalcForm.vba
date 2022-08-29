Private Sub UserForm_Initialize()

ActiveTb.Value = ""
DependencyTb.Value = ""
HousingTb.Value = ""
ResidencyTb.Value = ""
EnrollmentTb.Value = ""
SemestersTb.Value = ""

End Sub

' --------------------------------------------------------------------------------------------------------------

Private Sub CalculateBtn_Click()

Dim coa_key As String      ' key for the coa collection
Dim type_std As String     ' type of student based on dependency, housing, and residency status.
Dim type_enr As String     ' type of enrollment

Dim a As Integer            ' column for ActiveCell
Dim e As Integer            ' column for Enrollment
Dim d As Integer            ' column for Dependency
Dim h As Integer            ' column for Housing
Dim r As Integer            ' column for Residency
Dim curr_row As Integer     ' current row

curr_row = 2
a = Columns(ActiveTb.Value).Column
e = Columns(EnrollmentTb.Value).Column
d = Columns(DependencyTb.Value).Column
h = Columns(HousingTb.Value).Column
r = Columns(ResidencyTb.Value).Column

' Filling in collection "coa" with cost of attendance information from the "COA Simplified" sheet.
' First number in the key indicates student type. Second number indicates enrollment type.
Dim coa As Collection
Set coa = New Collection

' (1) Dependent, living with parents or off-campus, in-state
coa.Add Sheets("COA Simplified").Range("B10").Value, "11"
coa.Add Sheets("COA Simplified").Range("C10").Value, "12"
coa.Add Sheets("COA Simplified").Range("D10").Value, "13"
coa.Add Sheets("COA Simplified").Range("E10").Value, "14"
' (2) Dependent, living on campus, in-state
coa.Add Sheets("COA Simplified").Range("B9").Value, "21"
coa.Add Sheets("COA Simplified").Range("C9").Value, "22"
coa.Add Sheets("COA Simplified").Range("D9").Value, "23"
coa.Add Sheets("COA Simplified").Range("E9").Value, "24"
' (3) Independent, living with parents or off-campus, in-state
coa.Add Sheets("COA Simplified").Range("B4").Value, "31"
coa.Add Sheets("COA Simplified").Range("C4").Value, "32"
coa.Add Sheets("COA Simplified").Range("D4").Value, "33"
coa.Add Sheets("COA Simplified").Range("E4").Value, "34"
' (4) Independent, living on campus, in-state
coa.Add Sheets("COA Simplified").Range("B3").Value, "41"
coa.Add Sheets("COA Simplified").Range("C3").Value, "42"
coa.Add Sheets("COA Simplified").Range("D3").Value, "43"
coa.Add Sheets("COA Simplified").Range("E3").Value, "44"
' (5) Dependent/Independent, living with parents or off-campus, Out-Of-state
coa.Add Sheets("COA Simplified").Range("B16").Value, "51"
coa.Add Sheets("COA Simplified").Range("C16").Value, "52"
coa.Add Sheets("COA Simplified").Range("D16").Value, "53"
coa.Add Sheets("COA Simplified").Range("E16").Value, "54"
' (6) Dependent/Independent, living on campus, Out-of-state
coa.Add Sheets("COA Simplified").Range("E15").Value, "61"
coa.Add Sheets("COA Simplified").Range("E15").Value, "62"
coa.Add Sheets("COA Simplified").Range("E15").Value, "63"
coa.Add Sheets("COA Simplified").Range("E15").Value, "64"


For counter = 1 To SemestersTb.Value        ' Counter 1 checks for selected semester from the beginning of the list
    If counter > 1 Then     ' Checks for subsequent semester from the beginning of the list
        a = a + 1  ' Incrementing ActiveCell column
        e = e + 1  ' Incrementing Enrollment column
        curr_row = 2
    End If

    ' Loops until column 'd' runs out of rows
    ' Fills rows in column 'a' with expected COA calculation
    ' Checks student dependency, living status, and residency status first
    ' After determining student's information, evaluates COA based on enrollment
    Do Until IsEmpty(Cells(curr_row, d))
        ActiveCell.Cells(curr_row, a).Select

        ' Dependent, living with parents or off-campus, in-state
        If Cells(curr_row, d).Value = "D" And _
           (Cells(curr_row, h).Value = "WITH_PARENT" Or Cells(curr_row, h).Value = "OFF_CAMPUS") And _
           Cells(curr_row, r) = "IN" Then
            type_std = "1"

        ' Dependent, living on campus, in-state
        ElseIf Cells(curr_row, d).Value = "D" And _
               Cells(curr_row, h).Value = "ON_COMPUS" And _
               Cells(curr_row, r) = "IN" Then
            type_std = "2"

        ' Independent, living with parents or off-campus, in-state
        ElseIf Cells(curr_row, d).Value = "I" And _
               (Cells(curr_row, h).Value = "WITH_PARENT" Or Cells(curr_row, h).Value = "OFF_CAMPUS") And _
               Cells(curr_row, r) = "IN" Then
            type_std = "3"

        ' Independent, living on campus, in-state
        ElseIf Cells(curr_row, d).Value = "I" And _
               Cells(curr_row, h).Value = "ON_COMPUS" And _
               Cells(curr_row, r) = "IN" Then
            type_std = "4"

        ' Dependent/Independent, living with parents or off-campus, Out-Of-state
        ElseIf (Cells(curr_row, h).Value = "WITH_PARENT" Or Cells(curr_row, h).Value = "OFF_CAMPUS") And _
               Cells(curr_row, r) = "OUT" Then
            type_std = "5"

        ' Dependent/Independent, living on campus, Out-of-state
        ElseIf Cells(curr_row, h).Value = "ON_COMPUS" And Cells(curr_row, r) = "OUT" Then
            type_std = "6"

        ' If none of the above evaluate to true, set the cell to "False" and jump to the NoMatch label.
        Else
            Cells(curr_row, a).Value = False
            GoTo NoMatch
        End If
        
        ' Finding student's enrollment
        If Cells(curr_row, e).Value >= 12 Then
            type_enr = "1"
        ElseIf Cells(curr_row, e).Value >= 9 And Cells(curr_row, e).Value < 12 Then
            type_enr = "2"
        ElseIf Cells(curr_row, e).Value >= 6 And Cells(curr_row, e).Value < 9 Then
            type_enr = "3"
        ElseIf Cells(curr_row, e).Value < 6 And Cells(curr_row, e).Value > 0 Then
            type_enr = "4"
        End If
           
        ' Checking if student is not enrolled.
        If Cells(curr_row, e).Value = 0 Then
            Cells(curr_row, a).Value = 0
        Else
            coa_key = type_std & type_enr
            Cells(curr_row, a).Value = coa(coa_key) ' Printing the COA.
        End If
        
NoMatch:

        ' Increments current row to continue to the next row.
        curr_row = curr_row + 1
    Loop

Next counter

Unload Me

End Sub
