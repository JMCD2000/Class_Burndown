Option Explicit

Function vertLookUp_Event_Dates(ByVal HullNum As Long, ByVal GetEventDate As String) As Variant
'This function looks up the date of a passed in Event and returns a Date
    Dim myLookupValue As String 'Passed in String of the Event Name
    'Dim myFirstColumn As Long 'Starting Column left most
    'Dim myLastColumn As Long 'Ending Column right most
    Dim myColumnIndex As Long 'is the return value column in the matched row
    'Dim myFirstRow As Long 'Upper Row
    'Dim myLastRow As Long 'Lowwer Row
    Dim myVLookupResult As Variant 'This will be the function return date
    'Dim myTableArray As Range 'This is the Range Array
 
    Select Case GetEventDate
        Case "BT"
            myColumnIndex = 3
        Case "AT"
            myColumnIndex = 5
        Case "DEL"
            myColumnIndex = 6
        Case "SAIL"
            myColumnIndex = 12
        Case "FCT"
            myColumnIndex = 18
        Case "EXP"
            myColumnIndex = 22
        Case "OWLD"
            myColumnIndex = 23
        End Select
    
    myLookupValue = HullNum
    'myFirstColumn = 1
    'myLastColumn = 23
    'myColumnIndex = x 'Is assigned via Select
    'myFirstRow = 21
    'myLastRow = 31
    'With Worksheets("Hull Milestone Dates")
    '    Set myTableArray = .Range(.Cells(myFirstRow, myFirstColumn), .Cells(myLastRow, myLastColumn))
    'End With
    'myVLookupResult = Application.WorksheetFunction.VLookup(myLookupValue, myTableArray, myColumnIndex, False)
    
    myVLookupResult = Application.VLookup(HullNum, ThisWorkbook.Worksheets("Hull Milestone Dates").Range("A21:W31"), myColumnIndex, False)
    Dim datetest As Boolean
    Dim myVLookupResult_date As Date
    datetest = testCDate(myVLookupResult)
    If datetest = True Then
        myVLookupResult_date = CDate(myVLookupResult)
        vertLookUp_Event_Dates = myVLookupResult_date
    Else
        Debug.Print ("myVLookupResult : " & myVLookupResult)
        Debug.Print ("myVLookupResult_date : " & myVLookupResult_date)
        vertLookUp_Event_Dates = myVLookupResult
    End If
        
End Function

Public Function RunAllSheets() As Boolean
'This is to run through all the data sheets
    Dim myResults As Boolean
    
    'Make speed and screen changes
    'Application.ScreenUpdating = False 'This stops the redrawing of the screen while updating each row
    'Application.Calculation = xlCalculationManual 'This stops other formulas from updating while writing data to this sheet
    
    Dim mySheetNamesArray As Variant
    mySheetNamesArray = Array("17 Cards", "18 Cards", "19 Cards", "20 Cards", "21 Cards", "22 Cards", "23 Cards", "24 Cards", "25 Cards", "26 Cards", "27 Cards")
    
    Dim CurSheet As Variant
    For Each CurSheet In mySheetNamesArray
        Debug.Print ("Starting CurSheet : " & CurSheet)
        myResults = ReFresh_WorkSheet_Data(CurSheet)
        Debug.Print ("Completed CurSheet : " & CurSheet)
    Next
    
    RunAllSheets = myResults
    
    'Un do speed and screen changes
    'Application.ScreenUpdating = True 'This turns back on the screen updating
    'Application.Calculation = xlCalculationAutomatic 'This turns back on the auto calculation

    'Tell me it has completed
    'If myResults = False Then
    '    MsgBox "Refresh has Failed.", vbOKOnly, "Data Refresh"
    'ElseIf myResults = True Then
    '    MsgBox "Refresh was successfull!", vbOKOnly, "Data Refresh"
    'Else
    '    MsgBox "Function/Sub error,Refresh has Failed", vbOKOnly, "Data Refresh"
    'End If

End Function

Function testCDate(testDate As Variant) As Boolean
'This Function test for data type miss match ERRORs _
and fails with out stopping beacuse of the conversion _
ERROR. As a stand alone function the ERROR is trapped _
and the program contiuns

    On Error Resume Next
    testCDate = Not IsError(CDate(testDate))
    
End Function

Function ConvertDateToNum(myDateVal As Variant) As Variant
'This Function will convert a Date to its numerical value

    ConvertDateToNum = (myDateVal) - (#1/1/1900#) + 2
    
End Function

Public Function Convert_OracleDate_to_MicrosoftDate(ByVal Passed_Date_String As String) As Date 'As Variant
'This Function converts the Oracle Date format to the Microsoft Date format _
that is needed inorder to compute the Date Diference values. The return _
as Date has values of an actual date closed or if no date the empty date _
value is #12:00:00 AM#
'TODO: If there is a bad date value the default is Open, The below _
commented out functions is part of the solution but gets messy when expecting #DATE#
    
    'Declare my Var
    Dim temp_DATE_str As String 'temp value while converting from Oracle Date to MS Date
    Dim Return_Date_date As Date 'As Variant 'This is returned to the caller
    
    'Replace the Oracle Hyphens with Microsoft Forward Slashes
    temp_DATE_str = Replace(Passed_Date_String, "-", "/")
    'If empty or not found no ERROR is raised
    
    If IsDate(temp_DATE_str) = True Then
    'IsDate is True if it looks like a Microsoft date in the range of Microsoft dates
        If temp_DATE_str = "" Then
        'Empty String becomes an Empty Date
            Return_Date_date = Empty 'The Empty Date value is #12:00:00 AM#
        Else
        'Date String is converted to a Date data type
            Return_Date_date = CDate(temp_DATE_str)
        End If
    ElseIf testCDate(temp_DATE_str) = True Then
    'Call my UDF to trap miss match ERRORs, The CDate() throws and ERROR if string is not a date
        Return_Date_date = CDate(temp_DATE_str)
    
    'Bad cell values trapping
    'ElseIf testCDate(temp_DATE_str) = False Then
    'Call my UDF to trap miss match ERRORs, The CDate() returned value is not a date
    '    Return_Date_date = "Date ERROR: testCDate=False; String or not Date format"
    'ElseIf IsDate(temp_DATE_str) = False Then
    'IsDate is False if it doesnt look like a date or is outside of the range for Microsoft dates
    '    Return_Date_date = "Date ERROR: IsDate=False; String or not Date format"
        
    Else
    'IsDate is False and testCDate is False
        Return_Date_date = Empty  'The Empty Date value is #12:00:00 AM#
    End If
    
    'Return Value
    Convert_OracleDate_to_MicrosoftDate = Return_Date_date
    
End Function

Public Function TC_Column_Populate( _
    ByVal Trial_ID As String, _
    ByVal Selected_ID_1 As String, _
    ByVal Selected_ID_2 As String, _
    ByVal Selected_ID_3 As String, _
    ByVal Selected_ID_4 As String, _
    ByVal Selected_ID_5 As String, _
    ByVal TC_Event As String, _
    ByVal Selected_Event As String, _
    ByVal Baseline_Date As Date, _
    ByVal TC_Closed_Date As Date, _
    ByVal Event_or_Date_Switch As String _
    ) As Variant
'This Function is to compute the Column value regardless if it is counted in another column. _
The varibles are passed in By Value to avoid changing the varibles. The computed varible _
is returned based or requested data type, a string for the Event or a date for the days _
from the baseline Event Date as a Date. This Funtion only returns one varible and is used _
to compute values using the same criterion and is written only once.
    
    'Declare my Var
    Dim Return_Event As String 'This is the assigned string for the Event
    Dim Return_Date As Variant 'This is the Date Difference for the TC_Closed_Date minus Baseline_Date
    
    If Len(Trial_ID) < 1 Then
    'Trial ID is Empty, only looking for Selected Event(s) cards
        If InStr(Selected_Event, TC_Event) > 0 Then
        'This is looking for Split Selected Event(s) cards that had _
        the Selected Trial ID removed but is not from another Event
            'It contains the Selected Event
            If Event_or_Date_Switch = "ByEvent" Then
                Return_Event = TC_Event
            ElseIf Event_or_Date_Switch = "ByDate" Then
                If TC_Closed_Date = Empty Then
                    Return_Date = "OPEN"
                ElseIf TC_Closed_Date = #12:00:00 AM# Then
                    Return_Date = "OPEN"
                ElseIf TC_Closed_Date > #12:00:00 AM# Then
                    Return_Date = TC_Closed_Date - Baseline_Date
                Else
                    'Date closed ERROR
                    Return_Date = "ERROR: Date Closed"
                End If
            Else
                Return_Event = "ERROR: no Switch Value Passed"
                Return_Date = "ERROR: no Switch Value Passed"
            End If
        Else
        'This has no Trial ID and was not written during TB/BT _
        and would be a split card from another event. Record _
        not a TB/BT Event and missing Trial ID and is NOT counted here
            If Event_or_Date_Switch = "ByEvent" Then
                Return_Event = "NULL"
            ElseIf Event_or_Date_Switch = "ByDate" Then
                Return_Date = "NULL"
            Else
                Return_Event = "ERROR: no Switch Value Passed"
                Return_Date = "ERROR: no Switch Value Passed"
            End If
        End If
    'Check if Trial Card Trial ID is in the passed in select by Trial IDs
    ElseIf InStr(Trial_ID, Selected_ID_1) _
    Or InStr(Trial_ID, Selected_ID_2) _
    Or InStr(Trial_ID, Selected_ID_3) _
    Or InStr(Trial_ID, Selected_ID_4) _
    Or InStr(Trial_ID, Selected_ID_5) _
     > 0 Then
    'Contains the Selected Trial ID in the TC Trial ID
        If InStr(Selected_Event, TC_Event) > 0 Then
        'It contains the Selected Event
            If Event_or_Date_Switch = "ByEvent" Then
                Return_Event = TC_Event
            ElseIf Event_or_Date_Switch = "ByDate" Then
                If TC_Closed_Date = Empty Then
                    Return_Date = "OPEN"
                ElseIf TC_Closed_Date = #12:00:00 AM# Then
                    Return_Date = "OPEN"
                ElseIf TC_Closed_Date > #12:00:00 AM# Then
                    Return_Date = TC_Closed_Date - Baseline_Date
                Else
                    'Date closed ERROR
                    Return_Date = "ERROR: Date Closed"
                End If
            Else
                Return_Event = "ERROR: no Switch Value Passed"
                Return_Date = "ERROR: no Switch Value Passed"
            End If
        Else
        'It contains the Selected Trial ID and was not written during the _
        Event(s) passed in, or rolled into this Event from a Prior Event
            If Event_or_Date_Switch = "ByEvent" Then
                Return_Event = "NULL"
            ElseIf Event_or_Date_Switch = "ByDate" Then
                Return_Date = "NULL"
            Else
                Return_Event = "ERROR: no Switch Value Passed"
                Return_Date = "ERROR: no Switch Value Passed"
            End If
        End If
    'Check if Trial Card Trial ID is NOT in the passed in select by Trial IDs
    ElseIf InStr(Trial_ID, Selected_ID_1) _
    And InStr(Trial_ID, Selected_ID_2) _
    And InStr(Trial_ID, Selected_ID_3) _
    And InStr(Trial_ID, Selected_ID_4) _
    And InStr(Trial_ID, Selected_ID_5) _
    = 0 Then
    'Does not contain Selected Trial ID in the TC Trial ID
        If Event_or_Date_Switch = "ByEvent" Then
            Return_Event = "NULL"
        ElseIf Event_or_Date_Switch = "ByDate" Then
            Return_Date = "NULL"
        Else
            Return_Event = "ERROR: no Switch Value Passed"
            Return_Date = "ERROR: no Switch Value Passed"
        End If
    Else
    'Trial ID is Empty, and was not written during Selected Event(s) _
    and would be a split card from another event. Record _
    not a Selected Event(s) and is missing Trial ID and is NOT counted here
        If Event_or_Date_Switch = "ByEvent" Then
            Return_Event = "NULL"
        ElseIf Event_or_Date_Switch = "ByDate" Then
            Return_Date = "NULL"
        Else
            Return_Event = "ERROR: no Switch Value Passed"
            Return_Date = "ERROR: no Switch Value Passed"
        End If
    End If
    
    'Return Value
    If Event_or_Date_Switch = "ByEvent" Then
        TC_Column_Populate = Return_Event
    ElseIf Event_or_Date_Switch = "ByDate" Then
        TC_Column_Populate = Return_Date
    Else
        TC_Column_Populate = "ERROR: no Switch Value Passed"
    End If

End Function


Public Function ReFresh_WorkSheet_Data(ByVal CurSheet As String) As Boolean
'This Function takes a passed in worksheet _
name and uses that to refresh the sheet. This Function _
returns True if no ERRORs occur.

    'Declare my Vars
    Dim mySuccess As Boolean 'The return value of this function
    Dim rowPointer As Integer 'Current row copy from, paste too
    Dim myCurRange As Integer 'The upper limit/boundary for the number of rows
    'These are the Reference Varibles
    Dim dt_val_BT_Trial_Date_data As Date 'This is the date of the BT Trial, is used for date diff Col:Rw Q:2
    Dim dt_val_AT_Trial_Date_data As Date 'This is the date of the AT Trial, is used for date diff Col:Rw X:2
    Dim Baseline_Switch_Date As Date 'This uses the BT or AT Trial Date and is Picked based on FULL count ALL or INSURV count INSURV only Col:Rw S:2
    Dim str_val_Count_ALL_or_INSURV_data As String 'This switches the TC counts from FULL count ALL or INSURV count INSURV only Col:Rw S:2
    Dim str_val_Expansion_Event_data As String 'This is for additional Trial Events, is used with SAIL, FCT2, FT, NULL Col:Rw AE:1
    Dim dt_val_Date_of_DEL As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
    Dim dt_val_Date_of_SAIL As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
    Dim dt_val_Date_of_OWLD As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
    Dim dt_val_DaysFrom_Baseline_to_DEL As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
    Dim dt_val_DaysFrom_Baseline_to_SAIL As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
    Dim dt_val_DaysFrom_Baseline_to_OWLD As Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
    
    'These are the Data columns varibles
    Dim str_val_DSP_data As String 'Col A DSP
    Dim str_val_STAR_data As String 'Col B STAR
    Dim str_val_PRI_data As String 'Col C PRI
    Dim str_val_SAFE_data As String 'Col D SAFETY
    Dim str_val_SCREEN_data As String 'Col E SCREEN
    Dim str_val_ACT1_data As String 'Col F ACTION CODE 1
    Dim str_val_ACT2_data As String 'Col G ACTION CODE 2
    Dim str_val_STATUS_data As String 'Col H STATUS
    Dim str_val_ACT_TKN_data As String 'Col I ACTION TAKEN
    Dim var_val_DATE_DISC_data As String 'Col J DATE DISCOVERED
    Dim var_val_DATE_CLOSED_data As String 'Col K DATE CLOSED
    Dim str_val_TRIAL_ID_data As String 'Col L TRIAL ID
    Dim str_val_EVENT_data As String 'Col M EVENT
    'These are the computed new values
    'moved into another function: Dim temp_DATE_DISC_str As String 'temp value while converting from Oracle Date to MS Date
    Dim dt_val_DATE_DISC_data As Variant 'As Date
    'moved into another function: Dim temp_DATE_CLOSED_str As String 'temp value while converting from Oracle Date to MS Date
    Dim dt_val_DATE_CLOSED_data As Variant 'As Date
    Dim var_val_BT_DaysFrom_ALL_data As Variant
    Dim str_val_EVENT_ALL_data As String
    Dim var_val_BT_DaysFrom_data As Variant
    Dim str_val_BT_EVENT_data  As String
    Dim str_val_conEVENT_STAR_PRI_SAFETY_ALL_data As String
    Dim str_val_conEVENT_DATE_CLOSED_ALL_data As String
    Dim var_val_AT_DaysFrom_INSURV_data As Variant
    Dim str_val_EVENT_INSURV_data As String
    Dim str_val_EVENT_AT_ONLY_data As String
    Dim var_val_AT_DaysFrom_AT_ONLY_data As Variant
    Dim str_val_EVENT_FCT_ONLY_data As String
    Dim var_val_AT_DaysFrom_FCT_ONLY_data As Variant
    Dim str_val_EVENT_EXPANTION_ONLY_data As String
    Dim var_val_AT_DaysFrom_EXPANTION_ONLY_data As Variant
    Dim str_val_conEVENT_DATE_CLOSED_INSURV_data As String
    Dim str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data As String
    Dim str_val_EVENT_INSURV_AT_ONLY_data As String
    Dim str_val_SCREEN_G_K_S_data As String
    Dim str_val_conEVENT_SCREEN_AT_ONLY_data As String
    Dim str_val_DEL_BT_DaysFrom_AT_data As String
    Dim str_val_SAIL_BT_DaysFrom_AT_data As String
    Dim str_val_OWLD_BT_DaysFrom_AT_data As String
    Dim str_val_PRI1S_AT_ONLY_data As String
    Dim str_val_PRI1S_DEL_BT_DaysFrom_AT_data As String
    Dim str_val_PRI1S_SAIL_BT_DaysFrom_AT_data As String
    Dim str_val_PRI1S_OWLD_BT_DaysFrom_AT_data As String
    Dim str_val_FCT_FCT_not_in_AT_ONLY_data As String
    Dim str_val_FCT_OWLD_not_in_AT_ONLY_data As String
    Dim str_val_FCT_OWLD_BT_DaysFrom_AT_data As String
    Dim str_val_STAR_NEW_BT_DaysFrom_AT_data As String
    Dim str_val_STAR_DEL_BT_DaysFrom_AT_data As String
    Dim str_val_STAR_SAIL_BT_DaysFrom_AT_data As String
    Dim str_val_STAR_OWLD_BT_DaysFrom_AT_data As String
    
    'Load the Event dates from the Hull Milestones
    Dim getHullNum As Long
    getHullNum = Left(CurSheet, 2)
    'vertLookUp_Event_Dates( Hull Number, (BT,AT,DEL,SAIL,FCT,EXP,OWLD) )
    ThisWorkbook.Worksheets(CurSheet).Range("Q2").Value = vertLookUp_Event_Dates(getHullNum, "BT")
    ThisWorkbook.Worksheets(CurSheet).Range("S2").Value = ThisWorkbook.Worksheets("BT All Hulls Chart Data").Range("CP6").Value
    ThisWorkbook.Worksheets(CurSheet).Range("X2").Value = vertLookUp_Event_Dates(getHullNum, "AT")
    ThisWorkbook.Worksheets(CurSheet).Range("AK2").Value = vertLookUp_Event_Dates(getHullNum, "DEL")
    ThisWorkbook.Worksheets(CurSheet).Range("AL2").Value = vertLookUp_Event_Dates(getHullNum, "SAIL")
    ThisWorkbook.Worksheets(CurSheet).Range("AM2").Value = vertLookUp_Event_Dates(getHullNum, "OWLD")
    ThisWorkbook.Worksheets(CurSheet).Range("AR2").Value = vertLookUp_Event_Dates(getHullNum, "FCT")

    'Assign values to my Vars
    mySuccess = False 'Default as False if exit on ERROR
    rowPointer = 0 'The first row in all the sheets
    myCurRange = ThisWorkbook.Worksheets(CurSheet).Range("A1").CurrentRegion.Rows.Count 'Get the total rows of data as a number
    'Debug.Print (vbTab & "-RF_WK_Data-Count Upper Bound (" & curSheet & ") myCurRange : " & myCurRange)
    dt_val_BT_Trial_Date_data = CDate(ThisWorkbook.Worksheets(CurSheet).Range("Q2").Value) 'Date of the BT Trial
    dt_val_AT_Trial_Date_data = CDate(ThisWorkbook.Worksheets(CurSheet).Range("X2").Value) 'Date of the AT Trial
    Baseline_Switch_Date = Empty 'This is used to switch the baseline date dif between the BT and the AT Trial dates
    str_val_Count_ALL_or_INSURV_data = ThisWorkbook.Worksheets(CurSheet).Range("S2").Value 'This is used in setting the baseline date and the look left columns to prevent double counting
    str_val_Expansion_Event_data = ThisWorkbook.Worksheets(CurSheet).Range("AE1").Value 'This holds the expansion Event name ("SAIL", "TF", "FCT2", "NULL")
    dt_val_Date_of_DEL = ThisWorkbook.Worksheets(CurSheet).Range("AK2").Value 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
    dt_val_Date_of_SAIL = ThisWorkbook.Worksheets(CurSheet).Range("AL2").Value 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
    dt_val_Date_of_OWLD = ThisWorkbook.Worksheets(CurSheet).Range("AM2").Value 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
    dt_val_DaysFrom_Baseline_to_DEL = Empty 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
    dt_val_DaysFrom_Baseline_to_SAIL = Empty 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
    dt_val_DaysFrom_Baseline_to_OWLD = Empty 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff

    For rowPointer = 3 To (myCurRange - 1)
        'rowPointer = row number that is zero based index _
        row 0 is the column header found in A1:Col(X)Row(X) _
        if there are no column headers used then .Offset((rowPointer - 1), 0) to index correctly _
        if there is column headers used then .Offset(rowPointer, 0) to index correctly

        'Debug.Print (vbTab & vbTab & "-RF_WK_Data-Count Upper limit/boundary for total rows, myCurRange : " & myCurRange)
        'Debug.Print (vbTab & vbTab & "-RF_WK_Data-Current selected row rowPointer : " & rowPointer)
        'Debug.Print (vbTab & vbTab & "-RF_WK_Data-Current selected row (rowPointer - 1) : " & (rowPointer - 1))

        'Zero based index pointer is used in the "With" & "Offset" Statements
        With ThisWorkbook.Worksheets(CurSheet).Range("A1")
            '.Offset(rowPointer, 0) is (Variable row pointer Int, Fixed column pointer)
            
        '''Read Data Columns
            str_val_DSP_data = .Offset((rowPointer), 0).Value2 'Col A DSP
            'Debug.Print "str_val_DSP_data :" & str_val_DSP_data 'Used for record data ERRORs
            str_val_STAR_data = .Offset((rowPointer), 1).Value2 'Col B STAR
            str_val_PRI_data = .Offset((rowPointer), 2).Value2 'Col C PRI
            str_val_SAFE_data = .Offset((rowPointer), 3).Value2 'Col D SAFETY
            str_val_SCREEN_data = .Offset((rowPointer), 4).Value2 'Col E SCREEN
            str_val_ACT1_data = .Offset((rowPointer), 5).Value2 'Col F ACTION CODE 1
            str_val_ACT2_data = .Offset((rowPointer), 6).Value2 'Col G ACTION CODE 2
            str_val_STATUS_data = .Offset((rowPointer), 7).Value2 'Col H STATUS
            str_val_ACT_TKN_data = .Offset((rowPointer), 8).Value2 'Col I ACTION TAKEN
            var_val_DATE_DISC_data = .Offset((rowPointer), 9).Value2 'Col J DATE DISCOVERED' = Format(Now, "mm/dd/yyyy")
            'Debug.Print "var_val_DATE_DISC_data :" & var_val_DATE_DISC_data 'Used for record data ERRORs
            var_val_DATE_CLOSED_data = .Offset((rowPointer), 10).Value2 'Col K DATE CLOSED' = Format(Now, "mm/dd/yyyy")
            'Debug.Print "var_val_DATE_DISC_data :" & var_val_DATE_DISC_data 'Used for record data ERRORs
            str_val_TRIAL_ID_data = .Offset((rowPointer), 11).Value2 'Col L TRIAL ID
            str_val_EVENT_data = .Offset((rowPointer), 12).Value2 'Col M EVENT
            
        '''Set the switch baseline to either AT(INSURV) or BT(FULL)
            'TODO: The below is currently implemented in line. When code is refactored _
            the switch assignment should happen here, once. The hard coded cell date refs _
            will also have to be cleaned up.
            'Set the Baseline Date
            'If str_val_Count_ALL_or_INSURV_data = "FULL" Then
            '    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
            'ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
            '    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
            'Else
            dt_val_DaysFrom_Baseline_to_DEL = dt_val_Date_of_DEL - dt_val_AT_Trial_Date_data 'Baseline_Switch_Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
            dt_val_DaysFrom_Baseline_to_SAIL = dt_val_Date_of_SAIL - dt_val_AT_Trial_Date_data 'Baseline_Switch_Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
            dt_val_DaysFrom_Baseline_to_OWLD = dt_val_Date_of_OWLD - dt_val_AT_Trial_Date_data 'Baseline_Switch_Date 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
            
        '''Find out if there is any missing key values
            'TODO: If there is missing a value in the EVENT, the whole record should be skipped
            'IF Len(str_val_EVENT_data) > 0
            'Do the row actions to the current record

            '''Compute new values
            '''Col O TRIM DATE DISCOVERED
                'Call to Date conversion function Convert_OracleDate_to_MicrosoftDate(STRING)DATE
                'TODO: If there is a bad date value the default is Open, The Function call below _
                has commented out sections taht is part of the solution but gets messy when expecting #DATE#
                dt_val_DATE_DISC_data = Convert_OracleDate_to_MicrosoftDate(var_val_DATE_DISC_data)
                'TODO; could write the reformated date back to the data column, as a string
    
            '''Col P TRIM DATE CLOSED
                'Call to Date conversion function Convert_OracleDate_to_MicrosoftDate(STRING)DATE
                'TODO: If there is a bad date value the default is Open, The Function call below _
                has commented out sections taht is part of the solution but gets messy when expecting #DATE#
                dt_val_DATE_CLOSED_data = Convert_OracleDate_to_MicrosoftDate(var_val_DATE_CLOSED_data)
                'TODO; could write the reformated date back to the data column, as a string
    
        '''BT Values
            '''Col Q ALL COL M Days From BT, This is counting ALL cards _
                contained within the sheet. This counts from the date of the TB/BT Trial _
                it is not part of the Baseline Switching.
                'TODO: If there is a bad date value the default is Open, The below _
                commented out functions is part of the solution but gets messy when expecting #DATE#
                'Debug.Print "InStr(dt_val_DATE_CLOSED_data, Error,1) :" & Str(InStr(1, dt_val_DATE_CLOSED_data, "Error", 1))
                'If InStr(1, dt_val_DATE_CLOSED_data, "Error", 1) > 0 Then
                'Date Closed has an ERROR most likley a word string
                    'var_val_BT_DaysFrom_ALL_data = dt_val_DATE_CLOSED_data
                'ElseIf InStr(1, dt_val_DATE_CLOSED_data, "Error", 1) = 0 Then
                'Date Closed has a valid date or is empty and thus still open
                    If dt_val_DATE_CLOSED_data = Empty Then
                        var_val_BT_DaysFrom_ALL_data = "OPEN"
                    ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        var_val_BT_DaysFrom_ALL_data = "OPEN"
                    ElseIf dt_val_DATE_CLOSED_data > #12:00:00 AM# Then
                        var_val_BT_DaysFrom_ALL_data = dt_val_DATE_CLOSED_data - dt_val_BT_Trial_Date_data
                    Else
                        var_val_BT_DaysFrom_ALL_data = "ERROR"
                    End If
                'Else
                '    var_val_BT_DaysFrom_ALL_data = "ERROR: Date Closed has bad data"
                'End If
    
            '''Col R ALL COL M EVENT, This is Listing ALL cards _
                contained within the sheet. This is redundant, but allows for manual _
                changes without changing the actual data
                If Len(str_val_EVENT_data) = 0 Then
                    str_val_EVENT_ALL_data = "ERROR: No Event value"
                    'str_val_EVENT_ALL_data = Empty
                ElseIf Len(str_val_EVENT_data) > 0 Then
                    str_val_EVENT_ALL_data = str_val_EVENT_data
                Else
                    str_val_EVENT_ALL_data = "ERROR"
                End If
    
            '''Col S BT (%B%) COL M Days From BT, This counts from the _
                date of the TB/BT Trial it is not part of the Baseline Switching _
                This will inclued any of the pre-BT Inspections like Shock, MIST or IN
                If Len(str_val_TRIAL_ID_data) > 0 Then
                'Has a Trial ID
                    If InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) > 0 Then
                    'Has a matching Event
                        var_val_BT_DaysFrom_data = _
                            TC_Column_Populate( _
                            str_val_TRIAL_ID_data, _
                            "B", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            str_val_EVENT_data, _
                            " TB BT IN SHK MIST ", _
                            dt_val_BT_Trial_Date_data, _
                            dt_val_DATE_CLOSED_data, _
                            "ByDate" _
                            )
                    ElseIf InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) = 0 Then
                    'The record was not written durring or before BT
                        var_val_BT_DaysFrom_data = "NULL"
                    Else
                    'ERROR in the Event field
                        var_val_BT_DaysFrom_data = "ERROR: The Event field has bad data or data type"
                    End If
                ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                'This is to account for splits that had the Trial ID _
                emptied and were written in the BT event or before
                    If InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) > 0 Then
                    'Has a matching Event
                        var_val_BT_DaysFrom_data = _
                            TC_Column_Populate( _
                            str_val_TRIAL_ID_data, _
                            "B", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            str_val_EVENT_data, _
                            " TB BT IN SHK MIST ", _
                            dt_val_BT_Trial_Date_data, _
                            dt_val_DATE_CLOSED_data, _
                            "ByDate" _
                            )
                    ElseIf InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were not written in the BT event or before
                        var_val_BT_DaysFrom_data = "NULL"
                    Else
                    'ERROR in the Event field
                        var_val_BT_DaysFrom_data = "ERROR: The Event field has bad data or data type"
                    End If
                Else
                'This is an untrapped ERROR
                    var_val_BT_DaysFrom_data = "ERROR: TRIAL_ID field has bad data or data type"
                End If
                
            '''Col T BT (%B%) COL M EVENT, This counts the TB/BT Trial it is _
                not part of the Baseline Switching This will inclued any of the _
                pre-BT Inspections like Shock, MIST or IN
                If Len(str_val_TRIAL_ID_data) > 0 Then
                'Has a Trial ID
                    If InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) > 0 Then
                    'Has a matching Event
                        str_val_BT_EVENT_data = _
                            TC_Column_Populate( _
                            str_val_TRIAL_ID_data, _
                            "B", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            str_val_EVENT_data, _
                            " TB BT IN SHK MIST ", _
                            dt_val_BT_Trial_Date_data, _
                            dt_val_DATE_CLOSED_data, _
                            "ByEvent" _
                            )
                    ElseIf InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) = 0 Then
                    'The record was not written durring or before BT
                        str_val_BT_EVENT_data = "NULL"
                    Else
                    'ERROR in the Event field
                        str_val_BT_EVENT_data = "ERROR: The Event field has bad data or data type"
                    End If
                ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                'This is to account for splits that had the Trial ID _
                emptied and were written in the BT event or before
                    If InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) > 0 Then
                    'Has a matching Event
                        str_val_BT_EVENT_data = _
                            TC_Column_Populate( _
                            str_val_TRIAL_ID_data, _
                            "B", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            "N/A", _
                            str_val_EVENT_data, _
                            " TB BT IN SHK MIST ", _
                            dt_val_BT_Trial_Date_data, _
                            dt_val_DATE_CLOSED_data, _
                            "ByEvent" _
                            )
                    ElseIf InStr(" TB BT IN SHK MIST ", str_val_EVENT_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were not written in the BT event or before
                        str_val_BT_EVENT_data = "NULL"
                    Else
                    'ERROR in the Event field
                        str_val_BT_EVENT_data = "ERROR: The Event field has bad data or data type"
                    End If
                Else
                'This is an untrapped ERROR
                    str_val_BT_EVENT_data = "ERROR: TRIAL_ID field has bad data or data type"
                End If
                
            '''Col U Concat ALL Event / Date Closed, This counts the TB/BT Trial it is _
                not part of the Baseline Switching
                'TODO: This may need to become part of the Baseline Switching _
                because Column AF has the INSURV only count
                If dt_val_DATE_CLOSED_data = Empty Then
                    'str_val_conEVENT_DATE_CLOSED_ALL_data = "NULL"
                    'This would be an OPEN card
                    str_val_conEVENT_DATE_CLOSED_ALL_data = "OPEN"
                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                    'str_val_conEVENT_DATE_CLOSED_ALL_data = "NULL"
                    'This would be an OPEN card
                    str_val_conEVENT_DATE_CLOSED_ALL_data = "OPEN"
                ElseIf dt_val_DATE_CLOSED_data > #12:00:00 AM# Then
                    str_val_conEVENT_DATE_CLOSED_ALL_data = _
                        str_val_EVENT_data & var_val_BT_DaysFrom_ALL_data
                Else
                    str_val_conEVENT_DATE_CLOSED_ALL_data = "ERROR"
                End If
    
            '''Col V Concat ALL STAR / PRI / SAFETY, This counts the TB/BT Trial it is _
                not part of the Baseline Switching
                'TODO: This may need to become part of the Baseline Switching _
                because Column AG has the INSURV only count
                str_val_conEVENT_STAR_PRI_SAFETY_ALL_data = _
                    str_val_EVENT_data & ";" & str_val_STAR_data & str_val_PRI_data & str_val_SAFE_data
    
        '''INSURV Values
            '''Col X INSURV (%C%F%S) COL K DATE CLOSED Days From AT This is _
                not part of the switch of baseline Trial event date of AT or BT
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                'Baseline_Switch_Date = dt_val_BT_Trial_Date_data
' Commented out inorder to count the BT event rollovers and re-identified
'                    If str_val_BT_EVENT_data <> "NULL" Then
'                    'The current record is already counted in the BT Column
'                        var_val_AT_DaysFrom_INSURV_data = "NULL"
'                    ElseIf str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_INSURV_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "F", _
                                        "S", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", _
                                        dt_val_AT_Trial_Date_data, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                var_val_AT_DaysFrom_INSURV_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_INSURV_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_INSURV_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "F", _
                                        "S", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " AT TC SAIL FCT FCT2 TF ", _
                                        dt_val_AT_Trial_Date_data, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                var_val_AT_DaysFrom_INSURV_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_INSURV_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'ERROR in the TRIAL_ID field
                            var_val_AT_DaysFrom_INSURV_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
'                    Else
'                    'ERROR in the Event field
'                        var_val_AT_DaysFrom_INSURV_data = "ERROR: The Event field has bad data or data type, Frist use of..."
'                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    'Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If Len(str_val_TRIAL_ID_data) > 0 Then
                    'Has a Trial ID
                        If InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            var_val_AT_DaysFrom_INSURV_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "F", _
                                    "S", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", _
                                    dt_val_AT_Trial_Date_data, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByDate" _
                                    )
                        ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                        'The record was not written in one of the tracked Events
                            var_val_AT_DaysFrom_INSURV_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_INSURV_data = "ERROR: The Event field has bad data or data type"
                        End If
                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were written in the INSURV Events
                        If InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            var_val_AT_DaysFrom_INSURV_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "F", _
                                    "S", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " AT TC SAIL FCT FCT2 TF ", _
                                    dt_val_AT_Trial_Date_data, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByDate" _
                                    )
                        ElseIf InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were not written in the INSURV Events
                            var_val_AT_DaysFrom_INSURV_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_INSURV_data = "ERROR: The Event field has bad data or data type"
                        End If
                    Else
                    'ERROR in the TRIAL_ID field
                        var_val_AT_DaysFrom_INSURV_data = "ERROR: TRIAL_ID field has bad data or data type"
                    End If
                Else
                    var_val_AT_DaysFrom_INSURV_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "var_val_AT_DaysFrom_INSURV_data : " & var_val_AT_DaysFrom_INSURV_data
    
            '''Col Y INSURV (%C%F%S) COL M EVENT, INSURV EVENTS this is _
                not part of the switch of baseline Trial event date of AT or BT
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                'Baseline_Switch_Date = dt_val_BT_Trial_Date_data
' Commented out inorder to count the BT event rollovers and re-identified
'                    If str_val_BT_EVENT_data <> "NULL" Then
'                    'The current record is already counted in the BT Column
'                        var_val_AT_DaysFrom_INSURV_data = "NULL"
'                    ElseIf str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_INSURV_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "F", _
                                        "S", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", _
                                        dt_val_AT_Trial_Date_data, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                str_val_EVENT_INSURV_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_INSURV_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_INSURV_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "F", _
                                        "S", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " AT TC SAIL FCT FCT2 TF ", _
                                        dt_val_AT_Trial_Date_data, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                str_val_EVENT_INSURV_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_INSURV_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'This is an untrapped ERROR
                            str_val_EVENT_INSURV_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
'                    Else
'                    'ERROR in the Event field
'                        str_val_EVENT_INSURV_data = "ERROR: The Event field has bad data or data type, First use of..."
'                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    'Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If Len(str_val_TRIAL_ID_data) > 0 Then
                    'Has a Trial ID
                        If InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            str_val_EVENT_INSURV_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "F", _
                                    "S", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", _
                                    dt_val_AT_Trial_Date_data, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByEvent" _
                                    )
                        ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                        'The record was not written in one of the tracked Events
                            str_val_EVENT_INSURV_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_INSURV_data = "ERROR: The Event field has bad data or data type"
                        End If
                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were written in the INSURV Events
                        If InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            str_val_EVENT_INSURV_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "F", _
                                    "S", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " AT TC SAIL FCT FCT2 TF ", _
                                    dt_val_AT_Trial_Date_data, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByEvent" _
                                    )
                        ElseIf InStr(" AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were not written in the INSURV Events
                            str_val_EVENT_INSURV_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_INSURV_data = "ERROR: The Event field has bad data or data type"
                        End If
                    Else
                    'ERROR in the TRIAL_ID field
                        str_val_EVENT_INSURV_data = "ERROR: TRIAL_ID field has bad data or data type"
                    End If
                Else
                    str_val_EVENT_INSURV_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "str_val_EVENT_INSURV_data : " & str_val_EVENT_INSURV_data
    
            '''Col Z INSURV (AT/TC %C%) COL M EVENT, Cards written or Re-ID during AT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                'Set the Baseline Date
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                    If str_val_BT_EVENT_data <> "NULL" Then
                    'The current record is already counted in the BT Column
                        str_val_EVENT_AT_ONLY_data = "NULL"
                    ElseIf str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_AT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                str_val_EVENT_AT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" AT TC ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_AT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " AT TC ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" AT TC ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                str_val_EVENT_AT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'This is an untrapped ERROR
                            str_val_EVENT_AT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
                    Else
                    'ERROR in the Event field
                        str_val_EVENT_AT_ONLY_data = "ERROR: The Event field has bad data or data type, Frist use of..."
                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If Len(str_val_TRIAL_ID_data) > 0 Then
                    'Has a Trial ID
                        If InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            str_val_EVENT_AT_ONLY_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " TB BT IN SHK MIST AT TC ", _
                                    Baseline_Switch_Date, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByEvent" _
                                    )
                        ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                        'The record was not written in one of the tracked Events
                            str_val_EVENT_AT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                        End If
                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were written in the INSURV Events
                        If InStr(" AT TC ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            str_val_EVENT_AT_ONLY_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " AT TC ", _
                                    Baseline_Switch_Date, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByEvent" _
                                    )
                        ElseIf InStr(" AT TC ", str_val_EVENT_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were not written in the INSURV Events
                            str_val_EVENT_AT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                        End If
                    Else
                    'ERROR in the TRIAL_ID field
                        str_val_EVENT_AT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                    End If
                Else
                    str_val_EVENT_AT_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "str_val_EVENT_AT_ONLY_data : " & str_val_EVENT_AT_ONLY_data
    
            '''Col AA INSURV (AT) COL X Days From AT or BT, Cards written or Re-ID during AT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                'Set the Baseline Date
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                    If str_val_BT_EVENT_data <> "NULL" Then
                    'The current record is already counted in the BT Column
                        var_val_AT_DaysFrom_AT_ONLY_data = "NULL"
                    ElseIf str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_AT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                var_val_AT_DaysFrom_AT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" AT TC ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_AT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "C", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " AT TC ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" AT TC ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                var_val_AT_DaysFrom_AT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'This is an untrapped ERROR
                            var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
                    Else
                    'ERROR in the Event field
                        var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: The Event field has bad data or data type, Frist use of..."
                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If Len(str_val_TRIAL_ID_data) > 0 Then
                    'Has a Trial ID
                        If InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            var_val_AT_DaysFrom_AT_ONLY_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " TB BT IN SHK MIST AT TC ", _
                                    Baseline_Switch_Date, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByDate" _
                                    )
                        ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                        'The record was not written in one of the tracked Events
                            var_val_AT_DaysFrom_AT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                        End If
                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                    'This is to account for splits that had the Trial ID _
                    emptied and were written in the INSURV Events
                        If InStr(" AT TC ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            var_val_AT_DaysFrom_AT_ONLY_data = _
                                TC_Column_Populate( _
                                    str_val_TRIAL_ID_data, _
                                    "C", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    "N/A", _
                                    str_val_EVENT_data, _
                                    " AT TC ", _
                                    Baseline_Switch_Date, _
                                    dt_val_DATE_CLOSED_data, _
                                    "ByDate" _
                                    )
                        ElseIf InStr(" AT TC ", str_val_EVENT_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were not written in the INSURV Events
                            var_val_AT_DaysFrom_AT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: The Event field has bad data or data type"
                        End If
                    Else
                    'ERROR in the TRIAL_ID field
                        var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                    End If
                Else
                    var_val_AT_DaysFrom_AT_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "var_val_AT_DaysFrom_AT_ONLY_data : " & var_val_AT_DaysFrom_AT_ONLY_data
    
            '''Col AB INSURV (FCT) COL M EVENT, Cards written or Re-ID during FCT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                'Set the Baseline Date
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                    If str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If str_val_EVENT_AT_ONLY_data = "NULL" Then
                        'The current record is not already counted in the AT Column
                            If Len(str_val_TRIAL_ID_data) > 0 Then
                            'Has a Trial ID
                                If InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) > 0 Then
                                'Has a matching Event
                                    str_val_EVENT_FCT_ONLY_data = _
                                        TC_Column_Populate( _
                                            str_val_TRIAL_ID_data, _
                                            "F", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            str_val_EVENT_data, _
                                            " TB BT IN SHK MIST AT TC SAIL FCT ", _
                                            Baseline_Switch_Date, _
                                            dt_val_DATE_CLOSED_data, _
                                            "ByEvent" _
                                            )
                                ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) = 0 Then
                                'The record was not written in one of the tracked Events
                                    str_val_EVENT_FCT_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    str_val_EVENT_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                                End If
                            ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were written in the INSURV Events
                                If InStr(" FCT ", str_val_EVENT_data) > 0 Then
                                'Has a matching Event
                                    str_val_EVENT_FCT_ONLY_data = _
                                        TC_Column_Populate( _
                                            str_val_TRIAL_ID_data, _
                                            "F", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            str_val_EVENT_data, _
                                            " FCT ", _
                                            Baseline_Switch_Date, _
                                            dt_val_DATE_CLOSED_data, _
                                            "ByEvent" _
                                            )
                                ElseIf InStr(" FCT ", str_val_EVENT_data) = 0 Then
                                'This is to account for splits that had the Trial ID _
                                emptied and were not written in the INSURV Events
                                    str_val_EVENT_FCT_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    str_val_EVENT_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                                End If
                            Else
                            'This is an untrapped ERROR
                                str_val_EVENT_FCT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                            End If
                        ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                        'The current record is already counted in the AT Column
                            str_val_EVENT_FCT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_FCT_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                        End If
                    ElseIf str_val_BT_EVENT_data <> "NULL" Then
                    'The current record is already counted in the BT Column
                        str_val_EVENT_FCT_ONLY_data = "NULL"
                    Else
                    'ERROR in the Event field
                        str_val_EVENT_FCT_ONLY_data = "ERROR: The BT Event field has bad data or data type, Frist use of..."
                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If str_val_EVENT_AT_ONLY_data = "NULL" Then
                        'The current record is not already counted in the AT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_FCT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "F", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC SAIL FCT ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                str_val_EVENT_FCT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" FCT ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_EVENT_FCT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "F", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " FCT ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByEvent" _
                                        )
                            ElseIf InStr(" FCT ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                str_val_EVENT_FCT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'ERROR in the TRIAL_ID field
                            str_val_EVENT_FCT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
                    ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                    'The current record is already counted in the AT Column
                        str_val_EVENT_FCT_ONLY_data = "NULL"
                    Else
                    'ERROR in the Event field
                        str_val_EVENT_FCT_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                    End If
                Else
                    str_val_EVENT_FCT_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "str_val_EVENT_FCT_ONLY_data Get A Value: " & str_val_EVENT_FCT_ONLY_data
    
            '''Col AC INSURV (FCT) COL X Days From AT, Cards written or Re-ID during AT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                'Set the Baseline Date
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                    If str_val_BT_EVENT_data = "NULL" Then
                    'The current record is not already counted in the BT Column
                        If str_val_EVENT_AT_ONLY_data = "NULL" Then
                        'The current record is not already counted in the AT Column
                            If Len(str_val_TRIAL_ID_data) > 0 Then
                            'Has a Trial ID
                                If InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) > 0 Then
                                'Has a matching Event
                                    var_val_AT_DaysFrom_FCT_ONLY_data = _
                                        TC_Column_Populate( _
                                            str_val_TRIAL_ID_data, _
                                            "F", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            str_val_EVENT_data, _
                                            " TB BT IN SHK MIST AT TC SAIL FCT ", _
                                            Baseline_Switch_Date, _
                                            dt_val_DATE_CLOSED_data, _
                                            "ByDate" _
                                            )
                                ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) = 0 Then
                                'The record was not written in one of the tracked Events
                                    var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                                End If
                            ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were written in the INSURV Events
                                If InStr(" FCT ", str_val_EVENT_data) > 0 Then
                                'Has a matching Event
                                    var_val_AT_DaysFrom_FCT_ONLY_data = _
                                        TC_Column_Populate( _
                                            str_val_TRIAL_ID_data, _
                                            "F", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            "N/A", _
                                            str_val_EVENT_data, _
                                            " FCT ", _
                                            Baseline_Switch_Date, _
                                            dt_val_DATE_CLOSED_data, _
                                            "ByDate" _
                                            )
                                ElseIf InStr(" FCT ", str_val_EVENT_data) = 0 Then
                                'This is to account for splits that had the Trial ID _
                                emptied and were not written in the INSURV Events
                                    var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                                End If
                            Else
                            'This is an untrapped ERROR
                                var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                            End If
                        ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                        'The current record is already counted in the AT Column
                            var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                        End If
                    ElseIf str_val_BT_EVENT_data <> "NULL" Then
                    'The current record is already counted in the BT Column
                        var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                    Else
                    'ERROR in the Event field
                        var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The BT Event field has bad data or data type, Frist use of..."
                    End If
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                    If str_val_EVENT_AT_ONLY_data = "NULL" Then
                        'The current record is not already counted in the AT Column
                        If Len(str_val_TRIAL_ID_data) > 0 Then
                        'Has a Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_FCT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "F", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " TB BT IN SHK MIST AT TC SAIL FCT ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT ", str_val_EVENT_data) = 0 Then
                            'The record was not written in one of the tracked Events
                                var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" FCT ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                var_val_AT_DaysFrom_FCT_ONLY_data = _
                                    TC_Column_Populate( _
                                        str_val_TRIAL_ID_data, _
                                        "F", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        "N/A", _
                                        str_val_EVENT_data, _
                                        " FCT ", _
                                        Baseline_Switch_Date, _
                                        dt_val_DATE_CLOSED_data, _
                                        "ByDate" _
                                        )
                            ElseIf InStr(" FCT ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The Event field has bad data or data type"
                            End If
                        Else
                        'ERROR in the TRIAL_ID field
                            var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                        End If
                    ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                    'The current record is already counted in the AT Column
                        var_val_AT_DaysFrom_FCT_ONLY_data = "NULL"
                    Else
                    'ERROR in the Event field
                        var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                    End If
                Else
                    var_val_AT_DaysFrom_FCT_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                'Debug.Print "var_val_AT_DaysFrom_FCT_ONLY_data Get A Value: " & var_val_AT_DaysFrom_FCT_ONLY_data
     
            '''Col AD INSURV (Expansion) COL X Days From AT or BT, Cards written or Re-ID during AT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                If str_val_Expansion_Event_data = "SAIL" _
                Or str_val_Expansion_Event_data = "TF" _
                Or str_val_Expansion_Event_data = "FCT2" Then
                'If value is equal to "SAIL" or "TF" or "FCT2" then this Column will be counted
                    If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    'Set the Baseline Date
                        Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                        If str_val_BT_EVENT_data = "NULL" Then
                        'The current record is not already counted in the BT Column
                            If str_val_EVENT_AT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the AT Column
                                If str_val_EVENT_FCT_ONLY_data = "NULL" Then
                                'The current record is not already counted in the FCT Column
                                    If Len(str_val_TRIAL_ID_data) > 0 Then
                                    'Has a Trial ID
                                        If InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) > 0 Then
                                        'Has a matching Event
                                            str_val_EVENT_EXPANTION_ONLY_data = _
                                                TC_Column_Populate( _
                                                    str_val_TRIAL_ID_data, _
                                                    "F", _
                                                    "S", _
                                                    "N/A", _
                                                    "N/A", _
                                                    "N/A", _
                                                    str_val_EVENT_data, _
                                                    " TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", _
                                                    Baseline_Switch_Date, _
                                                    dt_val_DATE_CLOSED_data, _
                                                    "ByEvent" _
                                                    )
                                        ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) = 0 Then
                                        'The record was not written in one of the tracked Events
                                            str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                                        Else
                                        'ERROR in the Event field
                                            str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                        End If
                                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                                    'This is to account for splits that had the Trial ID _
                                    emptied and were written in the INSURV Events
                                        If InStr(" SAIL ", str_val_EVENT_data) > 0 _
                                        Or InStr(" TF ", str_val_EVENT_data) > 0 _
                                        Or InStr(" FCT2 ", str_val_EVENT_data) > 0 Then
                                        'Has a matching Event
                                            str_val_EVENT_EXPANTION_ONLY_data = _
                                                TC_Column_Populate( _
                                                    str_val_TRIAL_ID_data, _
                                                    "F", _
                                                    "S", _
                                                    "N/A", _
                                                    "N/A", _
                                                    "N/A", _
                                                    str_val_EVENT_data, _
                                                    str_val_Expansion_Event_data, _
                                                    Baseline_Switch_Date, _
                                                    dt_val_DATE_CLOSED_data, _
                                                    "ByEvent" _
                                                    )
                                        ElseIf InStr(" SAIL ", str_val_EVENT_data) = 0 _
                                        Or InStr(" TF ", str_val_EVENT_data) = 0 _
                                        Or InStr(" FCT2 ", str_val_EVENT_data) = 0 Then
                                        'This is to account for splits that had the Trial ID _
                                        emptied and were not written in the INSURV Events
                                            str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                                        Else
                                        'ERROR in the Event field
                                            str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                        End If
                                    Else
                                    'This is an untrapped ERROR
                                        str_val_EVENT_EXPANTION_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                                    End If
                                ElseIf str_val_EVENT_FCT_ONLY_data <> "NULL" Then
                                'The current record is not already counted in the FCT Column
                                str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The FCT Event field has bad data or data type, Frist use of..."
                                End If
                            ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                            'The current record is already counted in the AT Column
                                str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                            End If
                        ElseIf str_val_BT_EVENT_data <> "NULL" Then
                        'The current record is already counted in the BT Column
                            str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The BT Event field has bad data or data type, Frist use of..."
                        End If
                    ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                        Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                        If str_val_EVENT_AT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the AT Column
                            If str_val_EVENT_FCT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the FCT Column
                                If Len(str_val_TRIAL_ID_data) > 0 Then
                                'Has a Trial ID
                                    If InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) > 0 Then
                                    'Has a matching Event
                                        str_val_EVENT_EXPANTION_ONLY_data = _
                                            TC_Column_Populate( _
                                                str_val_TRIAL_ID_data, _
                                                "F", _
                                                "s", _
                                                "N/A", _
                                                "N/A", _
                                                "N/A", _
                                                str_val_EVENT_data, _
                                                " TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", _
                                                Baseline_Switch_Date, _
                                                dt_val_DATE_CLOSED_data, _
                                                "ByEvent" _
                                                )
                                    ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) = 0 Then
                                    'The record was not written in one of the tracked Events
                                        str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                                    Else
                                    'ERROR in the Event field
                                        str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                    End If
                                ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                                'This is to account for splits that had the Trial ID _
                                emptied and were written in the INSURV Events
                                    If InStr(" SAIL ", str_val_EVENT_data) > 0 _
                                    Or InStr(" TF ", str_val_EVENT_data) > 0 _
                                    Or InStr(" FCT2 ", str_val_EVENT_data) > 0 Then
                                    'Has a matching Event
                                        str_val_EVENT_EXPANTION_ONLY_data = _
                                            TC_Column_Populate( _
                                                str_val_TRIAL_ID_data, _
                                                "F", _
                                                "S", _
                                                "N/A", _
                                                "N/A", _
                                                "N/A", _
                                                str_val_EVENT_data, _
                                                str_val_Expansion_Event_data, _
                                                Baseline_Switch_Date, _
                                                dt_val_DATE_CLOSED_data, _
                                                "ByEvent" _
                                                )
                                    ElseIf InStr(" SAIL ", str_val_EVENT_data) = 0 _
                                    Or InStr(" TF ", str_val_EVENT_data) = 0 _
                                    Or InStr(" FCT2 ", str_val_EVENT_data) = 0 Then
                                    'This is to account for splits that had the Trial ID _
                                    emptied and were not written in the INSURV Events
                                        str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                                    Else
                                    'ERROR in the Event field
                                        str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                    End If
                                Else
                                'ERROR in the TRIAL_ID field
                                    str_val_EVENT_EXPANTION_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                                End If
                            ElseIf str_val_EVENT_FCT_ONLY_data <> "NULL" Then
                            'The current record is not already counted in the FCT Column
                            str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The FCT Event field has bad data or data type, Frist use of..."
                            End If
                        ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                        'The current record is already counted in the AT Column
                            str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            str_val_EVENT_EXPANTION_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                        End If
                    Else
                        str_val_EVENT_EXPANTION_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                    End If
                ElseIf str_val_Expansion_Event_data = "NULL" Then
                'The Expansion Events are not being used
                    str_val_EVENT_EXPANTION_ONLY_data = "NULL"
                Else
                'ERROR in the Expansion Event field
                    str_val_EVENT_EXPANTION_ONLY_data = "ERROR: Expansion value has bad data or data type in Cell AE1"
                End If
                'Debug.Print "str_val_EVENT_EXPANTION_ONLY_data Get A Value: " & str_val_EVENT_EXPANTION_ONLY_data
                
            '''Col AE INSURV (Expansion) COL M EVENT, Cards written or Re-ID during AT _
                is part of the switch of baseline Trial event date of AT or BT and checks _
                left to make sure not already counted
                If str_val_Expansion_Event_data = "SAIL" _
                Or str_val_Expansion_Event_data = "TF" _
                Or str_val_Expansion_Event_data = "FCT2" Then
                'If value is equal to "SAIL" or "TF" or "FCT2" then this Column will be counted
                    If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    'Set the Baseline Date
                        Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                        If str_val_BT_EVENT_data = "NULL" Then
                        'The current record is not already counted in the BT Column
                            If str_val_EVENT_AT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the AT Column
                                If str_val_EVENT_FCT_ONLY_data = "NULL" Then
                                'The current record is not already counted in the FCT Column
                                    If Len(str_val_TRIAL_ID_data) > 0 Then
                                    'Has a Trial ID
                                        If InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) > 0 Then
                                        'Has a matching Event
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = _
                                                TC_Column_Populate( _
                                                    str_val_TRIAL_ID_data, _
                                                    "F", _
                                                    "S", _
                                                    "N/A", _
                                                    "N/A", _
                                                    "N/A", _
                                                    str_val_EVENT_data, _
                                                    " TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", _
                                                    Baseline_Switch_Date, _
                                                    dt_val_DATE_CLOSED_data, _
                                                    "ByDate" _
                                                    )
                                        ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) = 0 Then
                                        'The record was not written in one of the tracked Events
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                                        Else
                                        'ERROR in the Event field
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                        End If
                                    ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                                    'This is to account for splits that had the Trial ID _
                                    emptied and were written in the INSURV Events
                                        If InStr(" SAIL ", str_val_EVENT_data) > 0 _
                                        Or InStr(" TF ", str_val_EVENT_data) > 0 _
                                        Or InStr(" FCT2 ", str_val_EVENT_data) > 0 Then
                                        'Has a matching Event
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = _
                                                TC_Column_Populate( _
                                                    str_val_TRIAL_ID_data, _
                                                    "F", _
                                                    "S", _
                                                    "N/A", _
                                                    "N/A", _
                                                    "N/A", _
                                                    str_val_EVENT_data, _
                                                    str_val_Expansion_Event_data, _
                                                    Baseline_Switch_Date, _
                                                    dt_val_DATE_CLOSED_data, _
                                                    "ByDate" _
                                                    )
                                        ElseIf InStr(" SAIL ", str_val_EVENT_data) = 0 _
                                        Or InStr(" TF ", str_val_EVENT_data) = 0 _
                                        Or InStr(" FCT2 ", str_val_EVENT_data) = 0 Then
                                        'This is to account for splits that had the Trial ID _
                                        emptied and were not written in the INSURV Events
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                                        Else
                                        'ERROR in the Event field
                                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                        End If
                                    Else
                                    'This is an untrapped ERROR
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                                    End If
                                ElseIf str_val_EVENT_FCT_ONLY_data <> "NULL" Then
                                'The current record is not already counted in the FCT Column
                                var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                                Else
                                'ERROR in the Event field
                                    var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The FCT Event field has bad data or data type, Frist use of..."
                                End If
                            ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                            'The current record is already counted in the AT Column
                                var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                            End If
                        ElseIf str_val_BT_EVENT_data <> "NULL" Then
                        'The current record is already counted in the BT Column
                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The BT Event field has bad data or data type, Frist use of..."
                        End If
                    ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                        Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                        If str_val_EVENT_AT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the AT Column
                            If str_val_EVENT_FCT_ONLY_data = "NULL" Then
                            'The current record is not already counted in the FCT Column
                                If Len(str_val_TRIAL_ID_data) > 0 Then
                                'Has a Trial ID
                                    If InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) > 0 Then
                                    'Has a matching Event
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = _
                                            TC_Column_Populate( _
                                                str_val_TRIAL_ID_data, _
                                                "F", _
                                                "s", _
                                                "N/A", _
                                                "N/A", _
                                                "N/A", _
                                                str_val_EVENT_data, _
                                                " TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", _
                                                Baseline_Switch_Date, _
                                                dt_val_DATE_CLOSED_data, _
                                                "ByDate" _
                                                )
                                    ElseIf InStr(" TB BT IN SHK MIST AT TC SAIL FCT TF FCT2 ", str_val_EVENT_data) = 0 Then
                                    'The record was not written in one of the tracked Events
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                                    Else
                                    'ERROR in the Event field
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                    End If
                                ElseIf Len(str_val_TRIAL_ID_data) = 0 Then
                                'This is to account for splits that had the Trial ID _
                                emptied and were written in the INSURV Events
                                    If InStr(" SAIL ", str_val_EVENT_data) > 0 _
                                    Or InStr(" TF ", str_val_EVENT_data) > 0 _
                                    Or InStr(" FCT2 ", str_val_EVENT_data) > 0 Then
                                    'Has a matching Event
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = _
                                            TC_Column_Populate( _
                                                str_val_TRIAL_ID_data, _
                                                "F", _
                                                "S", _
                                                "N/A", _
                                                "N/A", _
                                                "N/A", _
                                                str_val_EVENT_data, _
                                                str_val_Expansion_Event_data, _
                                                Baseline_Switch_Date, _
                                                dt_val_DATE_CLOSED_data, _
                                                "ByDate" _
                                                )
                                    ElseIf InStr(" SAIL ", str_val_EVENT_data) = 0 _
                                    Or InStr(" TF ", str_val_EVENT_data) = 0 _
                                    Or InStr(" FCT2 ", str_val_EVENT_data) = 0 Then
                                    'This is to account for splits that had the Trial ID _
                                    emptied and were not written in the INSURV Events
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                                    Else
                                    'ERROR in the Event field
                                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The Event field has bad data or data type"
                                    End If
                                Else
                                'ERROR in the TRIAL_ID field
                                    var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: TRIAL_ID field has bad data or data type"
                                End If
                            ElseIf str_val_EVENT_FCT_ONLY_data <> "NULL" Then
                            'The current record is not already counted in the FCT Column
                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                            Else
                            'ERROR in the Event field
                                var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The FCT Event field has bad data or data type, Frist use of..."
                            End If
                        ElseIf str_val_EVENT_AT_ONLY_data <> "NULL" Then
                        'The current record is already counted in the AT Column
                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                        Else
                        'ERROR in the Event field
                            var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: The AT Event field has bad data or data type, Frist use of..."
                        End If
                    Else
                        var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                    End If
                ElseIf str_val_Expansion_Event_data = "NULL" Then
                'The Expansion Events are not being used
                    var_val_AT_DaysFrom_EXPANTION_ONLY_data = "NULL"
                Else
                'ERROR in the Expansion Event field
                    var_val_AT_DaysFrom_EXPANTION_ONLY_data = "ERROR: Expansion value has bad data or data type in Cell AE1"
                End If
                'Debug.Print "var_val_AT_DaysFrom_EXPANTION_ONLY_data Get A Value: " & var_val_AT_DaysFrom_EXPANTION_ONLY_data
                
            '''Col AF Concat INSURV Event / Date Closed
                'TODO: This may need to become part of the Baseline Switching _
                because Column U has the ALL Event count
                'TODO: make this select all or only select the INSURV
                'Set the Baseline Date
                If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                Else
                'Currently this will get over written by the next IF Statement
                    str_val_conEVENT_DATE_CLOSED_INSURV_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                End If
                                
                If dt_val_DATE_CLOSED_data = Empty Then
                    str_val_conEVENT_DATE_CLOSED_INSURV_data = "OPEN"
                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                    str_val_conEVENT_DATE_CLOSED_INSURV_data = "OPEN"
                ElseIf dt_val_DATE_CLOSED_data > #12:00:00 AM# Then
                    str_val_conEVENT_DATE_CLOSED_INSURV_data = _
                        str_val_EVENT_data & ConvertDateToNum(dt_val_DATE_CLOSED_data)
                Else
                    str_val_conEVENT_DATE_CLOSED_INSURV_data = "ERROR"
                End If
                'Debug.Print "str_val_conEVENT_DATE_CLOSED_INSURV_data Get A Value: " & str_val_conEVENT_DATE_CLOSED_INSURV_data
                
            '''Col AG Concat INSURV STAR / PRI / SAFETY
                'TODO: This may need to become part of the Baseline Switching _
                because Column V has the ALL Event count
                'TODO: make this select all or only select the INSURV
                'Set the Baseline Date
                'If str_val_Count_ALL_or_INSURV_data = "FULL" Then
                '    Baseline_Switch_Date = dt_val_BT_Trial_Date_data
                'ElseIf str_val_Count_ALL_or_INSURV_data = "INSURV" Then
                '    Baseline_Switch_Date = dt_val_AT_Trial_Date_data
                'Else
                'Currently this will get over written by the next IF Statement
                '    str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data = "ERROR: Count ALL or INSURV value missing from Cell S2"
                'End If
                str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data = _
                    str_val_EVENT_data & ";" & str_val_STAR_data & str_val_PRI_data & str_val_SAFE_data
                'Debug.Print "str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data Get A Value: " & str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data
    
                
            '''Col AH INSURV AT New and Roll Count (manual entry)
                'str_val_EVENT_AT_ONLY_data
                If str_val_EVENT_data <> "" _
                Or str_val_EVENT_data <> Empty Then
                'The Event is not empty, needed error check
                    If InStr(1, str_val_TRIAL_ID_data, "C", vbTextCompare) > 0 Then
                    'Has a matching Trial ID
                        If InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) > 0 Then
                        'Has a matching Event
                            str_val_EVENT_INSURV_AT_ONLY_data = "AT"
                        ElseIf str_val_EVENT_data = Empty Then
                        'Is not in the AT INSURV new or rollover count
                            str_val_EVENT_INSURV_AT_ONLY_data = Empty
                        ElseIf str_val_EVENT_data = "" Then
                        'Is not in the AT INSURV new or rollover count
                            str_val_EVENT_INSURV_AT_ONLY_data = Empty
                        ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                        'Is not in the AT INSURV new or rollover count
                            str_val_EVENT_INSURV_AT_ONLY_data = Empty
                        Else
                            str_val_EVENT_INSURV_AT_ONLY_data = "ERROR"
                        End If
                    ElseIf str_val_EVENT_data = Empty Then
                    'Event is is empty, this is to catch the empty cells that would show as InStr()=0
                        str_val_EVENT_INSURV_AT_ONLY_data = Empty
                    ElseIf str_val_EVENT_data = "" Then
                    'Event is is empty, this is to catch the empty cells that would show as InStr()=0
                        str_val_EVENT_INSURV_AT_ONLY_data = Empty
                    ElseIf InStr(1, str_val_TRIAL_ID_data, "C", vbTextCompare) = 0 Then
                    'This could be a card split in the AT Event
                        If Len(str_val_TRIAL_ID_data) = 0 Then
                        'This is to account for splits that had the Trial ID _
                        emptied and were written in the INSURV Events
                            If InStr(" AT TC ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event was a split
                                str_val_EVENT_INSURV_AT_ONLY_data = "AT"
                            ElseIf str_val_EVENT_data = Empty Then
                            'Event is empty and is an ERROR, record is not in the AT INSURV new or rollover count
                                str_val_EVENT_INSURV_AT_ONLY_data = Empty
                            ElseIf InStr(" AT TC ", str_val_EVENT_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were not written in the INSURV Events
                                str_val_EVENT_INSURV_AT_ONLY_data = Empty
                            Else
                                str_val_EVENT_INSURV_AT_ONLY_data = "ERROR"
                            End If
                        Else
                        'Is not a split in the AT INSURV new or rollover count
                        str_val_EVENT_INSURV_AT_ONLY_data = Empty
                        End If
                    ElseIf str_val_TRIAL_ID_data = Empty Then
                        str_val_EVENT_INSURV_AT_ONLY_data = Empty
                    ElseIf str_val_TRIAL_ID_data = "" Then
                        str_val_EVENT_INSURV_AT_ONLY_data = Empty
                    Else
                        str_val_EVENT_INSURV_AT_ONLY_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_data = "" _
                Or str_val_EVENT_data = Empty Then
                    str_val_EVENT_INSURV_AT_ONLY_data = Empty
                Else
                    str_val_EVENT_INSURV_AT_ONLY_data = "ERROR"
                End If
                
                'Debug.Print "str_val_EVENT_INSURV_AT_ONLY_data Get A Value: " & str_val_EVENT_INSURV_AT_ONLY_data

            '''Col AI Screen Groupings (manual entry) (G & K & S)
'               str_val_SCREEN_G_K_S_data
'               str_val_SCREEN_data 'Col E SCREEN
                If InStr(1, " GA GI GN GD KA KI KN KD GF ", str_val_SCREEN_data, vbTextCompare) > 0 Then
                    If InStr(1, " GA GI GN GD ", str_val_SCREEN_data, vbTextCompare) > 0 Then
                        str_val_SCREEN_G_K_S_data = "G"
                    ElseIf InStr(1, " KA KI KN KD ", str_val_SCREEN_data, vbTextCompare) > 0 Then
                        str_val_SCREEN_G_K_S_data = "K"
                    ElseIf InStr(1, " GF ", str_val_SCREEN_data, vbTextCompare) > 0 Then
                        str_val_SCREEN_G_K_S_data = "S"
                    Else
                        str_val_SCREEN_G_K_S_data = "ERROR"
                    End If
                ElseIf InStr(1, " GA GI GN GD KA KI KN KD GF ", str_val_SCREEN_data, vbTextCompare) = 0 Then
                    str_val_SCREEN_G_K_S_data = Empty
                Else
                    str_val_SCREEN_G_K_S_data = "ERROR"
                End If

            '''Col AJ INSURV AT CONCATENATE Event & Screen Group
                'str_val_conEVENT_SCREEN_AT_ONLY_data
                'Look left to see if this is INSURV
                'Commented out for now to deconflect with FCT and Expansion Events
                If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                    str_val_conEVENT_SCREEN_AT_ONLY_data = str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                    str_val_conEVENT_SCREEN_AT_ONLY_data = Empty
                Else
                    str_val_conEVENT_SCREEN_AT_ONLY_data = "ERROR"
                End If

            '''Col AK INSURV AT Open at DEL CONCATENATE Mile & Event & Screen
                'str_val_DEL_BT_DaysFrom_AT_data 'Col AK INSURV AT Open at DEL CONCATENATE Mile & Event & Screen
                    'dt_val_DaysFrom_Baseline_to_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'dt_val_Date_of_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
                'dt_val_DaysFrom_Baseline_to_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'Look left to see if this is INSURV AT countable
                If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                    'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                    For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                    If dt_val_DATE_CLOSED_data = Empty Then
                        str_val_DEL_BT_DaysFrom_AT_data = "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        str_val_DEL_BT_DaysFrom_AT_data = "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_DEL Then
                        str_val_DEL_BT_DaysFrom_AT_data = "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_DEL Then
                        str_val_DEL_BT_DaysFrom_AT_data = "CLOSED"
                    ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        str_val_DEL_BT_DaysFrom_AT_data = "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    Else
                        str_val_DEL_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                    str_val_DEL_BT_DaysFrom_AT_data = Empty
                Else
                    str_val_DEL_BT_DaysFrom_AT_data = "ERROR"
                End If
                
            '''Col AL INSURV AT Open at SAIL CONCATENATE Mile & Event & Screen
                'str_val_SAIL_BT_DaysFrom_AT_data 'Col AL INSURV AT Open at SAIL CONCATENATE Mile & Event & Screen
                    'dt_val_DaysFrom_Baseline_to_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
                'dt_val_DaysFrom_Baseline_to_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'Look left to see if this is INSURV AT countable
                If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                    'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                    For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                    If dt_val_DATE_CLOSED_data = Empty Then
                        str_val_SAIL_BT_DaysFrom_AT_data = "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        str_val_SAIL_BT_DaysFrom_AT_data = "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_SAIL Then
                        str_val_SAIL_BT_DaysFrom_AT_data = "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_SAIL Then
                        str_val_SAIL_BT_DaysFrom_AT_data = "CLOSED"
                    ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        str_val_SAIL_BT_DaysFrom_AT_data = "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    Else
                        str_val_SAIL_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                    str_val_SAIL_BT_DaysFrom_AT_data = Empty
                Else
                    str_val_SAIL_BT_DaysFrom_AT_data = "ERROR"
                End If
                
            '''Col AM INSURV AT Open at OWLD CONCATENATE Mile & Event & Screen
                'str_val_OWLD_BT_DaysFrom_AT_data 'Col AM INSURV AT Open at OWLD CONCATENATE Mile & Event & Screen
                    'dt_val_DaysFrom_Baseline_to_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'dt_val_Date_of_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
                'dt_val_DaysFrom_Baseline_to_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff
                'Look left to see if this is INSURV AT countable
                If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                    'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                    For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                    If dt_val_DATE_CLOSED_data = Empty Then
                        str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                        str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                        str_val_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                    ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                    Else
                        str_val_OWLD_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" Then
                    If str_val_EVENT_INSURV_data = "NULL" _
                    Or str_val_EVENT_INSURV_data = "" _
                    Or str_val_EVENT_INSURV_data = Empty Then
                        str_val_OWLD_BT_DaysFrom_AT_data = Empty
                    ElseIf str_val_EVENT_INSURV_data <> "NULL" Then
                        If dt_val_DATE_CLOSED_data = Empty Then
                            str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                            str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                            str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                            str_val_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                        ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                            str_val_OWLD_BT_DaysFrom_AT_data = "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data
                        Else
                            str_val_OWLD_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    ElseIf str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                        str_val_OWLD_BT_DaysFrom_AT_data = Empty
                    Else
                        str_val_OWLD_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                    str_val_OWLD_BT_DaysFrom_AT_data = Empty
                Else
                    str_val_OWLD_BT_DaysFrom_AT_data = "ERROR"
                End If
                
            '''Col AN INSURV AT New and Roll Count HIGH PRI (1S)
'               str_val_PRI1S_AT_ONLY_data 'Col AN INSURV AT New and Roll Count HIGH PRI (1S)
                'This will count C and empty AT/TC Splits. This is only counting the New and Rolled _
                cards @ the AT Event that are not STARED CARDS.
'                    dt_val_AT_Trial_Date_data 'Date of the AT Trial
'                    dt_val_Date_of_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
'                    dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
'                    dt_val_Date_of_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
'                    str_val_STAR_data = .Offset((rowPointer), 1).Value2 'Col B STAR
'                    str_val_PRI_data = .Offset((rowPointer), 2).Value2 'Col C PRI
'                    str_val_SAFE_data = .Offset((rowPointer), 3).Value2 'Col D SAFETY
                If str_val_EVENT_INSURV_data <> "NULL" _
                Or str_val_EVENT_INSURV_data <> Empty Then
                'Look left to see if this is INSURV AT countable
                    If (str_val_PRI_data & str_val_SAFE_data) = "1S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "2S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "1" Then
                    'This last OR might need to be = "1 " vice "1"
                        If str_val_STAR_data <> "STAR" _
                        Or str_val_STAR_data <> "*" Then
                        'This only counts the High Pri cards that are not equal to STAR/*
                            If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                            'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                            For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                                If dt_val_DATE_CLOSED_data = Empty Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_AT_ONLY_data = "PRI" & ";" & "AT" & ";" & str_val_PRI_data & str_val_SAFE_data '& ";" & str_val_SCREEN_G_K_S_data
                                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_AT_ONLY_data = "PRI" & ";" & "AT" & ";" & str_val_PRI_data & str_val_SAFE_data '& ";" & str_val_SCREEN_G_K_S_data
                                ElseIf dt_val_DATE_CLOSED_data > dt_val_AT_Trial_Date_data Then
                                'Closed after the AT Trial, OPEN
                                    str_val_PRI1S_AT_ONLY_data = "PRI" & ";" & "AT" & ";" & str_val_PRI_data & str_val_SAFE_data '& ";" & str_val_SCREEN_G_K_S_data
                                ElseIf dt_val_DATE_CLOSED_data <= dt_val_AT_Trial_Date_data Then
                                'Closed at or during the AT trial, OPEN
                                'Philisoficly, this is counting all cards as OPEN for the AT count
                                    'str_val_PRI1S_AT_ONLY_data = "CLOSED"
                                    str_val_PRI1S_AT_ONLY_data = "PRI" & ";" & "AT" & ";" & str_val_PRI_data & str_val_SAFE_data '& ";" & str_val_SCREEN_G_K_S_data
                                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                                'Is OPEN
                                    str_val_PRI1S_AT_ONLY_data = "PRI" & ";" & "AT" & ";" & str_val_PRI_data & str_val_SAFE_data '& ";" & str_val_SCREEN_G_K_S_data
                                Else
                                'Bad data or data type
                                    str_val_PRI1S_AT_ONLY_data = "ERROR"
                                End If
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" _
                            Or str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                            'Not written or contained in the AT INSURV count
                                str_val_PRI1S_AT_ONLY_data = Empty
                            Else
                            'Bad data or data type
                                str_val_PRI1S_AT_ONLY_data = "ERROR"
                            End If
                        ElseIf str_val_STAR_data = "STAR" _
                        Or str_val_STAR_data = "*" Then
                        'STARED cards are counted else where and not here to prevent double counting
                            str_val_PRI1S_AT_ONLY_data = Empty
                        Else
                        'Bad data or data type
                            str_val_PRI1S_AT_ONLY_data = "ERROR"
                        End If
                    Else
                    'Was not a High Pri trial card
                        str_val_PRI1S_AT_ONLY_data = Empty
                    End If
                ElseIf str_val_EVENT_INSURV_data = Empty Then
                'Not written or contained in the AT INSURV count
                    str_val_DEL_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_DEL_BT_DaysFrom_AT_data = "ERROR"
                End If

            '''Col AO INSURV AT Open at DEL High Pri (1S)
'               str_val_PRI1S_DEL_BT_DaysFrom_AT_data 'Col AO INSURV AT Open at DEL High Pri (1S)
                'This will count C and empty AT/TC Splits. This is only counting the New and Rolled _
                cards @ the AT Event that are not STARED CARDS.
'                    dt_val_Date_of_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
                If str_val_EVENT_INSURV_data <> "NULL" _
                Or str_val_EVENT_INSURV_data <> Empty Then
                'Look left to see if this is INSURV AT countable
                    If (str_val_PRI_data & str_val_SAFE_data) = "1S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "2S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "1" Then
                    'This last OR might need to be = "1 " vice "1"
                        If str_val_STAR_data <> "STAR" _
                        Or str_val_STAR_data <> "*" Then
                        'This only counts the High Pri cards that are not equal to STAR/*
                            If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                            'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                            For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                                If dt_val_DATE_CLOSED_data = Empty Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "PRI" & ";" & "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "PRI" & ";" & "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_DEL Then
                                'Closed after the AT Trial, OPEN
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "PRI" & ";" & "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_DEL Then
                                'Closed at or during the AT trial, OPEN
                                'Philisoficly, this is counting all cards as OPEN for the AT count
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "CLOSED"
                                    'str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "PRI" & ";" & "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                                'Is OPEN
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "PRI" & ";" & "DEL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                Else
                                'Bad data or data type
                                    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "ERROR"
                                End If
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" _
                            Or str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                            'Not written or contained in the AT INSURV count
                                str_val_PRI1S_DEL_BT_DaysFrom_AT_data = Empty
                            Else
                            'Bad data or data type
                                str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_STAR_data = "STAR" _
                        Or str_val_STAR_data = "*" Then
                        'STARED cards are counted else where and not here to prevent double counting
                            str_val_PRI1S_DEL_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_PRI1S_DEL_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    Else
                    'Was not a High Pri trial card
                        str_val_PRI1S_DEL_BT_DaysFrom_AT_data = Empty
                    End If
                ElseIf str_val_EVENT_INSURV_data = Empty Then
                'Not written or contained in the AT INSURV count
                    str_val_DEL_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_DEL_BT_DaysFrom_AT_data = "ERROR"
                End If

            '''Col AP INSURV AT Open at SAIL High Pri (1S)
'               str_val_PRI1S_SAIL_BT_DaysFrom_AT_data 'Col AP INSURV AT Open at SAIL High Pri (1S)
                'This will count C and empty AT/TC Splits. This is only counting the New and Rolled _
                cards @ the AT Event that are not STARED CARDS. This could count the SAIL Inspection, but it is not.
'                    dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
                If str_val_EVENT_INSURV_data <> "NULL" _
                Or str_val_EVENT_INSURV_data <> Empty Then
                'Look left to see if this is INSURV AT countable
                    If (str_val_PRI_data & str_val_SAFE_data) = "1S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "2S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "1" Then
                    'This last OR might need to be = "1 " vice "1"
                        If str_val_STAR_data <> "STAR" _
                        Or str_val_STAR_data <> "*" Then
                        'This only counts the High Pri cards that are not equal to STAR/*
                            If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                            'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                            For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                                If dt_val_DATE_CLOSED_data = Empty Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "PRI" & ";" & "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "PRI" & ";" & "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_SAIL Then
                                'Closed after the AT Trial, OPEN
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "PRI" & ";" & "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_SAIL Then
                                'Closed at or during the AT trial, OPEN
                                'Philisoficly, this is counting all cards as OPEN for the AT count
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "CLOSED"
                                    'str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "PRI" & ";" & "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                                'Is OPEN
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "PRI" & ";" & "SAIL" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                Else
                                'Bad data or data type
                                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "ERROR"
                                End If
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" _
                            Or str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                            'Not written or contained in the AT INSURV count
                                str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty
                            Else
                            'Bad data or data type
                                str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_STAR_data = "STAR" _
                        Or str_val_STAR_data = "*" Then
                        'STARED cards are counted else where and not here to prevent double counting
                            str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    Else
                    'Was not a High Pri trial card
                        str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty
                    End If
                ElseIf str_val_EVENT_INSURV_data = Empty Then
                'Not written or contained in the AT INSURV count
                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "ERROR"
                End If

            '''Col AQ INSURV AT Open at OWLD High Pri (1S)
'               str_val_PRI1S_OWLD_BT_DaysFrom_AT_data 'Col AQ INSURV AT Open at OWLD High Pri (1S)
                'This will count C,F,S and empty AT/TC,SAIL,FCT,FCT2/TF
                'This will count C and empty AT/TC Splits. This is only counting the New and Rolled _
                cards @ the AT Event that are not STARED CARDS.
'                    str_val_EVENT_INSURV_AT_ONLY_data 'Col AH
'                    dt_val_Date_of_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
                If str_val_EVENT_INSURV_data <> "NULL" _
                Or str_val_EVENT_INSURV_data <> Empty Then
                'Look left to see if this is INSURV AT countable
                    If (str_val_PRI_data & str_val_SAFE_data) = "1S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "2S" _
                    Or (str_val_PRI_data & str_val_SAFE_data) = "1" Then
                    'This last OR might need to be = "1 " vice "1"
                        If str_val_STAR_data <> "STAR" _
                        Or str_val_STAR_data <> "*" Then
                        'This only counts the High Pri cards that are not equal to STAR/*
                            If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                            'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                            For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                                If dt_val_DATE_CLOSED_data = Empty Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                                'Closed after the AT Trial, OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                                'Closed at or during the AT trial, OPEN
                                'Philisoficly, this is counting all cards as OPEN for the AT count
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                                    'str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                                'Is OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                Else
                                'Bad data or data type
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "ERROR"
                                End If
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" _
                            Or str_val_EVENT_INSURV_AT_ONLY_data = Empty Then
                            'Not written or contained in the AT INSURV count
                                'str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = Empty
                                If dt_val_DATE_CLOSED_data = Empty Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                                'Presumed to be OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                                'Closed after the AT Trial, OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                                'Closed at or during the AT trial, OPEN
                                'Philisoficly, this is counting all cards as OPEN for the AT count
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                                    'str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                                'Is OPEN
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                                Else
                                'Bad data or data type
                                    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "ERROR"
                                End If
                            Else
                            'Bad data or data type
                                str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_STAR_data = "STAR" _
                        Or str_val_STAR_data = "*" Then
                        'STARED cards are counted else where and not here to prevent double counting
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    Else
                    'Was not a High Pri trial card
                        str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = Empty
                    End If
                ElseIf str_val_EVENT_INSURV_data = Empty Then
                'Not written or contained in the AT INSURV count
                    If var_val_AT_DaysFrom_INSURV_data <> "NULL" _
                    Or var_val_AT_DaysFrom_INSURV_data <> "AT" Then
                    'Not written or contained in the AT INSURV or a pure BT card
                        If dt_val_DATE_CLOSED_data = Empty Then
                        'Presumed to be OPEN
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                        ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        'Presumed to be OPEN
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                        ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                        'Closed after the AT Trial, OPEN
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                        ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                        'Closed at or during the AT trial, OPEN
                        'Philisoficly, this is counting all cards as OPEN for the AT count
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                            'str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                        ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        'Is OPEN
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "PRI" & ";" & "OWLD" & ";" & str_val_EVENT_INSURV_AT_ONLY_data & ";" & str_val_SCREEN_G_K_S_data & ";" & str_val_PRI_data & str_val_SAFE_data
                        Else
                        'Bad data or data type
                            str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    Else
                    'Is written or contained in the AT INSURV or a pure BT card
                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty
                    End If
                Else
                'Bad data or data type
                    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = "ERROR"
                End If


            '''Col AR INSURV FCT New and Roll Count (manual entry)
'               str_val_FCT_FCT_not_in_AT_ONLY_data 'Col AR INSURV FCT New and Roll Count (manual entry)
                If str_val_EVENT_data <> "" _
                Or str_val_EVENT_data <> Empty Then
                'The Event is not empty, needed error check
                    If str_val_EVENT_INSURV_AT_ONLY_data = "AT" Then
                        str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                    ElseIf str_val_EVENT_INSURV_AT_ONLY_data <> "AT" Then
                        If InStr(1, str_val_TRIAL_ID_data, "S", vbTextCompare) > 0 _
                        Or InStr(1, str_val_TRIAL_ID_data, "F", vbTextCompare) > 0 Then
                        'Has a matching Trial ID
                            If InStr(" TB BT IN SHK MIST AT TC SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                            'Has a matching Event
                                str_val_FCT_FCT_not_in_AT_ONLY_data = "FCT"
                            ElseIf str_val_EVENT_data = Empty Then
                            'Is not in the AT INSURV new or rollover count
                                str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                            ElseIf str_val_EVENT_data = "" Then
                            'Is not in the AT INSURV new or rollover count
                                str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                            ElseIf InStr(" TB BT IN SHK MIST AT TC ", str_val_EVENT_data) = 0 Then
                            'Is not in the AT INSURV new or rollover count
                                str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                            Else
                                str_val_FCT_FCT_not_in_AT_ONLY_data = "ERROR"
                            End If
                        ElseIf str_val_EVENT_data = Empty Then
                        'Event is is empty, this is to catch the empty cells that would show as InStr()=0
                            str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                        ElseIf str_val_EVENT_data = "" Then
                        'Event is is empty, this is to catch the empty cells that would show as InStr()=0
                            str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                        ElseIf InStr(1, str_val_TRIAL_ID_data, "S", vbTextCompare) = 0 _
                        Or InStr(1, str_val_TRIAL_ID_data, "F", vbTextCompare) = 0 Then
                        'This could be a card split in the AT Event
                            If Len(str_val_TRIAL_ID_data) = 0 Then
                            'This is to account for splits that had the Trial ID _
                            emptied and were written in the INSURV Events
                                If InStr(" SAIL FCT FCT2 TF ", str_val_EVENT_data) > 0 Then
                                'Has a matching Event was a split
                                    str_val_FCT_FCT_not_in_AT_ONLY_data = "FCT"
                                ElseIf str_val_EVENT_data = Empty Then
                                'Event is empty and is an ERROR, record is not in the AT INSURV new or rollover count
                                    str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                                ElseIf InStr(" SAIL FCT FCT2 TF ", str_val_EVENT_data) = 0 Then
                                'This is to account for splits that had the Trial ID _
                                emptied and were not written in the INSURV Events
                                    str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                                Else
                                    str_val_FCT_FCT_not_in_AT_ONLY_data = "ERROR"
                                End If
                            Else
                            'Is not a split in the AT INSURV new or rollover count
                            str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                            End If
                        ElseIf str_val_TRIAL_ID_data = Empty Then
                            str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                        ElseIf str_val_TRIAL_ID_data = "" Then
                            str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                        Else
                            str_val_FCT_FCT_not_in_AT_ONLY_data = "ERROR"
                        End If
                    Else
                        str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                    End If
                ElseIf str_val_EVENT_data = "" _
                Or str_val_EVENT_data = Empty Then
                    str_val_FCT_FCT_not_in_AT_ONLY_data = Empty
                Else
                    str_val_FCT_FCT_not_in_AT_ONLY_data = "ERROR"
                End If

            '''Col AS INSURV FCT Not Counted In AT CONCATENATE Event & Screen
'               str_val_FCT_OWLD_not_in_AT_ONLY_data 'Col AS INSURV FCT Not Counted In AT CONCATENATE Event & Screen
                If str_val_FCT_FCT_not_in_AT_ONLY_data <> "" _
                Or str_val_FCT_FCT_not_in_AT_ONLY_data <> Empty Then
                'The Event is not empty, needed error check
                    If str_val_FCT_FCT_not_in_AT_ONLY_data = "FCT" Then
                        str_val_FCT_OWLD_not_in_AT_ONLY_data = "FCT" & ";" & str_val_SCREEN_G_K_S_data
                    Else
                        str_val_FCT_OWLD_not_in_AT_ONLY_data = Empty
                    End If
                Else
                    str_val_FCT_OWLD_not_in_AT_ONLY_data = Empty
                End If
                
            '''Col AT INSURV FCT Not Counted In AT CONCATENATE Event & Screen
'               str_val_FCT_OWLD_not_in_AT_ONLY_data 'Col AS INSURV FCT Not Counted In AT CONCATENATE Event & Screen
                If str_val_FCT_FCT_not_in_AT_ONLY_data <> "" _
                Or str_val_FCT_FCT_not_in_AT_ONLY_data <> Empty Then
                'The Event is not empty, needed error check
                    If str_val_FCT_FCT_not_in_AT_ONLY_data = "FCT" Then
                    'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                    For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                        If dt_val_DATE_CLOSED_data = Empty Then
                        'Presumed to be OPEN
                            str_val_FCT_OWLD_BT_DaysFrom_AT_data = "FCT" & ";" & "OWLD" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        'Presumed to be OPEN
                            str_val_FCT_OWLD_BT_DaysFrom_AT_data = "FCT" & ";" & "OWLD" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                        'Closed after the AT Trial, OPEN
                            str_val_FCT_OWLD_BT_DaysFrom_AT_data = "FCT" & ";" & "OWLD" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                        'Closed at or during the AT trial, OPEN
                        'Philisoficly, this is counting all cards as OPEN for the AT count
                            str_val_FCT_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                            'str_val_FCT_OWLD_BT_DaysFrom_AT_data = "FCT" & ";" & "OWLD" & ";" & str_val_SCREEN_G_K_S_data
                        'ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        'Is OPEN
                        '    str_val_FCT_OWLD_BT_DaysFrom_AT_data = "FCT" & ";" & "OWLD" & ";" & str_val_SCREEN_G_K_S_data
                        Else
                        'Bad data or data type
                            str_val_FCT_OWLD_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    Else
                        str_val_FCT_OWLD_BT_DaysFrom_AT_data = Empty
                    End If
                Else
                    str_val_FCT_OWLD_BT_DaysFrom_AT_data = Empty
                End If
    
            '''Col AU INSURV AT NEW STAR
                'str_val_STAR_NEW_BT_DaysFrom_AT_data 'Col AU INSURV AT NEW STAR
'               str_val_EVENT_data = .Offset((rowPointer), 12).Value2 'Col M EVENT
                'This will count C and empty AT/TC Splits. This is only counting the New and Rolled _
                cards @ the AT Event that are STARED CARDS.
'                    dt_val_AT_Trial_Date_data 'Date of the AT Trial
'                    dt_val_Date_of_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
'                    dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
'                    dt_val_Date_of_OWLD 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AM:2
'                    str_val_STAR_data = .Offset((rowPointer), 1).Value2 'Col B STAR
                If str_val_STAR_data = "STAR" _
                Or str_val_STAR_data = "*" Then
                'This only counts the cards that are equal to STAR/*
                    If str_val_EVENT_data = "AT" _
                    Or str_val_EVENT_data = "TC" Then
                    'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                    For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                        If dt_val_DATE_CLOSED_data = Empty Then
                        'Presumed to be OPEN
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "STAR" & ";" & "AT" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                        'Presumed to be OPEN
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "STAR" & ";" & "AT" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data > dt_val_AT_Trial_Date_data Then
                        'Closed after the AT Trial, OPEN
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "STAR" & ";" & "AT" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf dt_val_DATE_CLOSED_data <= dt_val_AT_Trial_Date_data Then
                        'Closed at or during the AT trial, OPEN
                        'Philisoficly, this is counting all cards as OPEN for the AT count
                            'str_val_STAR_NEW_BT_DaysFrom_AT_data = "CLOSED"
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "STAR" & ";" & "AT" & ";" & str_val_SCREEN_G_K_S_data
                        ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                        'Is OPEN
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "STAR" & ";" & "AT" & ";" & str_val_SCREEN_G_K_S_data
                        Else
                        'Bad data or data type
                            str_val_STAR_NEW_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    ElseIf str_val_EVENT_data <> "AT" _
                    Or str_val_EVENT_data <> "TC" _
                    Or str_val_EVENT_data = Empty Then
                    'Not written or contained in the AT INSURV count
                        str_val_STAR_NEW_BT_DaysFrom_AT_data = Empty
                    Else
                    'Bad data or data type
                        str_val_STAR_NEW_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_STAR_data <> "STAR" _
                Or str_val_STAR_data <> "*" Then
                'IF NOT A STARED cards DONT COUNT
                    str_val_STAR_NEW_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_STAR_NEW_BT_DaysFrom_AT_data = "ERROR"
                End If

            '''Col AV INSURV AT Open at DEL STAR
            'str_val_STAR_DEL_BT_DaysFrom_AT_data 'Col AV INSURV AT Open at DEL STAR
'               str_val_EVENT_data = .Offset((rowPointer), 12).Value2 'Col M EVENT
                'This will count C and empty AT/TC Splits. This is only counting the Open _
                cards @ the DEL Event that are STARED CARDS.
'                    dt_val_Date_of_DEL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AK:2
                If str_val_EVENT_data <> "NULL" _
                Or str_val_EVENT_data <> Empty Then
                    If str_val_STAR_data = "STAR" _
                    Or str_val_STAR_data = "*" Then
                    'This only counts the cards that are equal to STAR/*
                        If str_val_EVENT_data = "AT" _
                        Or str_val_EVENT_data = "TC" Then
                        'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                        For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                            If dt_val_DATE_CLOSED_data = Empty Then
                            'Presumed to be OPEN
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "STAR" & ";" & "DEL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                            'Presumed to be OPEN
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "STAR" & ";" & "DEL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_DEL Then
                            'Closed after the AT Trial, OPEN
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "STAR" & ";" & "DEL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_DEL Then
                            'Closed at or during the AT trial, OPEN
                            'Philisoficly, this is counting all cards as OPEN for the AT count
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "CLOSED"
                                'str_val_STAR_DEL_BT_DaysFrom_AT_data = "STAR" & ";" & "DEL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                            'Is OPEN
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "STAR" & ";" & "DEL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            Else
                            'Bad data or data type
                                str_val_STAR_DEL_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_EVENT_data <> "AT" _
                        Or str_val_EVENT_data <> "TC" _
                        Or str_val_EVENT_data = Empty Then
                        'Not written or contained in the AT INSURV count
                            str_val_STAR_DEL_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_STAR_DEL_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    ElseIf str_val_STAR_data <> "STAR" _
                    Or str_val_STAR_data <> "*" Then
                    'IF NOT A STARED cards DONT COUNT
                        str_val_STAR_DEL_BT_DaysFrom_AT_data = Empty
                    Else
                    'Bad data or data type
                        str_val_STAR_DEL_BT_DaysFrom_AT_data = "ERROR"
                    End If
                Else
                'Event is Empty
                    str_val_STAR_DEL_BT_DaysFrom_AT_data = Empty
                End If

            '''Col AW INSURV AT Open at SAIL STAR
            'str_val_STAR_SAIL_BT_DaysFrom_AT_data 'Col AW INSURV AT Open at SAIL STAR
'               str_val_EVENT_data = .Offset((rowPointer), 12).Value2 'Col M EVENT
                'This will count C and empty AT/TC Splits. This is only counting the Open _
                cards @ the SAIL Event that are STARED CARDS. This also counts the SAIL Inspection.
'                    dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
                If str_val_EVENT_data <> "NULL" _
                Or str_val_EVENT_data <> Empty Then
                    If str_val_STAR_data = "STAR" _
                    Or str_val_STAR_data = "*" Then
                    'This only counts the cards that are equal to STAR/*
                        If str_val_EVENT_data = "AT" _
                        Or str_val_EVENT_data = "TC" _
                        Or str_val_EVENT_data = "SAIL" Then
                        'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                        For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                            If dt_val_DATE_CLOSED_data = Empty Then
                            'Presumed to be OPEN
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "STAR" & ";" & "SAIL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                            'Presumed to be OPEN
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "STAR" & ";" & "SAIL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_SAIL Then
                            'Closed after the AT Trial, OPEN
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "STAR" & ";" & "SAIL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_SAIL Then
                            'Closed at or during the AT trial, OPEN
                            'Philisoficly, this is counting all cards as OPEN for the AT count
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "CLOSED"
                                'str_val_STAR_SAIL_BT_DaysFrom_AT_data = "STAR" & ";" & "SAIL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                            'Is OPEN
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "STAR" & ";" & "SAIL" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            Else
                            'Bad data or data type
                                str_val_STAR_SAIL_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_EVENT_data <> "AT" _
                        Or str_val_EVENT_data <> "TC" _
                        Or str_val_EVENT_data <> "SAIL" _
                        Or str_val_EVENT_data = Empty Then
                        'Not written or contained in the AT INSURV count
                            str_val_STAR_SAIL_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_STAR_SAIL_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    ElseIf str_val_STAR_data <> "STAR" _
                    Or str_val_STAR_data <> "*" Then
                    'STARED cards are counted else where and not here to prevent double counting
                        str_val_STAR_SAIL_BT_DaysFrom_AT_data = Empty
                    Else
                    'Bad data or data type
                        str_val_STAR_SAIL_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_data = Empty Then
                'Not written or contained in the AT INSURV count
                    str_val_STAR_SAIL_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_STAR_SAIL_BT_DaysFrom_AT_data = "ERROR"
                End If

            '''Col AX INSURV AT Open at OWLD STAR
            'str_val_STAR_OWLD_BT_DaysFrom_AT_data 'Col AX INSURV AT Open at OWLD STAR
'               str_val_EVENT_data = .Offset((rowPointer), 12).Value2 'Col M EVENT
                'This will count C and empty AT/TC Splits. This is only counting the Open _
                cards @ the SAIL Event that are STARED CARDS. This also counts the SAIL Inspection.
'                    dt_val_Date_of_SAIL 'This is the number of days from the baseline date of the BT or AT Trial, is used for date diff Col:Rw AL:2
                If str_val_EVENT_data <> "NULL" _
                Or str_val_EVENT_data <> Empty Then
                    If str_val_STAR_data = "STAR" _
                    Or str_val_STAR_data = "*" Then
                    'This only counts the cards that are equal to STAR/*
                        If str_val_EVENT_data = "AT" _
                        Or str_val_EVENT_data = "TC" _
                        Or str_val_EVENT_data = "SAIL" _
                        Or str_val_EVENT_data = "FCT" Then
                        'For DEL or SAIL, this will either show <DEL/SAIL>;EVENT;SCREEN or "CLOSED" _
                        For OWLD this either show "OPEN" or <DEL/SAIL>;EVENT;SCREEN
                            If dt_val_DATE_CLOSED_data = Empty Then
                            'Presumed to be OPEN
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "STAR" & ";" & "OWLD" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data = #12:00:00 AM# Then
                            'Presumed to be OPEN
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "STAR" & ";" & "OWLD" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data > dt_val_Date_of_OWLD Then
                            'Closed after the AT Trial, OPEN
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "STAR" & ";" & "OWLD" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf dt_val_DATE_CLOSED_data <= dt_val_Date_of_OWLD Then
                            'Closed at or during the AT trial, OPEN
                            'Philisoficly, this is counting all cards as OPEN for the AT count
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "CLOSED"
                                'str_val_STAR_OWLD_BT_DaysFrom_AT_data = "STAR" & ";" & "OWLD" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            ElseIf str_val_EVENT_INSURV_AT_ONLY_data = "OPEN" Then
                            'Is OPEN
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "STAR" & ";" & "OWLD" & ";" & str_val_EVENT_data & ";" & str_val_SCREEN_G_K_S_data
                            Else
                            'Bad data or data type
                                str_val_STAR_OWLD_BT_DaysFrom_AT_data = "ERROR"
                            End If
                        ElseIf str_val_EVENT_data <> "AT" _
                        Or str_val_EVENT_data <> "TC" _
                        Or str_val_EVENT_data <> "SAIL" _
                        Or str_val_EVENT_data <> "FCT" _
                        Or str_val_EVENT_data = Empty Then
                        'Not written or contained in the AT INSURV count
                            str_val_STAR_OWLD_BT_DaysFrom_AT_data = Empty
                        Else
                        'Bad data or data type
                            str_val_STAR_OWLD_BT_DaysFrom_AT_data = "ERROR"
                        End If
                    ElseIf str_val_STAR_data <> "STAR" _
                    Or str_val_STAR_data <> "*" Then
                    'STARED cards are counted else where and not here to prevent double counting
                        str_val_STAR_OWLD_BT_DaysFrom_AT_data = Empty
                    Else
                    'Bad data or data type
                        str_val_STAR_OWLD_BT_DaysFrom_AT_data = "ERROR"
                    End If
                ElseIf str_val_EVENT_data = Empty Then
                'Not written or contained in the AT INSURV count
                    str_val_STAR_OWLD_BT_DaysFrom_AT_data = Empty
                Else
                'Bad data or data type
                    str_val_STAR_OWLD_BT_DaysFrom_AT_data = "ERROR"
                End If
        
        '''Find out if there is any missing key values
            'TODO: If there is missing a value in the EVENT, the whole record should be skipped
            'ElseIF Len(str_val_EVENT_data) = 0
            'The Values remain as EMPTY or BLANK
            'Else
            'Error in the descrimination feild data
            'Need a hard exit, WEND
            
        '''Write computed values
            '''Col O TRIM DATE DISCOVERED
            If dt_val_DATE_DISC_data = Empty Then
                .Offset(rowPointer, 14).Value = "ERROR" '= Empty 'Col O TRIM DATE DISCOVERED
            Else
                .Offset(rowPointer, 14).Value = dt_val_DATE_DISC_data 'Col O TRIM DATE DISCOVERED
            End If
            
            '''Col P TRIM DATE CLOSED
            If dt_val_DATE_CLOSED_data = Empty Then
                .Offset(rowPointer, 15).Value = "OPEN" '= Empty 'Col P TRIM DATE CLOSED
            Else
                .Offset(rowPointer, 15).Value = dt_val_DATE_CLOSED_data 'Col P TRIM DATE CLOSED
            End If

            '''Col Q ALL COL M Days From BT
            .Offset(rowPointer, 16).Value = var_val_BT_DaysFrom_ALL_data 'Col Q ALL COL M Days From BT

            '''Col R ALL COL M EVENT
            .Offset(rowPointer, 17).Value = str_val_EVENT_ALL_data 'Col R ALL COL M EVENT
            
            '''Col S BT (%B%) COL M Days From BT
            .Offset(rowPointer, 18).Value = var_val_BT_DaysFrom_data 'Col S BT (%B%) COL M Days From BT
            
            '''Col T BT (%B%) COL M EVENT
            .Offset(rowPointer, 19).Value = str_val_BT_EVENT_data 'Col T BT (%B%) COL M EVENT

            '''Col U Concat ALL Event / Date Closed
            .Offset(rowPointer, 20).Value = str_val_conEVENT_DATE_CLOSED_ALL_data 'Col U Concat ALL Event / Date Closed

            '''Col V Concat ALL STAR / PRI / SAFETY
            .Offset(rowPointer, 21).Value = str_val_conEVENT_STAR_PRI_SAFETY_ALL_data 'Col V Concat ALL STAR / PRI / SAFETY

            '''Col W IS BLANK
            '.Offset(rowPointer, 22).Value = Empty 'Col W IS BLANK
            
            '''Col X INSURV (%C%F%S) COL M EVENT Days From AT
            .Offset(rowPointer, 23).Value = var_val_AT_DaysFrom_INSURV_data  'Col X INSURV (%C%F%S) COL M EVENT Days From AT
            
            '''Col Y INSURV (%C%F%S) COL M EVENT
            .Offset(rowPointer, 24).Value = str_val_EVENT_INSURV_data 'Col Y INSURV (%C%F%S) COL M EVENT

            '''Col Z INSURV (AT/TC %C%) COL M EVENT
            .Offset(rowPointer, 25).Value = str_val_EVENT_AT_ONLY_data 'Col Z INSURV (AT/TC %C%) COL M EVENT
            
            '''Col AA INSURV (AT) COL X Days From AT
            .Offset(rowPointer, 26).Value = var_val_AT_DaysFrom_AT_ONLY_data 'Col AA INSURV (AT) COL X Days From AT
            
            '''Col AB INSURV (FCT) COL M EVENT
            .Offset(rowPointer, 27).Value = str_val_EVENT_FCT_ONLY_data 'Col AB INSURV (FCT) COL M EVENT
            
            '''Col AC INSURV (FCT) COL X Days From AT
            .Offset(rowPointer, 28).Value = var_val_AT_DaysFrom_FCT_ONLY_data 'Col AC INSURV (FCT) COL X Days From AT
            
            '''Col AD INSURV (SAIL) COL M EVENT
            .Offset(rowPointer, 29).Value = str_val_EVENT_EXPANTION_ONLY_data 'Col AD INSURV (SAIL) COL M EVENT
            
            '''Col AE INSURV (SAIL) COL W Days AT
            .Offset(rowPointer, 30).Value = var_val_AT_DaysFrom_EXPANTION_ONLY_data 'Col AE INSURV (SAIL) COL W Days AT

            '''Col AF Concat INSURV Event / Date Closed
            .Offset(rowPointer, 31).Value = str_val_conEVENT_DATE_CLOSED_INSURV_data 'Col AF Concat INSURV Event / Date Closed

            '''Col AG Concat INSURV STAR / PRI / SAFETY
            .Offset(rowPointer, 32).Value = str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data 'Col AG Concat INSURV STAR / PRI / SAFETY

            '''Col AH INSURV AT New and Roll Count (manual entry)
            .Offset(rowPointer, 33).Value = str_val_EVENT_INSURV_AT_ONLY_data 'Col AH INSURV AT New and Roll Count (manual entry)

            '''Col AI Screen Groupings (manual entry) (G & K & S)
            .Offset(rowPointer, 34).Value = str_val_SCREEN_G_K_S_data 'Col AI Screen Groupings (manual entry) (G & K & S)

            '''Col AJ INSURV AT CONCATENATE Event & Screen Group
            .Offset(rowPointer, 35).Value = str_val_conEVENT_SCREEN_AT_ONLY_data 'Col AJ INSURV AT CONCATENATE Event & Screen Group

            '''Col AK INSURV AT Open at DEL CONCATENATE Mile & Event & Screen
            .Offset(rowPointer, 36).Value = str_val_DEL_BT_DaysFrom_AT_data 'Col AK INSURV AT Open at DEL CONCATENATE Mile & Event & Screen

            '''Col AL INSURV AT Open at SAIL CONCATENATE Mile & Event & Screen
            .Offset(rowPointer, 37).Value = str_val_SAIL_BT_DaysFrom_AT_data 'Col AL INSURV AT Open at SAIL CONCATENATE Mile & Event & Screen

            '''Col AM INSURV AT Open at OWLD CONCATENATE Mile & Event & Screen
            .Offset(rowPointer, 38).Value = str_val_OWLD_BT_DaysFrom_AT_data 'Col AM INSURV AT Open at OWLD CONCATENATE Mile & Event & Screen

            '''Col AN INSURV AT New and Roll Count HIGH PRI (1S)
            .Offset(rowPointer, 39).Value = str_val_PRI1S_AT_ONLY_data 'Col AN INSURV AT New and Roll Count HIGH PRI (1S)

            '''Col AO INSURV AT Open at DEL High Pri (1S)
            .Offset(rowPointer, 40).Value = str_val_PRI1S_DEL_BT_DaysFrom_AT_data 'Col AO INSURV AT Open at DEL High Pri (1S)

            '''Col AP INSURV AT Open at SAIL High Pri (1S)
            .Offset(rowPointer, 41).Value = str_val_PRI1S_SAIL_BT_DaysFrom_AT_data 'Col AP INSURV AT Open at SAIL High Pri (1S)

            '''Col AQ INSURV AT Open at OWLD High Pri (1S)
            .Offset(rowPointer, 42).Value = str_val_PRI1S_OWLD_BT_DaysFrom_AT_data 'Col AQ INSURV AT Open at OWLD High Pri (1S)

            '''Col AR INSURV FCT New and Roll Count (manual entry)
            .Offset(rowPointer, 43).Value = str_val_FCT_FCT_not_in_AT_ONLY_data 'Col AR INSURV FCT New and Roll Count (manual entry)

            '''Col AS INSURV FCT Not Counted In AT CONCATENATE Event & Screen
            .Offset(rowPointer, 44).Value = str_val_FCT_OWLD_not_in_AT_ONLY_data 'Col AS INSURV FCT Not Counted In AT CONCATENATE Event & Screen

            '''Col AT INSURV FCT Open at OWLD
            .Offset(rowPointer, 45).Value = str_val_FCT_OWLD_BT_DaysFrom_AT_data 'Col AT INSURV FCT Open at OWLD
            
            '''Col AU INSURV AT NEW STAR
            .Offset(rowPointer, 46).Value = str_val_STAR_NEW_BT_DaysFrom_AT_data 'Col AU INSURV AT NEW STAR

            '''Col AV INSURV AT Open at DEL STAR
            .Offset(rowPointer, 47).Value = str_val_STAR_DEL_BT_DaysFrom_AT_data 'Col AV INSURV AT Open at DEL STAR

            '''Col AW INSURV AT Open at SAIL STAR
            .Offset(rowPointer, 48).Value = str_val_STAR_SAIL_BT_DaysFrom_AT_data 'Col AW INSURV AT Open at SAIL STAR

            '''Col AX INSURV AT Open at OWLD STAR
            .Offset(rowPointer, 49).Value = str_val_STAR_OWLD_BT_DaysFrom_AT_data 'Col AX INSURV AT Open at OWLD STAR
        
        End With
    
    '''Empty the varible values
    'These will not change while updating the workbook
        'Dim dt_val_AT_Trial_Date_data As Date 'This is the date of the AT Trial, is used for date diff Col:Rw X:2
        'Dim Baseline_Switch_Date As Date 'This uses the BT or AT Trial Date and is Picked based on FULL count ALL or INSURV count INSURV only Col:Rw S:2
        'Dim str_val_Count_ALL_or_INSURV_data As String 'This switches the TC counts from FULL count ALL or INSURV count INSURV only Col:Rw S:2
        'Dim str_val_Expansion_Event_data As String 'This is for additional Trial Events, is used with SAIL, FCT2, FT, NULL Col:Rw AE:1
    'These are the Data columns varibles
    str_val_DSP_data = Empty 'As String 'Col A DSP
    str_val_STAR_data = Empty 'As String 'Col B STAR
    str_val_PRI_data = Empty 'As String 'Col C PRI
    str_val_SAFE_data = Empty 'As String 'Col D SAFETY
    str_val_SCREEN_data = Empty 'As String 'Col E SCREEN
    str_val_ACT1_data = Empty 'As String 'Col F ACTION CODE 1
    str_val_ACT2_data = Empty 'As String 'Col G ACTION CODE 2
    str_val_STATUS_data = Empty 'As String 'Col H STATUS
    str_val_ACT_TKN_data = Empty 'As String 'Col I ACTION TAKEN
    var_val_DATE_DISC_data = Empty 'As String 'Col J DATE DISCOVERED
    var_val_DATE_CLOSED_data = Empty 'As String 'Col K DATE CLOSED
    str_val_TRIAL_ID_data = Empty 'As String 'Col L TRIAL ID
    str_val_EVENT_data = Empty 'As String 'Col M EVENT
    'These are the computed new values
    'moved into another function: temp_DATE_DISC_str = Empty ' As String 'temp value while converting from Oracle Date to MS Date
    dt_val_DATE_DISC_data = Empty ' As Variant 'As Date
    'moved into another function: temp_DATE_CLOSED_str = Empty ' As String 'temp value while converting from Oracle Date to MS Date
    dt_val_DATE_CLOSED_data = Empty ' As Variant 'As Date
    var_val_BT_DaysFrom_ALL_data = Empty ' As Variant
    str_val_EVENT_ALL_data = Empty ' As String
    var_val_BT_DaysFrom_data = Empty ' As Variant
    str_val_BT_EVENT_data = Empty '  As String
    str_val_conEVENT_STAR_PRI_SAFETY_ALL_data = Empty ' As String
    str_val_conEVENT_DATE_CLOSED_ALL_data = Empty ' As String
    var_val_AT_DaysFrom_INSURV_data = Empty ' As Variant
    str_val_EVENT_INSURV_data = Empty ' As String
    str_val_EVENT_AT_ONLY_data = Empty ' As String
    var_val_AT_DaysFrom_AT_ONLY_data = Empty ' As Variant
    str_val_EVENT_FCT_ONLY_data = Empty ' As String
    var_val_AT_DaysFrom_FCT_ONLY_data = Empty ' As Variant
    str_val_EVENT_EXPANTION_ONLY_data = Empty ' As String
    var_val_AT_DaysFrom_EXPANTION_ONLY_data = Empty ' As Variant
    str_val_conEVENT_DATE_CLOSED_INSURV_data = Empty ' As String
    str_val_conEVENT_STAR_PRI_SAFETY_INSURV_data = Empty ' As String
    'Dim str_val_EVENT_AT_ONLY_data As String
    str_val_SCREEN_G_K_S_data = Empty ' As String
    str_val_conEVENT_SCREEN_AT_ONLY_data = Empty ' As String
    str_val_DEL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_SAIL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_OWLD_BT_DaysFrom_AT_data = Empty ' As String
    str_val_PRI1S_AT_ONLY_data = Empty ' As String
    str_val_PRI1S_DEL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_PRI1S_SAIL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_PRI1S_OWLD_BT_DaysFrom_AT_data = Empty ' As String
    str_val_FCT_FCT_not_in_AT_ONLY_data = Empty ' As String
    str_val_FCT_OWLD_not_in_AT_ONLY_data = Empty ' As String
    str_val_FCT_OWLD_BT_DaysFrom_AT_data = Empty 'As String
    str_val_STAR_NEW_BT_DaysFrom_AT_data = Empty ' As String
    str_val_STAR_DEL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_STAR_SAIL_BT_DaysFrom_AT_data = Empty ' As String
    str_val_STAR_OWLD_BT_DaysFrom_AT_data = Empty ' As String
        
    Next
       
'        'Format the Current Row
'        With Workbooks(exportWorkBook).Worksheets(exportWkSt).Range("A1:AR1")
'            .Offset(rowPointer).Font.Name = "Arial"
'            .Offset(rowPointer).Font.Size = 10
'            .Offset(rowPointer).Font.FontStyle = "Bold"
'            .Offset(rowPointer).HorizontalAlignment = xlCenter
'            .Offset(rowPointer).VerticalAlignment = xlTop
'            .Offset(rowPointer).WrapText = True
'            .Offset(rowPointer).NumberFormat = "@" 'Col A:L Format as text
'        End With
'        'Format the Dates
'        Workbooks(exportWorkBook).Worksheets(exportWkSt).Range("A1").Offset(rowPointer).NumberFormat = "mm/dd/yyyy;@"
'        Workbooks(exportWorkBook).Worksheets(exportWkSt).Range("AH1").Offset(rowPointer).NumberFormat = "mm/dd/yyyy;@"
        
'        expUsrDwnLd.MoveNext
'        rowPointer = rowPointer + 1
'    Wend
    
    '~~Debug.Print (vbTab & "-OWE_woH-end of oneWay_DataSheetExport_woHDR()")
mySuccess = True
'Return the Boolean to the calling sub
ReFresh_WorkSheet_Data = mySuccess

End Function


