Attribute VB_Name = "mAGLBrokerReport"
Option Explicit

Sub AGLBrokerReport()
    
    On Error GoTo Err_AGLBrokerReport
    
    Dim lngRow As Long
    Dim rTargetRow As Range
    Dim strDateVal As String
    
        
    Application.Cursor = xlWait
    
    strDateVal = InputBox(Prompt:="Please Enter the End of 30 Days Date in the format dd-mm-yyyy", _
          Title:="End of 30 days date", Default:=Date)

    
    Set rTargetRow = rOutput.Rows(1)
    
    For lngRow = 2 To rInput.Rows.Count Step 1
        Application.StatusBar = "Function 2 of 3: AGL - Broker Transpose; processing row " & lngRow & " of " & rInput.Rows.Count
        Set rTargetRow = rTargetRow.Offset(1)
        
        If lngRow > 2 Then
            rTargetRow.FillDown
            rTargetRow.ClearContents
        End If
            
        With rInput.Rows(lngRow)
        
            If DateDiff("d", CDate(Left(.Cells(75), 2) & "-" & Mid(.Cells(75), 4, 2) & "-" & Right(.Cells(75), 4)), strDateVal) < 31 Then
                rTargetRow.Cells(1) = "'" & .Cells(11)                                         'Lead ID
                rTargetRow.Cells(2) = MMslashDDslashYYYY(Date)                                 'Customer Cancellation Date
                rTargetRow.Cells(3) = fMain.tbxSourceFilename                                  'Customer Cancellation Reason
                rTargetRow.Cells(4) = "Yes"                                                    'Retention?
                rTargetRow.Cells(5) = "Clawback - Retention"                                   'Status
                'rTargetRow.Cells(6) = .Cells(75)
                'rTargetRow.Cells(7) = strDateVal
                'rTargetRow.Cells(8) = DateDiff("d", CDate(Left(.Cells(75), 2) & "-" & Mid(.Cells(75), 4, 2) & "-" & Right(.Cells(75), 4)), strDateVal)
            Else
                rTargetRow.Cells(1) = "'" & .Cells(11)                                          'Lead ID
                'rTargetRow.Cells(2) = MMslashDDslashYYYY(Date)                                 'Customer Cancellation Date
                rTargetRow.Cells(3) = fMain.tbxSourceFilename                                   'Customer Cancellation Reason
                'rTargetRow.Cells(4) = "Yes"                                                    'Retention?
                rTargetRow.Cells(5) = "Balance Paid"                               'Status
                'rTargetRow.Cells(6) = .Cells(75)
                'rTargetRow.Cells(7) = strDateVal
                'rTargetRow.Cells(8) = DateDiff("d", CDate(Left(.Cells(75), 2) & "-" & Mid(.Cells(75), 4, 2) & "-" & Right(.Cells(75), 4)), strDateVal)
            End If

                        
        End With
    Next lngRow
    
    Set rTargetRow = Nothing
    Set rOutput = rOutput.CurrentRegion

Exit_AGLBrokerReport:
    Exit Sub
    
Err_AGLBrokerReport:
    Err.Raise Number:=Err.Number, Description:=Err.Description
    Resume Exit_AGLBrokerReport

End Sub
