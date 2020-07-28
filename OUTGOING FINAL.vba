Sub remove_all_space()

LRows = Range("H" & Rows.Count).End(xlUp).Row - 1

For i = 0 To LRows
      Range("h1").Offset(i, 0).Value = Application.WorksheetFunction.Trim(Range("h1").Offset(i, 0).Value)
      Range("h1").Offset(i, 0).Value = Application.WorksheetFunction.Substitute(Range("h1").Offset(i, 0).Value, " ", "")
Next i

End Sub

Sub Conversions()

    Columns(5).EntireColumn.Delete
    Columns(12).EntireColumn.Delete
    Columns(15).EntireColumn.Delete
    Columns(2).EntireColumn.Delete
    
    Range("N4").Value = "MOVE ASSET"
    Range("O4").Value = "PROCESSING RESPONSIBILITY"

End Sub


Sub sort()

 Dim ws As Worksheet
 Set ws = ThisWorkbook.Worksheets(1)

 
 ws.Rows("1").EntireRow.Hidden = True
 ws.Rows("2").EntireRow.Hidden = True
 ws.Rows("3").EntireRow.Hidden = True
 ws.Rows("4").EntireRow.Hidden = True

    ws.UsedRange.Activate
    
    Selection.sort Key1:=Range("H5"), order1:=xlAscending
    
    'ws.Rows("1").EntireRow.Hidden = False
    'ws.Rows("2").EntireRow.Hidden = False
    'ws.Rows("3").EntireRow.Hidden = False
    'ws.Rows("4").EntireRow.Hidden = False
 

End Sub


Sub InputBoxPractice()

    Dim traDate As Date
    Dim setDate As Date
   
    traDate = InputBox("Please type in Trade Date as sentence (Eg: January 5 2020)", "Trade Date")
    setDate = InputBox("Please type in Settlement Date as sentence (Eg: January 5 2020)", "Settlement Date")
   
    Dim UsedRng As Range, LastRow As Long
    Set UsedRng = ActiveSheet.UsedRange

    LastRow = UsedRng(UsedRng.Cells.Count).Row
    Range("H5").Select
    
    Do Until ActiveCell.Row = LastRow + 1
      
                If traDate > ActiveCell Then
                    ActiveCell.Offset(0, 6).Value = "Expired"
                End If
                   
                If traDate <= ActiveCell And ActiveCell <= setDate Then
                    
                    If ActiveCell.Offset(0, -2) = "I" Then
                            If ActiveCell.Offset(0, -1) = "DVCA" Then
                                 ActiveCell.Offset(0, 6).Value = "YES"
                            ElseIf ActiveCell.Offset(0, -1) = "DRIP" Then
                                 ActiveCell.Offset(0, 6).Value = "YES"
                            Else: ActiveCell.Offset(0, 6).Value = "No"
                            End If
                    
                       ElseIf ActiveCell.Offset(0, -2) = "V" Then
                            If ActiveCell.Offset(0, -1) = "DVCA" Then
                                 ActiveCell.Offset(0, 6).Value = "YES"
                            ElseIf ActiveCell.Offset(0, -1) = "DRIP" Then
                                 ActiveCell.Offset(0, 6).Value = "YES"
                            Else: ActiveCell.Offset(0, 6).Value = "No"
                            End If
                        
                       ElseIf ActiveCell.Offset(0, -2).Value = "N" Then
                       ActiveCell.Offset(0, 6).Value = "Yes"
                            
                       ElseIf ActiveCell.Offset(0, -2) = "D" Then
                       ActiveCell.Offset(0, 6).Value = "Yes"
                    End If
                        
                End If
            
                If traDate < ActiveCell And ActiveCell >= setDate Then
                    ActiveCell.Offset(0, 6).Value = "YES"
                End If

    
 
                ActiveCell.Offset(1, 0).Select
    
    Loop
   
End Sub

Sub outgoingInput()

    Dim setDate As Date
   
    setDate = InputBox("Please type in Settlement Date as sentence (Eg: January 5 2020)", "Settlement Date")
   
    Dim UsedRng As Range, LastRow As Long
    Set UsedRng = ActiveSheet.UsedRange
    

    LastRow = UsedRng(UsedRng.Cells.Count).Row
    Range("I5").Select
    
    Do Until ActiveCell.Row = LastRow + 1
      
                If setDate >= ActiveCell And ActiveCell <> "00/00/00" Then
                    ActiveCell.Offset(0, 6).Value = "OLD CUSTODIAN SSB"
                    
                    ElseIf ActiveCell > setDate And ActiveCell <> "00/00/00" Then
                    ActiveCell.Offset(0, 6).Value = "NEW CUSTODIAN"
                
                End If
                
                If ActiveCell = "00/00/00" Then
                    If ActiveCell = "00/00/00" Then ActiveCell.Offset(0, -1).Activate
                        If setDate >= ActiveCell Then
                        ActiveCell.Offset(0, 7).Value = "OLD CUSTODIAN SSB"
                        ActiveCell.Offset(0, 1).Select
                        ElseIf ActiveCell > setDate Then
                        ActiveCell.Offset(0, 7).Value = "NEW CUSTODIAN"
                        ActiveCell.Offset(0, 1).Select
                        End If
                End If
                
                ActiveCell.Offset(1, 0).Select
    Loop
   
End Sub

Sub highlight()

    Dim UsedRng As Range, LastRow As Long
    Set UsedRng = ActiveSheet.UsedRange
    
   

    LastRow = UsedRng(UsedRng.Cells.Count).Row
    Range("N5").Select
    
        Do Until ActiveCell.Row = LastRow
    
                If ActiveCell = "No" Then
                    ActiveCell.EntireRow.Interior.ColorIndex = 8
                End If
                
                ActiveCell.Offset(1, 0).Select
        
        Loop
    
End Sub

Sub Unhide_ColumnsRows_On_All_Sheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
            ws.Cells.EntireColumn.Hidden = False
            ws.Cells.EntireRow.Hidden = False
    Next ws
    
    Worksheets(1).Columns("A:N").AutoFit
End Sub












