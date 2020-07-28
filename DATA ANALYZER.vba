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
    
    Range("###").Value = "###"
    Range("###").Value = "###"

End Sub

Sub InputBoxPractice()

    Dim traDate As Date
    Dim setDate As Date
   
    traDate = InputBox("###")
    setDate = InputBox("###")
   
    Dim UsedRng As Range, LastRow As Long
    Set UsedRng = ActiveSheet.UsedRange

    LastRow = UsedRng(UsedRng.Cells.Count).Row
    Range("###").Select
    
    Do Until ActiveCell.Row = LastRow + 1
      
                If traDate > ActiveCell Then
                    ActiveCell.Offset(0, 6).Value = "###"
                End If
                   
                If traDate <= ActiveCell And ActiveCell <= setDate Then
                    
                    If ActiveCell.Offset(0, -2) = "I" Then
                            If ActiveCell.Offset(0, -1) = "###" Then
                                 ActiveCell.Offset(0, 6).Value = "###"
                            ElseIf ActiveCell.Offset(0, -1) = "###" Then
                                 ActiveCell.Offset(0, 6).Value = "###"
                            Else: ActiveCell.Offset(0, 6).Value = "###"
                            End If
                    
                       ElseIf ActiveCell.Offset(0, -2) = "###" Then
                            If ActiveCell.Offset(0, -1) = "###" Then
                                 ActiveCell.Offset(0, 6).Value = "###"
                            ElseIf ActiveCell.Offset(0, -1) = "###" Then
                                 ActiveCell.Offset(0, 6).Value = "###"
                            Else: ActiveCell.Offset(0, 6).Value = "###"
                            End If
                        
                       ElseIf ActiveCell.Offset(0, -2).Value = "N" Then
                       ActiveCell.Offset(0, 6).Value = "###"
                            
                       ElseIf ActiveCell.Offset(0, -2) = "###" Then
                       ActiveCell.Offset(0, 6).Value = "###"
                    End If
                        
                End If
            
                If traDate < ActiveCell And ActiveCell >= setDate Then
                    ActiveCell.Offset(0, 6).Value = "###"
                End If

    
 
                ActiveCell.Offset(1, 0).Select
    
    Loop
   
End Sub

Sub outgoingInput()

    Dim setDate As Date
   
    setDate = InputBox("###")
   
    Dim UsedRng As Range, LastRow As Long
    Set UsedRng = ActiveSheet.UsedRange
    

    LastRow = UsedRng(UsedRng.Cells.Count).Row
    Range("###").Select
    
    Do Until ActiveCell.Row = LastRow + 1
      
                If setDate >= ActiveCell And ActiveCell <> "###" Then
                    ActiveCell.Offset(0, 6).Value = ""
                    
                    ElseIf ActiveCell > setDate And ActiveCell <> "###" Then
                    ActiveCell.Offset(0, 6).Value = ""
                
                End If
                
                If ActiveCell = "###" Then
                    If ActiveCell = "###" Then ActiveCell.Offset(0, -1).Activate
                        If setDate >= ActiveCell Then
                        ActiveCell.Offset(0, 7).Value = ""
                        ActiveCell.Offset(0, 1).Select
                        ElseIf ActiveCell > setDate Then
                        ActiveCell.Offset(0, 7).Value = ""
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
    Range("###").Select
    
        Do Until ActiveCell.Row = LastRow
    
                If ActiveCell = "###" Then
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












