Sub RisksAndIssues()

Dim riskRow As Integer
Dim issueRow As Integer

riskRow = 11
issueRow = 15


For i = 2 To 200
    Title = Sheet20.Cells(i, 2)
    RiskOrIssue = Sheet20.Cells(i, 3)
    Status = Sheet20.Cells(i, 4)
    Rating = Sheet20.Cells(i, 5)
    Mitigation = Sheet20.Cells(i, 11)
    EscalationLevel = Sheet20.Cells(i, 14)
    ProgramsAffected = Sheet20.Cells(i, 15)
    
    
    If RiskOrIssue = "Risk" And Rating = "16" And _
    Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If riskRow <= 13 Then
            Sheet18.Cells(riskRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            riskRow = riskRow + 1
           End If
              
    ElseIf RiskOrIssue = "Risk" And Rating = "12" And _
    Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If riskRow <= 13 Then
            Sheet18.Cells(riskRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            riskRow = riskRow + 1
            
        End If
        
    ElseIf RiskOrIssue = "Risk" And Rating = "9" And _
    Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
  
        If riskRow <= 13 Then
            Sheet18.Cells(riskRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            riskRow = riskRow + 1
            
        End If
        
     ElseIf RiskOrIssue = "Risk" And Rating = "8" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
     ElseIf RiskOrIssue = "Risk" And Rating = "6" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
     ElseIf RiskOrIssue = "Risk" And Rating = "4" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
     ElseIf RiskOrIssue = "Risk" And Rating = "3" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
     ElseIf RiskOrIssue = "Risk" And Rating = "2" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
     ElseIf RiskOrIssue = "Risk" And Rating = "1" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
    
          If riskRow <= 13 Then
              Sheet18.Cells(riskRow, 2) = Title
              Sheet18.Cells(riskRow, 10) = Mitigation
              riskRow = riskRow + 1
              
          End If
          
'   Below starts sequence for Issues
    
    ElseIf RiskOrIssue = "Issue" And Rating = "16" And _
    Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(issueRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "12" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(issueRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "9" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "8" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "6" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "4" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "3" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "2" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
        
     ElseIf RiskOrIssue = "Issue" And Rating = "1" And _
     Status = "Active" And InStr(UCase(ProgramsAffected), UCase(Range("N5"))) Then
      
        If issueRow <= 17 Then
            Sheet18.Cells(issueRow, 2) = Title
            Sheet18.Cells(riskRow, 10) = Mitigation
            issueRow = issueRow + 1
        End If
    
    
    End If

Next i


End Sub