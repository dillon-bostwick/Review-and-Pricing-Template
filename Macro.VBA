' Tile Review and Upgrade Pricing Macros
' 
' I created these macros when working at Brightwater Homes Inc.
' These macros were created specifically use with the xlsm included,
' "Tile Review and Upgrade Pricing"
'
' By Dillon Bostwick for Brightwater Homes
' May 2013
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Calculate cost of each tile option
Sub getcost()

    lastcomprw = Sheets("Tile Selections").Cells(1, 1).End(xlDown).Row
    
    ssrw = 5

    Do Until Cells(ssrw, 1) = "TOTAL"
        If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) And Not IsEmpty(Cells(ssrw, 16)) And Not IsEmpty(Cells(ssrw, 17)) And Not IsEmpty(Cells(ssrw, 18)) And Not IsEmpty(Cells(ssrw, 19)) And Not IsEmpty(Cells(ssrw, 20)) Then
    
            For csrw = 2 To lastcomprw
        
                If ActiveSheet.Cells(ssrw, 18) = Sheets("Tile Selections").Cells(csrw, 5) Then
                If ActiveSheet.Cells(ssrw, 19) = Sheets("Tile Selections").Cells(csrw, 6) Then
                If ActiveSheet.Cells(ssrw, 20) = Sheets("Tile Selections").Cells(csrw, 7) Then
        
                            ActiveSheet.Cells(ssrw, 11) = CCur(Sheets("Tile Selections").Cells(csrw, 8))
                        End If
                    End If
                End If
            Next csrw
        End If
        ssrw = ssrw + 1
    Loop
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Get Type
Sub gettype()

    'get last row - const through prog
    lastcomprw = Sheets("Tile Selections").Cells(1, 1).End(xlDown).Row

    '''''''''''''

    acl = 200
    Do Until Cells(acl - 195, 1) = "TOTAL"
        If Not Cells(acl - 195, 1).Font.Bold And Not IsEmpty(Cells(acl - 195, 1)) And Not IsEmpty(Cells(acl - 195, 16)) And Not IsEmpty(Cells(acl - 195, 17)) And Not IsEmpty(Cells(acl - 195, 18)) Then
            totalin = 1
            For csrw = 2 To lastcomprw ' csrw: compiled sheet row
                exists = False
                For arw = 2 To totalin ' arw: Temp sheet row
                    If ActiveWorkbook.Sheets("Temp").Cells(arw, acl) = Sheets("Tile Selections").Cells(csrw, 6) Then
                        exists = True
                        Exit For
                    End If
                Next arw
    
                If (exists = False) And (ActiveSheet.Cells(acl - 195, 18) = Sheets("Tile Selections").Cells(csrw, 5)) Then
                    ActiveWorkbook.Sheets("Temp").Cells(totalin + 1, acl) = Sheets("Tile Selections").Cells(csrw, 6)
                    totalin = totalin + 1
                End If
            Next csrw
    
        Let ActiveWorkbook.Sheets("Temp").Cells(1, acl) = totalin ' first row in Temp stores last row for that item col
    
        End If
        acl = acl + 1
    Loop

    '''''Create Drop Down''''''''

    ssrw = 5
    
    Do Until Cells(ssrw, 1) = "TOTAL"
    If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) And Not IsEmpty(Cells(ssrw, 16)) And Not IsEmpty(Cells(ssrw, 17)) And Not IsEmpty(Cells(ssrw, 18)) Then
        Let lastrw = ActiveWorkbook.Sheets("Temp").Cells(1, ssrw + 195).Value
        If lastrw = 2 Then
            TempRangeName = Cells(2, ssrw + 195).Address
        Else
            TempRangeName = Range(Cells(2, ssrw + 195), Cells(lastrw, ssrw + 195)).Address
        End If
        With ActiveWorkbook.ActiveSheet.Cells(ssrw, 19).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Temp!" & TempRangeName
            .InCellDropdown = True
        End With
    End If
    ssrw = ssrw + 1
        Loop
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Get Dimensions
Sub getdimension()

    'get last row - const through prog
    lastcomprw = Sheets("Tile Selections").Cells(1, 1).End(xlDown).Row

    '''''''''''''
    acl = 100
    Do Until Cells(acl - 95, 1) = "TOTAL"
        If Not Cells(acl - 95).Font.Bold And Not IsEmpty(Cells(acl - 95, 1)) And Not IsEmpty(Cells(acl - 95, 16)) And Not IsEmpty(Cells(acl - 95, 17)) And Not IsEmpty(Cells(acl - 95, 18)) And Not IsEmpty(Cells(acl - 95, 19)) Then
            totalin = 1
    
            For csrw = 2 To lastcomprw ' csrw: compiled sheet row
                exists = False
                For arw = 2 To totalin ' arw: Temp sheet row
            
                    If ActiveWorkbook.Sheets("Temp").Cells(arw, acl) = Sheets("Tile Selections").Cells(csrw, 7) Then
                      exists = True
                        Exit For
                    End If
            
                Next arw
    
                If (exists = False) And (Cells(acl - 95, 18) = Sheets("Tile Selections").Cells(csrw, 5)) And (Cells(acl - 95, 19) = Sheets("Tile Selections").Cells(csrw, 6)) Then
                    ActiveWorkbook.Sheets("Temp").Cells(totalin + 1, acl) = Sheets("Tile Selections").Cells(csrw, 7)
                    totalin = totalin + 1
                End If

            Next csrw
    
            Let ActiveWorkbook.Sheets("Temp").Cells(1, acl) = totalin ' first row in Temp stores last row for that item col
    
        End If
        acl = acl + 1
    Loop

    '''''Data Validation''''''''

    ssrw = 5
    Do Until Cells(ssrw, 1) = "TOTAL"
    If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) And Not IsEmpty(Cells(ssrw, 16)) And Not IsEmpty(Cells(ssrw, 17)) And Not IsEmpty(Cells(ssrw, 18)) And Not IsEmpty(Cells(ssrw, 19)) Then
        Let lastrw = ActiveWorkbook.Sheets("Temp").Cells(1, ssrw + 95).Value
    
        If lastrw = 2 Then
            TempRangeName = Cells(2, ssrw + 95).Address
        Else
            TempRangeName = Range(Cells(2, ssrw + 95), Cells(lastrw, ssrw + 95)).Address
        End If
    
        With ActiveWorkbook.ActiveSheet.Cells(ssrw, 20).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Temp!" & TempRangeName
            .InCellDropdown = True
        End With
    End If
    ssrw = ssrw + 1
    Loop

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Get Group and Menu
Sub getgroupandmanu()

    lastcomprw = Sheets("Tile Selections").Cells(1, 1).End(xlDown).Row

    '''''Make Temp'''''''

    Dim wsSheet As Worksheet
    On Error Resume Next
    Set wsSheet = Sheets("Temp")
    On Error GoTo 0

    If wsSheet Is Nothing Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Temp"
        ActiveWorkbook.Sheets(1).Activate
    End If

    ''''''''''''''populate grouping column''''''''''''''''
    'The algorithm is the same as in the future but it takes all rows from the compilation. later it will need to back reference

    cl = 1 ' cl is 1 when working in groupings
    totalin = 1 ' this needs to reset for each column

    For csrw = 1 To lastcomprw ' csrw: compiled sheet row
        exists = False
    
        For arw = 1 To totalin + 1 ' arw: Temp sheet row
            If ActiveWorkbook.Sheets("Temp").Cells(arw, cl) = Sheets("Tile Selections").Cells(csrw + 1, cl) Then
                exists = True
                Exit For
            End If
        Next arw
    
        If exists = False Then
            ActiveWorkbook.Sheets("Temp").Cells(totalin + 1, cl) = Sheets("Tile Selections").Cells(csrw + 1, cl)
            totalin = totalin + 1
        End If

    Next csrw

    '''''Create Drop Down''''''''
    ssrw = 5

    Do Until Cells(ssrw, 1) = "TOTAL"
        If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) Then
            With ActiveWorkbook.ActiveSheet.Cells(ssrw, 16).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Temp!$A$1:$A$" & totalin
            .InCellDropdown = True
            End With
        End If
        ssrw = ssrw + 1
    Loop

    '''''For Manufacturers''''''

    cl = 2
    totalin = 1 ' this needs to reset for each column

    For csrw = 1 To lastcomprw ' csrw: compiled sheet row
        exists = False
    
        For arw = 1 To totalin + 1 ' arw: Temp sheet row
            If ActiveWorkbook.Sheets("Temp").Cells(arw, cl) = Sheets("Tile Selections").Cells(csrw + 1, 3) Then
                exists = True
                Exit For
            End If
        Next arw
    
        If exists = False Then
            ActiveWorkbook.Sheets("Temp").Cells(totalin + 1, cl) = Sheets("Tile Selections").Cells(csrw + 1, 3)
            totalin = totalin + 1
        End If

    Next csrw

    '''''Create Drop Down''''''''

    ssrw = 5
    Do Until Cells(ssrw, 1) = "TOTAL"
        If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) Then
            With ActiveWorkbook.ActiveSheet.Cells(ssrw, 17).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Temp!$B$1:$B$" & totalin
                .InCellDropdown = True
            End With
        End If
        ssrw = ssrw + 1
    Loop
End Sub




' Get tile style
Sub getstyle()

' get last row - const through prog
lastcomprw = Sheets("Tile Selections").Cells(1, 1).End(xlDown).Row

'''''''''''''

acl = 4
Do Until Cells(acl + 1, 1) = "TOTAL"
    totalin = 1
    If Not Cells(acl + 1, 1).Font.Bold And Not IsEmpty(Cells(acl + 1, 1)) And Not IsEmpty(Cells(acl + 1, 16)) And Not IsEmpty(Cells(acl + 1, 17)) Then
        
        For csrw = 2 To lastcomprw ' csrw: compiled sheet row
            exists = False
            For arw = 2 To totalin ' arw: Temp sheet row
            
                If ActiveWorkbook.Sheets("Temp").Cells(arw, acl) = Sheets("Tile Selections").Cells(csrw, 5) Then
                    exists = True
                    Exit For
                End If
            
            Next arw
    
            If (exists = False) And(ActiveWorkbook.ActiveSheet.Cells(acl + 1, 16) = Sheets("Tile Selections").Cells(csrw, 1)) And (ActiveWorkbook.ActiveSheet.Cells(acl + 1, 17) = Sheets("Tile Selections").Cells(csrw, 3)) Then
                ActiveWorkbook.Sheets("Temp").Cells(totalin + 1, acl) = Sheets("Tile Selections").Cells(csrw, 5)
                totalin = totalin + 1
            End If

        Next csrw
    
        Let ActiveWorkbook.Sheets("Temp").Cells(1, acl) = totalin ' first row in
         Temp stores totalin for that item col
    
    End If
    acl = acl + 1
Loop

'''''Create Drop Down''''''''

ssrw = 5
Do Until Cells(ssrw, 1) = "TOTAL"
    If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) And Not 
    IsEmpty(Cells(ssrw, 16)) And Not IsEmpty(Cells(ssrw, 17)) Then
    
        Let lastrw = ActiveWorkbook.Sheets("Temp").Cells(1, ssrw - 1).Value

        TempRangeName = Range(Cells(2, ssrw - 1), Cells(lastrw, ssrw - 1)).Address
    
        With ActiveWorkbook.ActiveSheet.Cells(ssrw, 18).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=Temp!" & TempRangeName
            .InCellDropdown = True
        End With
    End If
    ssrw = ssrw + 1
Loop

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Reset all tile options when completed'
Sub resettiles()
    Dim wsSheet As Worksheet
    On Error Resume Next
    Set wsSheet = Sheets("Temp")
    On Error GoTo 0
    If Not wsSheet Is Nothing Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets("Temp").Delete
        Application.DisplayAlerts = True
    End If

    ssrw = 5
    Do Until Cells(ssrw, 1) = "TOTAL"
        If Not Cells(ssrw, 1).Font.Bold And Not IsEmpty(Cells(ssrw, 1)) Then
            ActiveSheet.Range("P" & ssrw & ":T" & ssrw).Clear
        End If
        ssrw = ssrw + 1
    Loop
End Sub

