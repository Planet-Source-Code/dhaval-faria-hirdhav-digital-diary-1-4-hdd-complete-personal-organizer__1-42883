Attribute VB_Name = "modInsertDates"

Public CurrentMonth As String
Public CurrentYear
Public FirstDay
Public LastDate
Public MyI
Public NowDates
Public LastDates
Public TotalDates
Public FirstDates
Public CurrentMonthNo

Public Function IDs()
FirstDay = Format("01/" & CurrentMonth & "/" & CurrentYear, "DDD")

If CurrentMonth = "Jan" Then
    TotalDates = 31
ElseIf CurrentMonth = "Feb" Then
    For z = 1 To 10000
        If CurrentYear / 4 = z Then
            TotalDates = 29
            GoTo 10
        Else
            TotalDates = 28
        End If
    Next z
ElseIf CurrentMonth = "Mar" Then
    TotalDates = 31
ElseIf CurrentMonth = "Apr" Then
    TotalDates = 30
ElseIf CurrentMonth = "May" Then
    TotalDates = 31
ElseIf CurrentMonth = "Jun" Then
    TotalDates = 30
ElseIf CurrentMonth = "Jul" Then
    TotalDates = 31
ElseIf CurrentMonth = "Aug" Then
    TotalDates = 31
ElseIf CurrentMonth = "Sep" Then
    TotalDates = 30
ElseIf CurrentMonth = "Oct" Then
    TotalDates = 31
ElseIf CurrentMonth = "Nov" Then
    TotalDates = 30
ElseIf CurrentMonth = "Dec" Then
    TotalDates = 31
End If

10

If FirstDay = "Fri" Then
    MyI = 5
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Sat" Then
    MyI = 6
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Tue" Then
    MyI = 2
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Sun" Then
    MyI = 0
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Mon" Then
    MyI = 1
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Wed" Then
    MyI = 3
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
ElseIf FirstDay = "Thu" Then
    MyI = 4
    NowDates = 0
    For I = MyI To TotalDates + MyI - 1
        NowDates = NowDates + 1
        frmCalender.lblDate(I).Enabled = True
        frmCalender.lblDate(I).Caption = NowDates
        LastDate = I + 1
    Next I
    LastDates = 0
    For j = 0 To MyI - 1
        frmCalender.lblDate(j).Caption = ""
    Next j
    For k = TotalDates + MyI To 36
        frmCalender.lblDate(k).Caption = ""
    Next k
    Exit Function
Else
    MsgBox FirstDay
End If
End Function

