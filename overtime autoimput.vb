Option Explicit
Function set_format(ByRef selection As Range)
Dim rng As Range
    With selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        For Each rng In Range("a1:p1")
            .EntireColumn.AutoFit
        Next
    End With
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With selection
        .Font.Size = 9
    End With
End Function
Function gen_header()
    Static arr_header() As String
    Static str_header As String
    Dim i As Integer
    str_header = "年度,月份,部门,科室,员工编号,姓名,加班类型,加班类型,加班地点,加班原因,加班日期,起始时间,结束时间,科室审批人,部门审批人,是否调休"
    arr_header = Split(str_header, ",")
    With Sheet1
        .Range("a1").Resize(1, 16) = arr_header
    End With
    Call set_format(Sheet1.Range("a1:p1"))
    Sheet1.Range("a1:p1").Interior.ColorIndex = 15
End Function
Function gen_cal()
    Dim firstday As Date
    Dim i, j, firstweek As Integer
    Dim datelist(1 To 32) As Date
    Dim rng As Range
    If Day(Now()) < 27 Then
        firstday = DateSerial(Year(Now()), Month(Now()) - 1, 26)
    Else
        firstday = DateSerial(Year(Now()), Month(Now()), 26)
    End If
    With Sheet1
        .Range("a38").Resize(1, 7) = Split("星期一,星期二,星期三,星期四,星期五,星期六,星期日", ",")
        firstweek = Weekday(firstday, vbMonday)
        datelist(1) = firstday
        For i = 2 To 32
            datelist(i) = datelist(i - 1) + 1
        Next i
        .Range("a39:g43").Clear
        Call set_format(.Range("a38:g44"))
        For i = 1 To 32
            If Day(datelist(i + 1)) = 26 Then
                Exit For
            End If
            If (i + firstweek - 1) Mod 7 = 0 Then
                Set rng = .Cells(Fix((i + firstweek - 1) / 7) + 38, 7)
            Else
                Set rng = .Cells(Fix((i + firstweek - 1) / 7) + 39, ((i + firstweek - 1) Mod 7))
            End If
                With rng
                    .NumberFormat = "MM-dd"
                    .Value = datelist(i)
                    .Interior.ColorIndex = 15
                End With
        Next i
    End With
End Function
Function gen_data(ByVal add_date As Date)
    Dim datalist As New Collection
    Dim i, k As Integer
    Dim rng As Range
    With datalist   '"年度,月份,部门,科室,员工编号,姓名,加班类型,加班类型,加班地点,加班原因,加班日期,起始时间,结束时间,科室审批人,部门审批人,是否调休"
        .Add Year(Now())
        .Add Month(Now())
        .Add "部门"
        .Add "科室"
        .Add 179518
        .Add "张涛"
        If Weekday(add_date, vbMonday) = 6 Or Weekday(add_date, vbMonday) = 7 Then
            .Add "假日加班"
        Else
            .Add "平日加班"
        End If
        .Add "白班"
        .Add "办公室"
        .Add "G项目"
        .Add add_date
        .Add add_date + TimeSerial(17, 30, 0)
        .Add add_date + TimeSerial(21, 30, 0)
        .Add "科室审批人"
        .Add "部门审批"
        If Weekday(add_date, vbMonday) = 6 Or Weekday(add_date, vbMonday) = 7 Then
            .Add 1
        Else
            .Add 0
        End If
    End With
    k = get_lastrow
    For i = 1 To 16 Step 1
        With Sheet1.Cells(k, i)
            .Value = datalist.Item(i)
            Select Case i
                Case Is = 7
                    Call valid_7(Sheet1.Cells(k, i))
                Case Is = 8
                    Call valid_8(Sheet1.Cells(k, i))
                Case Is = 11
                    .NumberFormat = "yyyy-MM-dd"
                Case Is = 12
                    .NumberFormat = "yyyy-MM-dd hh:mm"
                    Call date_valid(Sheet1.Cells(k, i), TimeSerial(0, 0, 0))
                Case Is = 13
                    .NumberFormat = "yyyy-MM-dd hh:mm"
                    Set rng = Sheet1.Cells(k, i - 1)
                    Call date_valid(Sheet1.Cells(k, i), TimeSerial(Hour(rng), Minute(rng), 0))
                Case Is = 16
                    Call valid_16(Sheet1.Cells(k, i))
            End Select
            .EntireRow.Hidden = False
        End With
        Call set_format(Sheet1.Cells(k, i))
    Next i
End Function
Function get_lastrow()
    Dim k As Integer
    k = 2
    While Sheet1.Cells(k, 11).Value <> ""
        k = k + 1
    Wend
    get_lastrow = k
End Function
Function delete_data(ByVal del_date As Date)
    Dim k, i As Integer
    k = get_lastrow
    For i = 2 To k
        If Sheet1.Cells(i, 11) = del_date Then
            Sheet1.Rows(i) = ""
        End If
    Next i
    With Range("a38:p38").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
    End With
End Function
Function choosedate(ByVal choisedate As Range)
    With choisedate.Interior
        If .ColorIndex = 6 Then
            delete_data (choisedate.Value)
            .ColorIndex = 15
        Else
            .ColorIndex = 6
            gen_data (choisedate.Value)
        End If
        sort_data
        hiddenrow
    End With
End Function
Function sort_data()
    With Sheet1
        .Range("a1:p32").Sort key1:=[k1], order1:=xlAscending, Header:=xlYes
    End With
End Function
Function hiddenrow()
    Dim k, i As Integer
    k = get_lastrow
    With Sheet1
        For i = k To 37
            .Cells(i, 1).EntireRow.Hidden = True
        Next i
    End With
End Function
Function date_valid(ByVal rng As Range, start As Date)
    Dim datalist As String
    Dim i, j As Integer
    For i = 0 To 23
        For j = 0 To 45 Step 15
            If TimeSerial(i, j, 0) >= start Then
            datalist = datalist & CStr(Format(TimeSerial(i, j, 0), "hh:mm")) & ","
            End If
        Next j
    Next i
    datalist = IIf(Len(datalist) > 0, Left(datalist, Len(datalist) - 1), datalist)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=datalist
    End With
End Function
Function valid_7(ByVal rng As Range)
    Static strlist As String
    strlist = "平日加班,假日加班,节日加班"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=strlist
    End With
End Function
Function valid_8(ByVal rng As Range)
    Static strlist As String
    strlist = "白班,中班,晚班,特殊排班"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=strlist
    End With
End Function
Function valid_16(ByVal rng As Range)
    Static strlist As String
    strlist = "0,1"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=strlist
    End With
End Function
Function base_info()
    With Sheet1
        Call set_format(.Range("h38:k41"))
        .Range("h38:h41") = WorksheetFunction.Transpose(Split("部门,科室,部门审批,科室审批", ","))
        
        .Range("j38:j41") = WorksheetFunction.Transpose(Split("姓名,员工编号,起始时间,结束时间", ","))
        .Range("k38:k41") = WorksheetFunction.Transpose(Split("张涛,179518,17:30,21:00", ","))
        Call date_valid(.Range("k40"), TimeSerial(0, 0, 0))
        Call date_valid(.Range("k41"), TimeSerial(0, 0, 0))
    End With
End Function
Sub worksheet_beforedoubleclick(ByVal target As Range, cancel As Boolean)
    Dim rng As Range
    cancel = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If target.Row = 1 Then
        Cells.Clear
        hiddenrow
        gen_header
        gen_cal
        base_info
    End If
    If Not Intersect(target(1), Range("d39:j43")) Is Nothing And target(1) <> "" Then
        choosedate (target(1))
    End If
    For Each rng In Sheet1.Range("a1:p1")
        rng.EntireColumn.AutoFit
    Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
Sub worksheet_change(ByVal target As Range)
    Dim m, n, p As Date
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If target.Column = 12 Or target.Column = 13 Then
        With Sheet1
            m = .Cells(target.Row, 11)
            n = .Cells(target.Row, 12)
            p = .Cells(target.Row, 13)
            .Cells(target.Row, 12) = m + TimeSerial(Hour(n), Minute(n), Second(n))
            .Cells(target.Row, 13) = m + TimeSerial(Hour(p), Minute(p), Second(p))
        End With
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Sub worksheet_activate()
    Sheet1.Range("1:1").EntireColumn.Hidden = False
    Sheet1.Range("a:a").EntireRow.Hidden = False
End Sub
