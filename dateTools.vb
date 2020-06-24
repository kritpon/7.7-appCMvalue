Module dateTools
    Function changeInt2Time(setTime As Integer) As String

        Dim ans As String = ""
        Dim total_millisec As Integer = setTime * 1000
        Dim millisecNow As Integer = total_millisec
        ' lbCounter.Text = intCounter
        setTime = setTime * -1
        Dim hour As Integer = setTime \ 60 ' Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        Dim min As Integer = setTime Mod 60 'Math.Floor(((millisecNow / (60 * 1000)) - (hour * 60)))
        'Dim sec As Integer = Math.Floor((millisecNow / 1000) - (hour * 60 * 60) - (min * 60))
        'Dim mil As Integer '= 100 * (((millisecNow / 1000) - (hour * 60 * 60) - (min * 60)) - sec)
        'mil = Math.Floor(millisecNow / 100 - (hour * 60 * 60) - (min * 60) - (sec * 100))
        ' lbMilSec.Text = mil
        ans = hour.ToString("00") + ":" + min.ToString("00")  '+ "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans

    End Function
    Function incDate_H(strDate As String, strTime As String) As String

        Dim strDD As String
        Dim strMM As String
        Dim strYY As String
        Dim strDate2 As String

        If CInt(Format(CDate(strTime), "HH").ToString) >= 0 And (CInt(Format(CDate(strTime), "HH").ToString) < 8) Then

            ' MsgBox(strTime & "&-Left = " & Format(CDate(strTime), "HH").ToString)
            strDD = (CInt(Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2))) + 1).ToString

            If strDD > 31 Then
                strDD = "01"

            End If
            strDate = DateAdd(DateInterval.Day, 1, CDate(strDate))
            strDD = Format(Microsoft.VisualBasic.DateAndTime.Day(strDate), "00")
            'strDD = Format((CInt(Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2)))), "0#").ToString
            '(CInt(Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2))) + 1).ToString
            'If strDD > 31 Then
            '    strDD = "1"
            'Else

            'End If
        Else
            '   strDD = Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2))
            strDD = Format(Microsoft.VisualBasic.DateAndTime.Day(strDate), "00")

        End If

        'strMM = Right(Left(strDate, 5), 2)  'Month(strDate) '
        strMM = Format(Month(strDate), "00")
        'strDD = Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2)) 'Microsoft.VisualBasic.DateAndTime.Day(strDate) '
        strYY = Trim(Microsoft.VisualBasic.Right(strDate, 4))
        If CInt(strYY) > 2562 Then
            strYY = Str(Int(Year(Now)) - 543)
        Else
            strYY = Str(Int(Year(Now)) - 543)
        End If

        strDate2 = strMM & "-" & strDD & "-" & Trim(strYY) & " " & strTime


        Return strDate2


    End Function


    Function strToDate(strDate As String, strTime As String) As String

        Dim strDD As String
        Dim strMM As String
        Dim strYY As String
        Dim strDate2 As String



        ' trh_Date = Format(Month((.Range("H" & countRow + 1).Value)), "00") 
        '& "/" & Format(Microsoft.VisualBasic.Day((.Range("H" & countRow + 1).Value)), "00")
        '& "/" & (Year((.Range("H" & countRow + 1).Value)) - 543)

        'If Len(strDate) = 17 Then

        strMM = Trim(Microsoft.VisualBasic.Left(strDate, 2))  'Month(strDate) '
        strDD = Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2)) 'Microsoft.VisualBasic.DateAndTime.Day(strDate) '
        strYY = Trim(Microsoft.VisualBasic.Right(strDate, 4))
        If CInt(strYY) > 2562 Then
            strYY = Str(Int(Year(Now)) - 543)
        Else
            strYY = Str(Int(Year(Now)) - 543)
        End If

        'If CInt(strMM) > 12 Then
        '    strDate2 = strDD & "-" & strMM & "-" & Trim(strYY) & " " & strTime
        'Else

        'End If
        strDate2 = strMM & "-" & strDD & "-" & Trim(strYY) & " " & strTime
        '   Dim strChkDate As DateTime = strDate2
        'Else
        '    strMM = (Microsoft.VisualBasic.Left(strDate, 2))
        '    strDD = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 5), 2)
        '    strYY = Year(strDate)
        '    If CInt(strYY) > 2562 Then
        '        strYY = Str(Int(Year(Now)) - 543)
        '    Else
        '        strYY = Str(Int(Year(Now)) - 543)
        '    End If
        '    strTime = Microsoft.VisualBasic.Right(strDate, 8)
        '    strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        'End If

        Return strDate2


    End Function

    Function strToDate2(strDate As String) As String
        Dim strDD As String
        Dim strMM As String
        Dim strYY As String
        Dim strDate2 As String
        Dim strTime As String


        ' trh_Date = Format(Month((.Range("H" & countRow + 1).Value)), "00") 
        '& "/" & Format(Microsoft.VisualBasic.Day((.Range("H" & countRow + 1).Value)), "00")
        '& "/" & (Year((.Range("H" & countRow + 1).Value)) - 543)

        'If Len(strDate) = 17 Then

        strMM = Month(strDate) 'Trim(Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 5)), 2)) '
        strDD = Microsoft.VisualBasic.DateAndTime.Day(strDate) 'Trim(Microsoft.VisualBasic.Left(strDate, 2)) '
        strYY = Trim(Year(strDate))
        If CInt(strYY) > 2562 Then
            strYY = Str(Int(Year(Now)) - 543)
        Else
            strYY = Str(Int(Year(Now)) - 543)
        End If
        strTime = Microsoft.VisualBasic.Right(strDate, 8)
        strDate2 = strMM & "-" & strDD & "-" & strYY & " " & strTime

        'Else
        '    strMM = (Microsoft.VisualBasic.Left(strDate, 2))
        '    strDD = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 5), 2)
        '    strYY = Year(strDate)
        '    If CInt(strYY) > 2562 Then
        '        strYY = Str(Int(Year(Now)) - 543)
        '    Else
        '        strYY = Str(Int(Year(Now)) - 543)
        '    End If
        '    strTime = Microsoft.VisualBasic.Right(strDate, 8)
        '    strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        'End If

        Return strDate2


    End Function

    Function strToDate3(strDate As String, intSel As Integer) As String  '  ถ้า intSel =1 เป็นวันที่กับเวลา ถ้า เป็น 2 มีแต่วันที่ 

        Dim strDD As String
        Dim strMM As String
        Dim strYY As String
        Dim strDate2 As String = ""
        Dim strTime As String


        ' trh_Date = Format(Month((.Range("H" & countRow + 1).Value)), "00") 
        '& "/" & Format(Microsoft.VisualBasic.Day((.Range("H" & countRow + 1).Value)), "00")
        '& "/" & (Year((.Range("H" & countRow + 1).Value)) - 543)

        'If Len(strDate) = 17 Then

        strMM = Month(strDate) 'Microsoft.VisualBasic.Right((Microsoft.VisualBasic.Left(strDate, 4)), 2)
        strDD = Microsoft.VisualBasic.DateAndTime.Day(strDate) 'Microsoft.VisualBasic.Left(strDate, 1)
        strYY = Year(strDate)
        If CInt(strYY) > 2562 Then
            strYY = Str(Int(Year(Now)) - 543)
        Else
            strYY = Str(Int(Year(Now)) - 543)
        End If
        strTime = Microsoft.VisualBasic.Right(strDate, 8)
        If intSel = 1 Then ' มีวันที่ และ เวลา

            strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        ElseIf intSel = 2 Then ' ไม่มีเวลาา

            strDate2 = strMM & "-" & strDD & "-" & strYY '& " " & strTime

        End If
        'Else
        '    strMM = (Microsoft.VisualBasic.Left(strDate, 2))
        '    strDD = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(strDate, 5), 2)
        '    strYY = Year(strDate)
        '    If CInt(strYY) > 2562 Then
        '        strYY = Str(Int(Year(Now)) - 543)
        '    Else
        '        strYY = Str(Int(Year(Now)) - 543)
        '    End If
        '    strTime = Microsoft.VisualBasic.Right(strDate, 8)
        '    strDate2 = strDD & "-" & strMM & "-" & strYY & " " & strTime

        'End If

        Return strDate2

    End Function
End Module
