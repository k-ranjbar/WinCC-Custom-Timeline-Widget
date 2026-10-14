 On Error Resume Next

    Dim catalogName, conn, cmd, rs
    Dim offsetMin, nowLocal, winStart, qStart, qEnd
    Dim dts(1), txt(1)
    Dim k, d, c
    Dim wmi, obj

    ' ------------------------------------------------------------------------
    '  ‰ŸÌ„ „‰ÿﬁÂ “„«‰Ì (TimeZone)
    ' ------------------------------------------------------------------------
    offsetMin = 0
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each obj In wmi.ExecQuery("SELECT CurrentTimeZone FROM Win32_ComputerSystem")
        offsetMin = obj.CurrentTimeZone
    Next
    Err.Clear

    nowLocal = Now
    winStart = DateAdd("h", -24, nowLocal)

    qStart = DateAdd("n", -offsetMin, DateAdd("h", -48, nowLocal))
    qEnd = DateAdd("n", -offsetMin, nowLocal)

    dts(0) = qStart
    dts(1) = qEnd

    For k = 0 To 1
        d = dts(k)
        txt(k) = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & " " & _
                 Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2) & ".000"
    Next

    ' ê—› ‰ ‰«„ œÌ «»Ì” ¬—‘ÌÊ —«‰ù «Ì„
    catalogName = HMIRuntime.Tags("@DatasourceNameRT").Read
    If catalogName = "" Then Exit Sub

    ' « ’«· »Â WinCC OLEDB Provider
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=WinCCOLEDBProvider.1;Catalog=" & catalogName & ";Data Source=.\WinCC"
    conn.CursorLocation = 3
    conn.Open
    If Err.Number <> 0 Then Exit Sub

    ' ------------------------------------------------------------------------
    ' ¬—«ÌÂùÂ«Ì  ÿ»ÌﬁÌ œ” ê«ÂùÂ«°  êùÂ«Ì Å—œ«“‘Ì ¬—‘ÌÊ Ê  êùÂ«Ì œ«Œ·Ì „ ‰Ì
    ' ------------------------------------------------------------------------
    Dim devices(5)
    devices(0) = "Machine1"
    devices(1) = "Machine2"
    devices(2) = "Machine3"
    devices(3) = "Machine4"
    devices(4) = "Machine5"
    devices(5) = "Machine8"

    Dim archiveTags(5)
    archiveTags(0) = "FilterPress1_WTC_Mode"
    archiveTags(1) = "FilterPress2_WTC_Mode"
    archiveTags(2) = "FilterPress3_WTC_Mode"
    archiveTags(3) = "FilterPress4_WTC_Mode"
    archiveTags(4) = "FilterPress5_WTC_Mode"
    archiveTags(5) = "FilterPress8_WTC_Mode"

    Const ARCHIVE_NAME = "FilterPress"

    Dim devIdx, devName, processTag, archivePath
    Dim eMode(), eTime(), cnt, nRows
    Dim lastV, haveLast, haveStart, startMode
    Dim maxEntries, firstIdx, i, m, tok, p, parts(3)
    Dim sql, opened, v, tLocal

    maxEntries = 144 ' 36 * 4 »—«Ì ?  ê „ ‰Ì

    ' Õ·ﬁÂ Å—œ«“‘ Â— ? œ” ê«Â
    For devIdx = 0 To 5
        devName = devices(devIdx)
        processTag = archiveTags(devIdx)
        archivePath = ARCHIVE_NAME & "\" & processTag

        opened = False
        Set rs = Nothing

        ' «Ã—«Ì òÊ∆—Ì TAG:R »« „”Ì— ò«„· ¬—‘ÌÊ
        sql = "TAG:R,'" & archivePath & "','" & txt(0) & "','" & txt(1) & "'"
        
        Set cmd = CreateObject("ADODB.Command")
        cmd.CommandType = 1
        Set cmd.ActiveConnection = conn
        cmd.CommandText = sql
        
        Err.Clear
        Set rs = cmd.Execute
        
        If Err.Number = 0 And Not rs.EOF Then
            opened = True
        Else
            ' «ê— „”Ì— ò«„· ÃÊ«» ‰œ«œ°  ”  »« ‰«„  ê Œ«·Ì
            If Not rs Is Nothing Then 
                If rs.State = 1 Then rs.Close
            End If
            Set rs = Nothing
            Set cmd = Nothing

            sql = "TAG:R,'" & processTag & "','" & txt(0) & "','" & txt(1) & "'"
            Set cmd = CreateObject("ADODB.Command")
            cmd.CommandType = 1
            Set cmd.ActiveConnection = conn
            cmd.CommandText = sql
            
            Err.Clear
            Set rs = cmd.Execute
            If Err.Number = 0 And Not rs.EOF Then opened = True
        End If

        If opened Then
            ReDim eMode(5000)
            ReDim eTime(5000)
            cnt = 0
            nRows = 0
            lastV = 0
            haveLast = False
            haveStart = False

            Do While Not rs.EOF
                nRows = nRows + 1
                v = CInt(rs.Fields("RealValue").Value)

                If (Not haveLast) Or (v <> lastV) Then
                    tLocal = DateAdd("n", offsetMin, rs.Fields("TimeStamp").Value)

                    If tLocal <= winStart Then
                        startMode = v
                        haveStart = True
                    Else
                        If haveStart Then
                            eMode(cnt) = startMode
                            eTime(cnt) = winStart
                            cnt = cnt + 1
                            haveStart = False
                        End If

                        If cnt < 5000 Then
                            eMode(cnt) = v
                            eTime(cnt) = tLocal
                            cnt = cnt + 1
                        End If
                    End If
                    lastV = v
                    haveLast = True
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            Set cmd = Nothing

            If haveStart Then
                eMode(cnt) = startMode
                eTime(cnt) = winStart
                cnt = cnt + 1
            ElseIf cnt = 0 Then
                ' ŒÊ«‰œ‰ „ﬁœ«— “‰œÂ «“ PLC œ— ’Ê—  ‰»Êœ ”«»ﬁÂ œ— »«“Â
                eMode(cnt) = CInt(HMIRuntime.Tags("WWTP_PumpStation " & devName & "_WTC_Mode").Read)
                eTime(cnt) = winStart
                cnt = cnt + 1
            End If

            firstIdx = 0
            If cnt > maxEntries Then
                firstIdx = cnt - maxEntries
            End If

            For p = 0 To 3
                parts(p) = ""
            Next

            If eTime(firstIdx) = winStart Then
                parts(0) = "S"
            End If

            For i = firstIdx To cnt - 1
                m = eMode(i)
                If m < 0 Or m > 9 Then m = 9
                d = eTime(i)
                tok = CStr(m) & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
                p = (i - firstIdx) \ 36
                parts(p) = parts(p) & tok
            Next

            ' ‰Ê‘ ‰ Œ—ÊÃÌ œ— ?  ê „ ‰Ìˆ œ«Œ·Ìˆ «Œ ’«’Ìˆ Â„«‰ œ” ê«Â
            HMIRuntime.Tags(devName & "_ModeData1").Write parts(0)
            HMIRuntime.Tags(devName & "_ModeData2").Write parts(1)
            HMIRuntime.Tags(devName & "_ModeData3").Write parts(2)
            HMIRuntime.Tags(devName & "_ModeData4").Write parts(3)
        End If
    Next

    conn.Close
    Set conn = Nothing