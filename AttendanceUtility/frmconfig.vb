Imports System.IO
Imports System.IO.Stream
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
Imports System.Text
'Ver 2.4.0-E62 Start
Imports System.Security.Cryptography
'Ver 2.4.0-E62 End
'Ver 2.3.8-2678 Start
Imports System.Text.RegularExpressions
Imports AttendanceUtility.ApplicationMain
Imports System.ServiceProcess
Imports System.Globalization

Public Class frmconfig
    Dim strFileName As String
    Dim strFileName1 As String
    Public strDonglePath As String
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim objReader As StreamReader
            Dim objWriter As StreamWriter

            If Not Directory.Exists("C:\AttDBPath") Then
                Directory.CreateDirectory("C:\AttDBPath")
            End If

            strFileName = "C:\AttDBPath\destpath.ini"
            'strFileName = Application.StartupPath & "\hostpath.ini"
            'If File.Exists(strFileName) = False Then
            '    File.CreateText(strFileName)
            'End If
            strDonglePath = ""
            objWriter = New StreamWriter(strFileName)
            objWriter.WriteLine("SERVER =" & txtServerName.Text)
            If OptWindows.Checked = True Then
                'objWriter.WriteLine("AUTHENTICATE=WINNT")
                'ra 2.0
                objWriter.WriteLine("AUTHENTICATE=Windows")
                'ra 2.0
            Else
                objWriter.WriteLine("AUTHENTICATE=SQL")
            End If
            objWriter.WriteLine("USERID=" & txtUser.Text)
            objWriter.WriteLine("PWD=" & txtPwd.Text)
            objWriter.WriteLine("DonglePath=" & strDonglePath)
            objWriter.WriteLine("DB=" & txtdb.Text)
            objWriter.Close()

            EnCrypt(strFileName)

            MsgBox("Created Connection for Database", MsgBoxStyle.Information, StrSQLDB)

        Catch ex As Exception
            MsgBox("Failed to Create Connection for Database", MsgBoxStyle.Information, StrSQLDB)
        End Try

    End Sub

    Private Sub btnconn_Click(sender As Object, e As EventArgs) Handles btnconn.Click
        Dim strLine As String
        'strFileName = Application.StartupPath & "\hostpath.ini"
        strFileName = "C:\AttDBPath\destpath.ini"
        If File.Exists(strFileName) = False Then

        Else

            Dim strmsg As String
            Dim IsFileEncrypted = CheckEncrypt(strFileName) 'Added code for encrypt check file

            If (IsFileEncrypted = True) Then
                strmsg = Decrypt(strFileName)
            Else
                strmsg = File.ReadAllText(strFileName)
            End If

            Dim mline() As String = strmsg.Split(Chr(13))
            strLine = mline(0)
            StrSvrName = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSvrName = "DESKTOP-FVN1B19\MSSQLSERVER1"
            strLine = mline(1)
            StrAuthentication = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(2)
            StrSQLUserId = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(3)
            StrSQLPwd = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(5)
            StrSQLDB = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(4)
            strDonglePath = Mid(strLine, InStr(1, strLine, "=") + 1)

        End If

        If StrAuthentication = "Windows" Then
            conDestData = New SqlClient.SqlConnection("user id=" & StrSQLUserId & ";Password=" & StrSQLPwd & ";integrated security=SSPI;initial catalog=" & StrSQLDB & ";data source=" & StrSvrName & ";MultipleActiveResultSets = true")
        Else
            conDestData = New SqlClient.SqlConnection("user id=" & StrSQLUserId & ";Password=" & StrSQLPwd & ";initial catalog=" & StrSQLDB & ";data source=" & StrSvrName & ";MultipleActiveResultSets = true")
        End If

        Try
            conDestData.Open()
            MsgBox(StrSQLDB & " Connection Established Successfully", MsgBoxStyle.Information, StrSQLDB)
        Catch ex As Exception
            conDestData.Close()
            conDestData.Dispose()

            MsgBox(StrSQLDB & " Connection Failed", MsgBoxStyle.Information, StrSQLDB)
        End Try


        'Dim dtblda As SqlDataAdapter
        'Dim dtbllist As New DataSet

        'dtblda = New SqlDataAdapter("SELECT table_name FROM information_schema.tables WHERE table_type = 'base table'", conDestData)
        'dtblda.Fill(dtbllist, "dsttbl")

        'For i = 0 To dtbllist.Tables(0).Rows.Count - 1
        '    cmbdest.Items.Add(dtbllist.Tables(0).Rows(i).Item(0))
        'Next


    End Sub


    Private Sub btnconn1_Click(sender As Object, e As EventArgs) Handles btnconn1.Click
        Dim strLine As String
        'strFileName = Application.StartupPath & "\srcpath.ini"
        strFileName = "C:\AttDBPath\srcpath.ini"
        If File.Exists(strFileName) = False Then

        Else

            Dim strmsg As String
            Dim IsFileEncrypted = CheckEncrypt(strFileName) 'Added code for encrypt check file
            'Dim strmsg As String = Decrypt(strFileName)
            If (IsFileEncrypted = True) Then
                strmsg = Decrypt(strFileName)
            Else
                strmsg = File.ReadAllText(strFileName)
            End If

            Dim mline() As String = strmsg.Split(Chr(13))
            strLine = mline(0)
            StrSvrName1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSvrName = "DESKTOP-FVN1B19\MSSQLSERVER1"
            strLine = mline(1)
            StrAuthentication1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(2)
            StrSQLUserId1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(3)
            StrSQLPwd1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(5)
            StrSQLDB1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSQLPwd = "1234"

            strLine = mline(4)
            strDonglePath = Mid(strLine, InStr(1, strLine, "=") + 1)

        End If

        'Open TdsPac Data
        If StrAuthentication = "Windows" Then
            conSrcData = New SqlClient.SqlConnection("user id=" & StrSQLUserId1 & ";Password=" & StrSQLPwd1 & ";integrated security=SSPI;initial catalog=" & StrSQLDB1 & ";data source=" & StrSvrName1 & ";MultipleActiveResultSets = true")
        Else
            conSrcData = New SqlClient.SqlConnection("user id=" & StrSQLUserId1 & ";Password=" & StrSQLPwd1 & ";initial catalog=" & StrSQLDB1 & ";data source=" & StrSvrName1 & ";MultipleActiveResultSets = true")
        End If

        Try
            conSrcData.Open()
            MsgBox(StrSQLDB1 & " Connection Established Successfully", MsgBoxStyle.Information, StrSQLDB1)
        Catch ex As Exception
            conSrcData.Close()
            conSrcData.Dispose()
            MsgBox(StrSQLDB1 & " Connection Failed", MsgBoxStyle.Information, StrSQLDB1)
        End Try

        Dim stblda As SqlDataAdapter
        Dim stbllist As New DataSet

        cmbsrc.Items.Clear()

        stblda = New SqlDataAdapter("SELECT table_name FROM information_schema.tables WHERE table_type = 'base table' order by table_name", conSrcData)
        stblda.Fill(stbllist, "dsttbl")

        For i = 0 To stbllist.Tables(0).Rows.Count - 1
            cmbsrc.Items.Add(stbllist.Tables(0).Rows(i).Item(0))
        Next


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim objReader As StreamReader
            Dim objWriter As StreamWriter
            If Not Directory.Exists("C:\AttDBPath") Then
                Directory.CreateDirectory("C:\AttDBPath")
            End If

            strFileName = "C:\AttDBPath\srcpath.ini"
            'strFileName = Application.StartupPath & "\srcpath.ini"
            'If File.Exists(strFileName) = False Then
            '    File.CreateText(strFileName)
            'End If
            strDonglePath = ""

            objWriter = New StreamWriter(strFileName)
            objWriter.WriteLine("SERVER =" & txtServerName1.Text)
            If OptWindows.Checked = True Then
                'objWriter.WriteLine("AUTHENTICATE=WINNT")
                'ra 2.0
                objWriter.WriteLine("AUTHENTICATE=Windows")
                'ra 2.0
            Else
                objWriter.WriteLine("AUTHENTICATE=SQL")
            End If
            objWriter.WriteLine("USERID=" & txtUser1.Text)
            objWriter.WriteLine("PWD=" & txtPwd1.Text)
            objWriter.WriteLine("DonglePath=" & strDonglePath)
            objWriter.WriteLine("DB=" & txtdb1.Text)
            objWriter.Close()

            EnCrypt(strFileName)


            MsgBox("Created Connection for Database", MsgBoxStyle.Information, StrSQLDB)

        Catch ex As Exception
            MsgBox("Failed to Create Connection for Database", MsgBoxStyle.Information, StrSQLDB)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim strLine As String
        'strFileName = Application.StartupPath & "\hostpath.ini"
        strFileName = "C:\AttDBPath\destpath.ini"

        If File.Exists(strFileName) = False Then
            MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
        Else

            Dim strmsg As String
            Dim IsFileEncrypted = CheckEncrypt(strFileName) 'Added code for encrypt check file

            If (IsFileEncrypted = True) Then
                strmsg = Decrypt(strFileName)
            Else
                strmsg = File.ReadAllText(strFileName)
            End If

            Dim mline() As String = strmsg.Split(Chr(13))
            strLine = mline(0)
            StrSvrName = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSvrName = "DESKTOP-FVN1B19\MSSQLSERVER1"
            strLine = mline(1)
            StrAuthentication = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(2)
            StrSQLUserId = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(3)
            StrSQLPwd = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(5)
            StrSQLDB = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(4)
            strDonglePath = Mid(strLine, InStr(1, strLine, "=") + 1)

        End If

        If StrAuthentication = "Windows" Then
            OptWindows.Select()
        End If
        If StrAuthentication = "SQL" Then
            OptSQL.Select()
        End If

        txtServerName.Text = StrSvrName
        txtUser.Text = StrSQLUserId
        txtPwd.Text = StrSQLPwd
        'txtDbrowse.Text = ""
        txtdb.Text = StrSQLDB


        Dim strLine1 As String
        'strFileName1 = Application.StartupPath & "\srcpath.ini"
        strFileName1 = "C:\AttDBPath\srcpath.ini"
        If File.Exists(strFileName1) = False Then
            MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
        Else

            Dim strmsg1 As String
            Dim IsFileEncrypted = CheckEncrypt(strFileName1) 'Added code for encrypt check file
            'Dim strmsg As String = Decrypt(strFileName)
            If (IsFileEncrypted = True) Then
                strmsg1 = Decrypt(strFileName1)
            Else
                strmsg1 = File.ReadAllText(strFileName1)
            End If

            Dim mline() As String = strmsg1.Split(Chr(13))
            strLine = mline(0)
            StrSvrName1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSvrName = "DESKTOP-FVN1B19\MSSQLSERVER1"
            strLine = mline(1)
            StrAuthentication1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(2)
            StrSQLUserId1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(3)
            StrSQLPwd1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            strLine = mline(5)
            StrSQLDB1 = Mid(strLine, InStr(1, strLine, "=") + 1)
            'StrSQLPwd = "1234"

            strLine = mline(4)
            strDonglePath = Mid(strLine, InStr(1, strLine, "=") + 1)

        End If

        If StrAuthentication1 = "Windows" Then
            OptWindows.Select()
        End If
        If StrAuthentication1 = "SQL" Then
            OptSQL.Select()
        End If

        txtServerName1.Text = StrSvrName1
        txtUser1.Text = StrSQLUserId1
        txtPwd1.Text = StrSQLPwd1
        'txtDbrowse.Text = ""
        txtdb1.Text = StrSQLDB1

    End Sub

    Private Sub btnexport_Click(sender As Object, e As EventArgs) Handles btnexport.Click

        insert_data()

        'Dim da As SqlDataAdapter
        'Dim ds As New DataSet

        'Dim daSAL As SqlDataAdapter
        'Dim dsSAL As New DataSet

        'da = New SqlDataAdapter("select * from DeviceLogs_9_2022_copy  where   isnull(Flag,0)=0 order by DeviceLogId", conSrcData)
        'da.Fill(ds, "emplog")


        ''UserId, LogDate, Direction, DeviceLogId


        'For i = 0 To ds.Tables(0).Rows.Count - 1
        '    Dim dataRow As DataRow = ds.Tables(0).Rows(i)
        '    Dim columnName As String = dataRow("UserId")
        '    Dim column As DataColumn = New DataColumn(columnName)
        '    'DataTable.Columns.Add(column)
        '    Dim cmdTrans As SqlCommand
        '    cmdTrans = New SqlCommand
        '    cmdTrans.Connection = conDestData
        '    Try
        '        cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
        '                   "Values ('" & dataRow("UserId") & "','" & dataRow("LogDate") & "','" & dataRow("Direction") & "','" & dataRow("DeviceLogId") & "')"
        '        cmdTrans.ExecuteNonQuery()
        '        cmdTrans.Dispose()

        '        Dim cmdAlter As New SqlClient.SqlCommand
        '        cmdAlter.Connection = conSrcData
        '        cmdAlter.CommandText = "update DeviceLogs_9_2022_copy set Flag=1 where UserId='" & dataRow("UserId") & "' and DeviceLogId='" & dataRow("DeviceLogId") & "'"
        '        cmdAlter.ExecuteNonQuery()
        '        cmdAlter.Dispose()

        '    Catch ex As Exception

        '    End Try


        'Next

    End Sub

    Public Function insertdata()
        Dim strline As String
        Dim strFile As String
        strFileName = "C:\AttDBPath\tablemap.ini"
        strFile = "C:\AttDBPath\tablemain.ini"

        If File.Exists(strFileName) = False And File.Exists(strFile) = False Then
            MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
        Else
            Dim strmsg As String
            'strmsg = File.ReadAllText(strFileName)


            'If (IsFileEncrypted = True) Then
            strmsg = Decrypt(strFile)
            'Else
            '    strmsg = File.ReadAllText(strFile)
            'End If
            Dim mline() As String = strmsg.Split(Chr(13))

            Dim mlinecount As Integer
            mlinecount = mline.Length


            Dim str_msg As String
            'If (IsFileEncrypted = True) Then
            str_msg = Decrypt(strFileName)
            'Else
            '    str_msg = File.ReadAllText(strFileName)
            'End If
            Dim mline1() As String = str_msg.Split(Chr(13))

            Dim mlinecount1 As Integer
            mlinecount1 = mline1.Length
            Dim tbl1, tbl2 As String
            strline = mline1(0)
            tbl1 = mline1(0)


            Dim empcode, direct, sno As String
            Dim logdate As String


            Dim tblcolumn As String = String.Empty

            Dim i As Integer
            For i = 2 To mline.Length - 2

                tblcolumn = tblcolumn + mline(i).Replace(vbLf, "")
                If i <> mline.Length - 2 Then
                    tblcolumn = tblcolumn + ","
                End If

            Next

            'tbl2 = mline(1).Replace(vbLf, "")
            'empcode = mline(2).Replace(vbLf, "")
            'logdate = mline(3).Replace(vbLf, "")
            'direct = mline(4).Replace(vbLf, "")
            'sno = mline(5).Replace(vbLf, "")



            sno = mline1(5).Replace(vbLf, "")

            Dim da As SqlDataAdapter
            Dim ds As New DataSet

            Dim daSAL As SqlDataAdapter
            Dim dsSAL As New DataSet
            'DeviceLogs_9_2022_copy
            'Parallal
            Dim strtb As String
            strtb = "select  top 10  * from " & tbl1 & "  where   isnull(Flag,0) = 0  order by  " & sno & ""
            da = New SqlDataAdapter(strtb, conSrcData)
            'da.Fill(ds, "emplog")

            Try
                If Not da Is Nothing Then
                    da.Fill(ds, "emplog")
                Else
                    MsgBox("There is no data!", MsgBoxStyle.Critical)
                End If
            Catch ex As Exception
                MsgBox("Invalid Table", MsgBoxStyle.Critical)
                Exit Function
            End Try

            Dim errcnt As Integer
            errcnt = 0

            'UserId, LogDate, Direction, DeviceLogId


            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim dataRow As DataRow = ds.Tables(0).Rows(i)

                Dim cmdTrans As SqlCommand
                cmdTrans = New SqlCommand
                cmdTrans.Connection = conDestData
                Try
                    'cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
                    '           "Values ('" & empcode & "','" & logdate & "','" & direct & "','" & sno & "')"


                    Dim strqry As String
                    Dim strcnt As Integer
                    Dim strvalue As String = String.Empty

                    strqry = "Insert into Parallel (" & tblcolumn & ")" & "Values ("
                    For strcnt = 2 To mline1.Length - 2
                        If strcnt <> 5 Then
                            Dim dttype As String
                            'dttype = dataRow(mline1(strcnt).Replace(vbLf, "")).Item("DataTypeName")
                            strvalue += "'" + dataRow(mline1(strcnt).Replace(vbLf, "")) + "'" + ","
                        End If

                        If strcnt = 5 Then
                            strvalue += dataRow(mline1(strcnt).Replace(vbLf, "")) & ","
                        End If
                    Next
                    strvalue = strvalue.Remove(strvalue.Length - 1, 1)
                    strqry += strvalue + ")"

                    '('" & dataRow(empcode) & "','" & dataRow(logdate) & "','" & dataRow(direct) & "','" & dataRow(sno) & "')"

                    Dim slnoint As Integer
                    slnoint = Convert.ToInt32(dataRow(mline1(5).Replace(vbLf, "")))



                    cmdTrans.CommandText = strqry
                    cmdTrans.ExecuteNonQuery()
                    cmdTrans.Dispose()

                    Dim cmdAlter As New SqlClient.SqlCommand
                    cmdAlter.Connection = conSrcData
                    Dim altstr As String
                    altstr = "update  " & tbl1 & "  set Flag=1 where " & mline1(2).Replace(vbLf, "") & "='" + dataRow(mline1(2).Replace(vbLf, "")) + "' and " & mline1(5).Replace(vbLf, "") & "= '" + Convert.ToString(dataRow(mline1(5).Replace(vbLf, ""))) + "'"
                    cmdAlter.CommandText = altstr
                    'cmdAlter.CommandText = "update  " & tbl1 & "  set Flag=1 where " & empcode & "='" & dataRow(empcode) & "' and " & sno & "='" & dataRow(sno) & "'"
                    cmdAlter.ExecuteNonQuery()
                    cmdAlter.Dispose()

                Catch ex As Exception
                    errcnt = errcnt + 1
                    'MsgBox("Failed to Export the Data  at Rowno '" & dataRow(sno) & "'", MsgBoxStyle.Critical)
                End Try

            Next
            If errcnt > 0 Then
                MsgBox("Failed to Export the Data", MsgBoxStyle.Critical)
            End If
            If errcnt = 0 Then
                MsgBox("Exporting the Data Successfully", MsgBoxStyle.Information)
            End If
        End If

    End Function

    Public Function insert_data()
        Dim strline As String
        Dim strfile As String
        strFileName = "C:\AttDBPath\tablemap.ini"

        strfile = "C:\AttDBPath\tablemain.ini"

        If File.Exists(strFileName) = False And File.Exists(strfile) = False Then
            MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
        Else
            Dim strmsg As String
            'strmsg = File.ReadAllText(strFileName)


            'If (IsFileEncrypted = True) Then
            strmsg = Decrypt(strFileName)
            'Else
            '    strmsg = File.ReadAllText(strFileName)
            'End If
            Dim mline() As String = strmsg.Split(Chr(13))

            Dim mlinecount As Integer
            mlinecount = mline.Length
            strline = mline(0)
            Dim tbl1, tbl2 As String
            Dim empcode, direct, sno, loc As String
            Dim logdate As String
            tbl1 = mline(0)
            tbl2 = mline(1).Replace(vbLf, "")
            empcode = mline(2).Replace(vbLf, "")
            logdate = mline(3).Replace(vbLf, "")
            direct = mline(4).Replace(vbLf, "")
            sno = mline(5).Replace(vbLf, "")
            loc = mline(6).Replace(vbLf, "")

            Dim da As SqlDataAdapter
            Dim ds As New DataSet

            Dim daSAL As SqlDataAdapter
            Dim dsSAL As New DataSet
            'DeviceLogs_9_2022_copy
            'Parallal
            Dim strtb As String
            strtb = "select  top 10  * from " & tbl1 & "  where  Flag=0 "
            'order by  " & sno & " asc "
            da = New SqlDataAdapter(strtb, conSrcData)
            'da.Fill(ds, "emplog")

            Try
                If Not da Is Nothing Then
                    da.Fill(ds, "emplog")
                Else
                    MsgBox("There Is no data!", MsgBoxStyle.Critical)
                End If
            Catch ex As Exception
                MsgBox("Invalid Table", MsgBoxStyle.Critical)
                Exit Function
            End Try

            Dim errcnt As Integer
            errcnt = 0

            'UserId, LogDate, Direction, DeviceLogId


            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim dataRow As DataRow = ds.Tables(0).Rows(i)

                Dim cmdTrans As SqlCommand
                cmdTrans = New SqlCommand
                cmdTrans.Connection = conDestData
                Try
                    ''cmdTrans.CommandText = "Insert into Parallel (EmployeeCode, LogDateTime, Direction, SerialNumber)" &
                    ''           "Values ('" & empcode & "','" & logdate & "','" & direct & "','" & sno & "')"

                    'Dim originalDateString As String = dataRow(logdate)
                    'Dim originalDateFormat As String = "dd/MM/yyyy hh:mm:ss tt"
                    'Dim originalDate_Format As String = "dd/MM/yyyy h:mm:ss tt"
                    'Dim desiredDateFormat As String = "MM/dd/yyyy hh:mm:ss tt"
                    'Dim desired_DateFormat As String = "yyyy-MM-dd hh:mm:ss tt"
                    'Dim desiredDateString As String
                    'desiredDateString = dataRow(logdate)
                    '' Parse original string into DateTime object using original format
                    'Dim originalDateTime As DateTime
                    'If DateTime.TryParseExact(originalDateString, originalDateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, originalDateTime) Then
                    '    ' Convert DateTime object back to string using desired format
                    '    desiredDateString = originalDateTime.ToString(desiredDateFormat)
                    '    'MsgBox("Original date string: " & originalDateString)
                    '    'Console.WriteLine("Desired date string: " & desiredDateString)
                    'Else
                    '    'Console.WriteLine("Failed to parse original date string.")
                    'End If

                    'If DateTime.TryParseExact(originalDateString, originalDate_Format, CultureInfo.InvariantCulture, DateTimeStyles.None, originalDateTime) Then
                    '    ' Convert DateTime object back to string using desired format
                    '    desiredDateString = originalDateTime.ToString(desiredDateFormat)
                    '    'MsgBox("Original date string: " & originalDateString)
                    '    'Console.WriteLine("Desired date string: " & desiredDateString)
                    'Else
                    '    'Console.WriteLine("Failed to parse original date string.")
                    'End If




                    'desiredDateString = originalDateTime.ToString(desired_DateFormat)
                    ''MsgBox("Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber,Location)" &
                    ''"Values ('" & dataRow(empcode) & "','" & desiredDateString & "','" & dataRow(direct) & "','" & dataRow(sno) & "','" & dataRow(loc) & "')")

                    'cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber,Location)" &
                    '"Values ('" & dataRow(empcode) & "','" & desiredDateString & "','" & dataRow(direct) & "','" & dataRow(sno) & "','" & dataRow(loc) & "')"

                    Dim formattedDate As String = CType(dataRow(logdate), DateTime).ToString("yyyy-MM-dd HH:mm:ss")


                    cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber,Location)" &
                    "Values ('" & dataRow(empcode) & "','" & formattedDate & "','" & dataRow(direct) & "','" & dataRow(sno) & "','" & dataRow(loc) & "')"





                    cmdTrans.ExecuteNonQuery()
                    cmdTrans.Dispose()

                    'MsgBox("update  " & tbl1 & "  set Flag=1 where " & empcode & "='" & dataRow(empcode) & "' and " & sno & "='" & dataRow(sno) & "'")
                    Dim cmdAlter As New SqlClient.SqlCommand
                    cmdAlter.Connection = conSrcData
                    cmdAlter.CommandText = "update  " & tbl1 & "  set Flag=1 where " & empcode & "='" & dataRow(empcode) & "' and " & sno & "='" & dataRow(sno) & "'"
                    cmdAlter.ExecuteNonQuery()
                    cmdAlter.Dispose()

                Catch ex As Exception
                    errcnt = errcnt + 1
                    'MsgBox("Failed to Export the Data  at Rowno '" & dataRow(sno) & "'", MsgBoxStyle.Critical)
                End Try

            Next
            If errcnt > 0 Then
                MsgBox("Failed to Export the Data", MsgBoxStyle.Critical)
            End If
            If errcnt = 0 Then
                MsgBox("Exporting the Data Successfully", MsgBoxStyle.Information)
            End If
        End If

    End Function


    Public Function startservice()
        Dim appPath As String = Application.StartupPath()
        'Shell(appPath + "\vpn.bat", AppWinStyle.Hide)
        Dim psi As New ProcessStartInfo("C:\Program Files\Default Company Name\Setup1\vpn_connect.exe", AppWinStyle.Hide)
        'Dim psi As New ProcessStartInfo(appPath + "\vpn_connect.exe", AppWinStyle.Hide)
        psi.RedirectStandardError = True
        psi.RedirectStandardOutput = True
        psi.CreateNoWindow = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False
        Dim securePass As New Security.SecureString()
        'Dim pass As String = "password"
        'Dim pass As String = "anu"
        'For Each c As Char In pass
        '    securePass.AppendChar(c)
        'Next
        Dim process1 As Process = Process.Start(psi)

        process1.WaitForExit()
    End Function

    Public Function callservice()



        If File.Exists(Application.StartupPath & "\WS\setup.exe") = False Then
            MsgBox("File is not found to Start the Service", MsgBoxStyle.Information)
        Else
            Try
                Dim strFilePath As String
                strFilePath = Application.StartupPath & "\WS\setup.exe"
                'strFilePath = Application.StartupPath & "\WS\WindowsService3.exe"
                Dim proc = Process.Start(strFilePath)
                proc.WaitForExit()

                'Dim Result As DialogResult = MessageBox.Show("Do you want to close the Application", "caption", MessageBoxButtons.YesNo)
                'If Result = DialogResult.No Then
                'ElseIf Result = DialogResult.Yes Then
                '    Application.Exit()
                'End If

                If System.IO.File.Exists("C:\Program Files (x86)\NGAttUtility\AttUtilityProject\AttendanceService.exe") = True Then
                    MsgBox("Attendance Service Installed Successfully ", MsgBoxStyle.Information)
                    Application.Exit()
                End If



            Catch ex As Exception
                MsgBox("Filed to Start the Service", MsgBoxStyle.Critical)
            End Try
        End If


    End Function
    Private Sub btnws_Click(sender As Object, e As EventArgs) Handles btnws.Click
        'startservice()
        callservice()


    End Sub

    Private Sub frmconfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not Directory.Exists("C:\AttDBPath") Then
            Directory.CreateDirectory("C:\AttDBPath")
        End If
        Button1.PerformClick()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If IsNothing(conDestData) = False Then
            If (conDestData.State = ConnectionState.Open) Then

            Else
                btnconn.PerformClick()
            End If
        Else
            btnconn.PerformClick()
        End If
        If IsNothing(conSrcData) = False Then
            If (conSrcData.State = ConnectionState.Open) Then

            Else
                btnconn1.PerformClick()
            End If
        Else
            btnconn1.PerformClick()
        End If


        If cmbsrc.Text = "" Then
            MsgBox("Select the Source Table", MsgBoxStyle.Information)
            Exit Sub
        End If


        Dim clmnda As SqlDataAdapter
        Dim clmnds As New DataSet

        DGColumn.Columns.Clear()

        clmnda = New SqlDataAdapter("select column_name ColumnName  from INFORMATION_SCHEMA.COLUMNS where table_name = 'Parallel'", conDestData)
        clmnda.Fill(clmnds, "dsclmn")

        DGColumn.AutoGenerateColumns = False

        Dim txc As New DataGridViewTextBoxColumn
        txc.HeaderText = "ColumnName"
        txc.DataPropertyName = "ColumnName"
        txc.Width = "150"
        txc.ReadOnly = True
        DGColumn.Columns.Add(txc)



        clmnds.Tables(0).Rows(0).Delete()
        clmnds.Tables(0).AcceptChanges()

        DGColumn.DataSource = clmnds.Tables("dsclmn")

        Dim sclmnda As SqlDataAdapter
        Dim sclmnds As New DataSet

        sclmnda = New SqlDataAdapter("select column_name from INFORMATION_SCHEMA.COLUMNS where table_name = '" & cmbsrc.Text & "'", conSrcData)
        sclmnda.Fill(sclmnds, "bindclmn")

        'Dim cbo = CType(DGColumn.Columns("MapColumn1"), DataGridViewComboBoxColumn)
        Dim cbo As New DataGridViewComboBoxColumn
        cbo.DataSource = sclmnds.Tables("bindclmn")
        cbo.ValueMember = "column_name"
        cbo.DisplayMember = "column_name"
        cbo.HeaderText = "MapColumn"
        cbo.Width = "300"
        DGColumn.Columns.Add(cbo)


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try

            '****************  Mapped Table   ***************

            If Not Directory.Exists("C:\AttDBPath") Then
                Directory.CreateDirectory("C:\AttDBPath")
            End If
            strFileName = "C:\AttDBPath\tablemap.ini"

            Dim objReader As StreamReader
            Dim objWriter As StreamWriter
            objWriter = New StreamWriter(strFileName)
            objWriter.WriteLine(cmbsrc.Text.ToString())
            objWriter.WriteLine("Parallel")
            For i As Integer = 0 To DGColumn.Rows.Count - 1
                'objWriter.WriteLine(DGColumn.Rows(i).Cells(0).Value.ToString() & "=" & DGColumn.Rows(i).Cells(1).Value.ToString())
                objWriter.WriteLine(DGColumn.Rows(i).Cells(1).Value.ToString())
            Next
            objWriter.Close()
            EnCrypt(strFileName)


            '****************  Main Table   ***************

            If Not Directory.Exists("C:\AttDBPath") Then
                Directory.CreateDirectory("C:\AttDBPath")
            End If
            strFileName1 = "C:\AttDBPath\tablemain.ini"

            Dim objReader1 As StreamReader
            Dim objWriter1 As StreamWriter
            objWriter = New StreamWriter(strFileName1)
            objWriter.WriteLine(cmbsrc.Text.ToString())
            objWriter.WriteLine("Parallel")
            For i As Integer = 0 To DGColumn.Rows.Count - 1
                'objWriter.WriteLine(DGColumn.Rows(i).Cells(0).Value.ToString() & "=" & DGColumn.Rows(i).Cells(1).Value.ToString())
                objWriter.WriteLine(DGColumn.Rows(i).Cells(0).Value.ToString())
            Next
            objWriter.Close()
            EnCrypt(strFileName1)




            MsgBox("Mapped the Table Columns", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Mapped to  Columns Failed", MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub cmbsrc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbsrc.SelectedIndexChanged

    End Sub
End Class