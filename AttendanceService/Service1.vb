Imports System.Data.SqlClient
Imports System.IO
Imports System.IO.Stream
Imports System.Data
Imports System.Timers
Imports System.Windows.Forms
Imports System.Security.Cryptography
Public Class Service1

    Dim mythread As Threading.Thread
    Dim dbthread As Threading.Thread
    Public conDestData As SqlClient.SqlConnection
    Public conSrcData As SqlClient.SqlConnection
    Public strPrivate As String = "<RSAKeyValue><Modulus>2b7LPLGvn+Gj7DBxSBi1pist5MLC5yb8Fo3Aly5fxDuvu2APFiCi4VLsBpWWKZC3ZAGD3JZ/0OLXyF6r2e+oDmxIBTS3ZlEYPqTZXEudoZRw5V5kQg0Uip+vXP65ztT9L9zKSVjf51gwsOUb8QpOJERhlgRYrgqBwfgpZyGhjHc=</Modulus><Exponent>AQAB</Exponent><P>/YHgY7dUragkw0SB5uXPCAxyr878lQXjP4EfrEeQF+8G95ptF2w6HIs06nJSSqiFE4IyEBuQycHTBK+1E5YGyQ==</P><Q>2+Ll3EMjqQrJD2ZTOH5lT7dTZEL6Op7NGQVhfY4+rq91WYwm1J8WHKS3rV7q8rdFC6p2HDGYmN/bTq6xB4pZPw==</Q><DP>hXiA9N9MZRX3LRv/rNrn8tvi8i9viuKLsB7C10jiU8eUin6y2zcvLWIZnSpNq2MolYnh49svkxpKiNgd5U8DCQ==</DP><DQ>mw4HVSkrDlsCqQ9ZA+9tdacq8PqiBZBRxKEcvDMAVKJ5t+mywCBmsVAeDe1u9DT0RWOw4fS/TJ4ewf9B6rVOdQ==</DQ><InverseQ>xTa1Xu+oS22WggRGgTIDCY5DpzjaKNghpyEwwHiqHsO/dGw04pQ7e+6JHXdR5BYJ+mo/kV7D+nk4r7KfeSzyMg==</InverseQ><D>awfDssPMhhRNlQ2CwWOT9mgHGQk68JBTHWr0HdvnqveDu+DNyZylM4ilB9+Dfk7qNjggbs9zaGP4mT8fzfJlcjjJhuqZ64hHQODvifYNDErk9hp8NQuv7WVJCjWtFPH7Ttn20Hj8ToQoxlUPwS2cg6cCS18E2mICXOWF+60M1aE=</D></RSAKeyValue>"
    Public strPublic As String = "<RSAKeyValue><Modulus>2b7LPLGvn+Gj7DBxSBi1pist5MLC5yb8Fo3Aly5fxDuvu2APFiCi4VLsBpWWKZC3ZAGD3JZ/0OLXyF6r2e+oDmxIBTS3ZlEYPqTZXEudoZRw5V5kQg0Uip+vXP65ztT9L9zKSVjf51gwsOUb8QpOJERhlgRYrgqBwfgpZyGhjHc=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>"
    Public StrSvrName As String
    Public StrAuthentication As String
    Public StrSQLUserId As String
    Public StrSQLPwd As String
    Public StrSQLDB As String
    Public StrSvrName1 As String
    Public StrAuthentication1 As String
    Public StrSQLUserId1 As String
    Public StrSQLPwd1 As String
    Public StrSQLDB1 As String
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work

        'Timer1.Interval = 20
        'Timer1.Start()

        'mythread = New Threading.Thread(AddressOf keepcounting)
        'mythread.Start()
        If Not Directory.Exists("C:\AttDBPath") Then
            Directory.CreateDirectory("C:\AttDBPath")
        End If
        destdbconnect()
        sourcedbconnect()
        ''dbconnect()


        System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "starttimer")
        Dim timer As System.Timers.Timer = New System.Timers.Timer()
        timer.Interval = 50000
        timer.Enabled = True
        timer.Start()

        AddHandler timer.Elapsed, AddressOf Me.OnTimer

        'Timer1.Elapsed += New ElapsedEventHandler(OnElapsedTime);

        'dbthread = New Threading.Thread(AddressOf exportdata)
        'dbthread.Start()
        'Threading.Thread.Sleep(50000)
    End Sub
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        ' TODO: Insert Monitoring Activies Here.
        Try
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Hello")
            'exportdata()
            insert_data()
        Catch ex As Exception

        End Try
    End Sub
    Private Shared Sub OnTimedEvent(source As Object, e As ElapsedEventArgs)
        System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Hello World!")
    End Sub
    Protected Sub keepcounting()
        Dim i As Integer = 0
        Do While True
            i = i + 1
            'System.IO.File.WriteAllText("D: \anu_text.txt", i)
            'Threading.Thread.Sleep(10000)

            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", i)
            Threading.Thread.Sleep(50000)
        Loop
    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        'dbthread.Abort()
    End Sub
    Protected Sub dbconnect()
        'dbthread.Abort()
        conSrcData = New SqlClient.SqlConnection("user id=sa;Password=Newgen123;integrated security=SSPI;initial catalog=ESSL;data source=FF-LP-003;MultipleActiveResultSets = true")

        conDestData = New SqlClient.SqlConnection("user id=sa;Password=Newgen123;integrated security=SSPI;initial catalog=TdsPac;data source=FF-LP-003;MultipleActiveResultSets = true")

        conDestData.Open()
        conSrcData.Open()

        'Dim cmdTrans As SqlCommand
        'cmdTrans = New SqlCommand
        'cmdTrans.Connection = conSrcData
        System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "dbconnected")
        'Dim str As String
        'str = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber) Values ('a','" & Now.Date & "','a','a')"
        'cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
        '                   "Values ('a','" & Now.Date & "','a','a')"
        'System.IO.File.WriteAllText("D:\anu_text.txt", str)
        'cmdTrans.ExecuteNonQuery()
        'cmdTrans.Dispose()
        'System.IO.File.WriteAllText("D:\anu_text.txt", "starttimer")
        'Timer1.Enabled = True
        ''Timer1.Interval = 2
        'Timer1.Start()

        'System.IO.File.WriteAllText("D:\anu_text.txt", "startedtimer")
        'exportdata()
    End Sub
    Protected Sub sourcedbconnect()
        Dim strLine As String
        Dim strFileName As String
        Dim strDonglePath As String
        strFileName = "C:\AttDBPath\srcpath.ini"
        'strFileName = Application.StartupPath & "\srcpath.ini"
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
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Source DB Connected")
        Catch ex As Exception
            conSrcData.Close()
            conSrcData.Dispose()
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Source DB Failed")
        End Try

    End Sub
    Protected Sub destdbconnect()
        Dim strFileName As String
        Dim strDonglePath As String
        Dim strLine As String
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

            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Destination DB Connected")

        Catch ex As Exception
            conDestData.Close()
            conDestData.Dispose()

            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "Destination DB Failed")
        End Try

    End Sub
    Public Function CheckEncrypt(ByVal filepath As String)
        Dim eFlag As Boolean = True
        Dim strReadTextFIle As String
        'Dim delimited As String = "[~`!@#$%^&*()-+|{}';,<>?]"
        'For Each match As Match In Regex.Matches(File.ReadAllText(filepath), delimited)
        '    eFlag = True
        'Next
        strReadTextFIle = File.ReadAllText(filepath)
        If (strReadTextFIle.Contains("SERVER")) Then
            eFlag = False
        End If
        Return eFlag
    End Function
    Public Function Decrypt(ByVal Text As String) As String
        Try
            Dim objPublicKeyGenerator = New clsEncrypt.PublicKeyGenerator
            Dim objMyKeyPair = objPublicKeyGenerator.MakeKeyPair
            objMyKeyPair.Publickey.key = strPublic
            objMyKeyPair.PrivateKey.Key = strPrivate
            objPublicKeyGenerator = Nothing
            Dim objCryptographyFunctions = New clsEncrypt.CryptographyFunctions
            Dim strEncMessage, strMessage As String
            Dim sr As StreamReader = File.OpenText(Text)
            Dim input, input1 As String
            input = sr.ReadToEnd
            sr.Close()
            strEncMessage = input
            strMessage = objCryptographyFunctions.DecryptAndAuthenticate(objMyKeyPair, strEncMessage)
            Return strMessage
            Dim objReader As StreamReader
            Dim objWriter As StreamWriter
            Dim strFileName As String
            objWriter = New StreamWriter(Text)
            objWriter.Write(strMessage)
            objWriter.Close()
        Catch ex As Exception
        End Try
    End Function
    Public Function insertdata()
        Dim strline As String
        Dim strFile As String
        Dim strFileName As String
        strFileName = "C:\AttDBPath\tablemap.ini"
        strFile = "C:\AttDBPath\tablemain.ini"

        If File.Exists(strFileName) = False And File.Exists(strFile) = False Then
            'MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
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
            strtb = "select  top 50  * from " & tbl1 & "  where   isnull(Flag,0) = 0  order by  " & sno & ""
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
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "anu")

            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim dataRow As DataRow = ds.Tables(0).Rows(i)

                Dim cmdTrans As SqlCommand
                cmdTrans = New SqlCommand
                cmdTrans.Connection = conDestData
                Dim altstr As String
                Try
                    'cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
                    '           "Values ('" & empcode & "','" & logdate & "','" & direct & "','" & sno & "')"

                    System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", i)
                    Dim strqry As String
                    Dim strcnt As Integer
                    Dim strvalue As String

                    strqry = "Insert into Parallel (" & tblcolumn & ")" & "Values ("
                    For strcnt = 2 To mline1.Length - 2
                        If strcnt <> 5 Then
                            strvalue += "'" + dataRow(mline1(strcnt).Replace(vbLf, "")) + "'" + ","
                        End If
                        If strcnt = 5 Then
                            strvalue += dataRow(mline1(strcnt).Replace(vbLf, "")) & ","
                        End If
                    Next
                    strvalue = strvalue.Remove(strvalue.Length - 1, 1)
                    strqry += strvalue + ")"

                    '('" & dataRow(empcode) & "','" & dataRow(logdate) & "','" & dataRow(direct) & "','" & dataRow(sno) & "')"

                    cmdTrans.CommandText = strqry
                    cmdTrans.ExecuteNonQuery()
                    cmdTrans.Dispose()

                    Dim cmdAlter As New SqlClient.SqlCommand
                    cmdAlter.Connection = conSrcData

                    altstr = "update  " & tbl1 & "  set Flag=1 where " & mline1(2).Replace(vbLf, "") & "='" + dataRow(mline1(2).Replace(vbLf, "")) + "' and " & mline1(5).Replace(vbLf, "") & "= '" + Convert.ToString(dataRow(mline1(5).Replace(vbLf, ""))) + "'"
                    cmdAlter.CommandText = altstr
                    cmdAlter.ExecuteNonQuery()
                    cmdAlter.Dispose()



                Catch ex As Exception
                    System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", altstr)
                End Try
            Next

        End If

    End Function
    Public Function insert_data()
        Dim strline As String
        Dim strFileName As String
        strFileName = "C:\AttDBPath\tablemap.ini"

        If File.Exists(strFileName) = False Then
            'MsgBox("File is not found to Configure the DB Connection", MsgBoxStyle.Information)
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
            strtb = "select  top 50  * from " & tbl1 & "  where  Flag = 0 order by  " & sno & ""
            da = New SqlDataAdapter(strtb, conSrcData)
            'da.Fill(ds, "emplog")
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", strtb)

            Try
                If Not da Is Nothing Then
                    da.Fill(ds, "emplog")
                Else
                    'MsgBox("There is no data!", MsgBoxStyle.Critical)
                End If
            Catch ex As Exception
                'MsgBox("Invalid Table", MsgBoxStyle.Critical)
                Exit Function
            End Try

            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "1")


            'UserId, LogDate, Direction, DeviceLogId


            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim dataRow As DataRow = ds.Tables(0).Rows(i)

                Dim cmdTrans As SqlCommand
                cmdTrans = New SqlCommand
                cmdTrans.Connection = conDestData
                Dim altstr As String
                Try
                    'cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
                    '           "Values ('" & empcode & "','" & logdate & "','" & direct & "','" & sno & "')"

                    cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber,Location)" &
                    "Values ('" & dataRow(empcode) & "','" & dataRow(logdate) & "','" & dataRow(direct) & "','" & dataRow(sno) & "','" & dataRow(loc) & "')"
                    cmdTrans.ExecuteNonQuery()
                    cmdTrans.Dispose()

                    Dim cmdAlter As New SqlClient.SqlCommand
                    cmdAlter.Connection = conSrcData

                    altstr = "update  " & tbl1 & "  set Flag=1 where " & empcode & "='" & dataRow(empcode) & "' and " & sno & "='" & dataRow(sno) & "'"
                    cmdAlter.CommandText = "update  " & tbl1 & "  set Flag=1 where " & empcode & "='" & dataRow(empcode) & "' and " & sno & "='" & dataRow(sno) & "'"
                    cmdAlter.ExecuteNonQuery()
                    cmdAlter.Dispose()
                    'System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", 3)

                Catch ex As Exception
                    System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", altstr)
                End Try

            Next

        End If

    End Function
    Protected Sub exportdata()
        Dim str1 As String = ""
        Dim str2 As String = ""
        Try
            Dim da As SqlDataAdapter
            Dim ds As New DataSet

            Dim daSAL As SqlDataAdapter
            Dim dsSAL As New DataSet

            da = New SqlDataAdapter("select  top 10 * from DeviceLogs_9_2022_copy  where   isnull(Flag,0)=0 order by DeviceLogId", conSrcData)
            da.Fill(ds, "emplog")


            'UserId, LogDate, Direction, DeviceLogId


            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim dataRow As DataRow = ds.Tables(0).Rows(i)

                Dim cmdTrans As SqlCommand
                cmdTrans = New SqlCommand
                cmdTrans.Connection = conDestData

                Try

                    str1 = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
                               "Values ('" & dataRow("UserId") & "','" & dataRow("LogDate") & "','" & dataRow("Direction") & "','" & dataRow("DeviceLogId") & "')"
                    cmdTrans.CommandText = "Insert into Parallel (EmployeeCode,LogDateTime,Direction,SerialNumber)" &
                               "Values ('" & dataRow("UserId") & "','" & dataRow("LogDate") & "','" & dataRow("Direction") & "','" & dataRow("DeviceLogId") & "')"
                    cmdTrans.ExecuteNonQuery()
                    cmdTrans.Dispose()



                    str2 = "update DeviceLogs_9_2022_copy set Flag=1 where UserId='" & dataRow("UserId") & "' and DeviceLogId='" & dataRow("DeviceLogId") & "'"
                    Dim cmdAlter As New SqlClient.SqlCommand
                    cmdAlter.Connection = conSrcData
                    cmdAlter.CommandText = "update DeviceLogs_9_2022_copy set Flag=1 where UserId='" & dataRow("UserId") & "' and DeviceLogId='" & dataRow("DeviceLogId") & "'"
                    cmdAlter.ExecuteNonQuery()
                    cmdAlter.Dispose()

                Catch ex As Exception

                End Try
            Next
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", "datainserted")
        Catch ex As Exception
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", str1)
            System.IO.File.WriteAllText("C:\AttDBPath\AttUtility.txt", str2)
        End Try

    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'System.IO.File.WriteAllText("D:\anu_text.txt", "timer started")
        'exportdata()

        Dim MyEventLog As System.Diagnostics.EventLog

        MyEventLog = New EventLog

        MyEventLog.Source = "timertest"

        MyEventLog.WriteEntry("data timer has fired")
    End Sub
End Class
