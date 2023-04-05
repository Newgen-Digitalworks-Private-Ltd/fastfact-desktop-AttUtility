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
'Ver 2.3.8-2678 End

Module utility

    'API Functions
    '<DllImport("win32dll.dll")> Public Function DogRead(ByVal DogBytes As Int16, ByVal DogAddr As Integer, ByVal DogData As String) As Integer
    'End Function

    'Declare Function DogRead Lib "Win32dll.dll" (ByVal DogBytes As Long, ByVal DogAddr As Long, ByVal DogData As String) As Long
    Declare Function DogRead Lib "Win32dll" (ByVal DogBytes As Integer, ByVal DogAddr As Integer, ByVal DogData As String) As Integer

    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
    As String, ByVal lpFile As String, ByVal lpParameters _
    As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    Public salTDSApplicationPath As String = Application.StartupPath + "\SalTds"
    'ra 1.3.109
    Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Int32
    Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
    'ra1.3.109


    'Global Declaration   

    Public conTdsPac As SqlClient.SqlConnection
    Public conDestData As SqlClient.SqlConnection
    Public conSrcData As SqlClient.SqlConnection
    Public resMessage As Resources.ResourceManager
    Public StrUserId As String
    Public StrCocd As String
    Public strCoName As String
    Public blnModify As Boolean
    Public blnAuditTrail As Boolean
    Public StrFinYear As String
    Public StrDbName As String
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
    Public blnCreate As Boolean
    Public dtOpDt, dtClDt As Date
    Public strDeductorType, strDeductorStatus As String
    Public intYearForm, intYearTo As Integer
    Public strSearchValue As String
    Public strSearchOutputType As String = "Code"
    Public blnLargeData As Boolean = False
    Public str3DType As String
    'For Transaction use
    Public strDefaultDate As String
    Public strGrossingUp As String
    Public strBookEntry As String
    Public strLocalIpAddress As String
    Public blnPrintCoverLetter As Boolean
    Public strEnqType As String
    Public blnCanModify As Boolean = True ' user right option
    Public blnCanDelete As Boolean = True ' ""
    Public blnCanAdd As Boolean = True    ' ""
    Public blnCanPrint As Boolean = True  ' ""   
    Public blnCanSearch As Boolean = True
    Public objPublicKeyGenerator As clsEncrypt.PublicKeyGenerator
    Public objMyKeyPair As clsEncrypt.KeyPair
    Public objCryptographyFunctions As clsEncrypt.CryptographyFunctions
    Public objMyKeyRing As New clsEncrypt.KeyRing
    Public strLevelID As String
    Public drTrans As SqlDataReader
    Public drChallan As SqlDataReader
    'ver 3.06-SoftLicense Start
    'Public DemoVersion As Boolean
    Public IsACFFound As Boolean = True
    Public DemoVersion As Boolean = True
    Public blnIsAsctivatedBySL As Boolean = True
    Public IsCalledFromCommandline As Boolean = False
    Public IsContinueWithCommandLineParam As Boolean = True
    'ver 3.06-SoftLicense End
    Public intSearchRate As Integer
    'ra 1.3.109
    Public blnChkOpenDb As Boolean
    Public strSearchCaption As String
    Public blnconnected As Boolean
    'ra 1.3.109

    Public intTotalRecordForPan As Long
    Public intTotalWrongPan As Long
    Public tsPanFile As System.IO.StreamWriter
    Public strPanFile As String
    'uj
    Public chlnNo As String
    Public strParent As String
    Public strPrivate As String = "<RSAKeyValue><Modulus>2b7LPLGvn+Gj7DBxSBi1pist5MLC5yb8Fo3Aly5fxDuvu2APFiCi4VLsBpWWKZC3ZAGD3JZ/0OLXyF6r2e+oDmxIBTS3ZlEYPqTZXEudoZRw5V5kQg0Uip+vXP65ztT9L9zKSVjf51gwsOUb8QpOJERhlgRYrgqBwfgpZyGhjHc=</Modulus><Exponent>AQAB</Exponent><P>/YHgY7dUragkw0SB5uXPCAxyr878lQXjP4EfrEeQF+8G95ptF2w6HIs06nJSSqiFE4IyEBuQycHTBK+1E5YGyQ==</P><Q>2+Ll3EMjqQrJD2ZTOH5lT7dTZEL6Op7NGQVhfY4+rq91WYwm1J8WHKS3rV7q8rdFC6p2HDGYmN/bTq6xB4pZPw==</Q><DP>hXiA9N9MZRX3LRv/rNrn8tvi8i9viuKLsB7C10jiU8eUin6y2zcvLWIZnSpNq2MolYnh49svkxpKiNgd5U8DCQ==</DP><DQ>mw4HVSkrDlsCqQ9ZA+9tdacq8PqiBZBRxKEcvDMAVKJ5t+mywCBmsVAeDe1u9DT0RWOw4fS/TJ4ewf9B6rVOdQ==</DQ><InverseQ>xTa1Xu+oS22WggRGgTIDCY5DpzjaKNghpyEwwHiqHsO/dGw04pQ7e+6JHXdR5BYJ+mo/kV7D+nk4r7KfeSzyMg==</InverseQ><D>awfDssPMhhRNlQ2CwWOT9mgHGQk68JBTHWr0HdvnqveDu+DNyZylM4ilB9+Dfk7qNjggbs9zaGP4mT8fzfJlcjjJhuqZ64hHQODvifYNDErk9hp8NQuv7WVJCjWtFPH7Ttn20Hj8ToQoxlUPwS2cg6cCS18E2mICXOWF+60M1aE=</D></RSAKeyValue>"
    Public strPublic As String = "<RSAKeyValue><Modulus>2b7LPLGvn+Gj7DBxSBi1pist5MLC5yb8Fo3Aly5fxDuvu2APFiCi4VLsBpWWKZC3ZAGD3JZ/0OLXyF6r2e+oDmxIBTS3ZlEYPqTZXEudoZRw5V5kQg0Uip+vXP65ztT9L9zKSVjf51gwsOUb8QpOJERhlgRYrgqBwfgpZyGhjHc=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>"
    'Ver 2.0.3-Cyn
    Public strLtrType As String
    'ver 2.3.1-1081 start
    Public ResetChallan As Boolean
    'ver 2.3.1-1081 end
    'Ver 2.3.6-E26 Start
    Public blnSplQuery As Boolean
    Public strSplQuery As String
    'Ver 2.3.6-E26 End
    'Ver 2.3.8-Priority 1 Start
    Public blnSinglePartyMultiTrans As Boolean
    'Ver 2.3.8-Priority 1 Start

    'Ver 2.3.8-1719 Start
    Public blnSinglePartyMultiTransAll As Boolean
    'Ver 2.3.8-1719 End
    'Ver 2.4.0-AXIS-Bank Start
    Public blnPrnRestrict As Boolean
    'Ver 2.4.0-AXIS-Bank end
    'Ver 2.4.0-E62 Start
    Const PagePass As String = "B10sai"
    Public TANLicence As Boolean
    'Ver 2.4.0-E62 End

    'Ver 2.4.0-Form27D Start
    Public strFormNo As String
    'Ver 2.4.0-Form27D End

    'Ver 2.4.9-E119 start
    Public mQuarterSelectedForeTds As String
    Public mFormNoForeTds As String
    'Ver 3.2.4.0-REQ840 start
    Public blnRFlagForEtds As Boolean
    'Ver 3.2.4.0-REQ840 end
    Public mCocds As String
    'Ver 2.4.9-E119 end
    'Ver 2.6.4-E631 start
    Public mPenalQtr As String
    Public mPenalFormNo As String
    Public mCocdCondn1 As String
    'Ver 2.6.4-E631 end
    'Ver 2.5.5-E92 Start
    Public TANDeductor As String
    'Ver 2.5.5-E92 End
    'Ver 2.5.9-E371 start 
    Public blnBookEntry As Boolean
    'Ver 2.5.9-E371 end
    Public blnSinglePartyCode As Boolean
    Public strDonglePath As String 'Ver 3.0.0-REQ301 start --'TdsPac SQL Multiuser Dongle Licensing on similar line of FAMS

    Public dbChallanGeneration As String

    'Ver 3.0.4-REQ417 start  'Jitendra
    Public blnCheckDNFDongle As Boolean
    Public DNFFound As Boolean
    Public DNFYear As String
    Public DNFPath As String
    '==> for Digital signer
    Public DongleType As String
    Public DigitalSignerFound As Boolean
    '<== for Digital signer
    'Ver 3.0.4-REQ417 end 

    'ver 3.06-SoftLicense Start
    Public strLICXMLFile As String
    Public blnIsSoftLicenceActivated As Boolean = True
    Public strCustomerCode As String
    Public blnIsSLDemoClicked As Boolean = False
    Public blnIsSALTDSOnly As Boolean = False
    Public blnIsTDSPACOnly As Boolean = False
    Public conSALTDS As SqlClient.SqlConnection
    Public DeductorCount As Integer = 2
    Public DeducteeCount As Integer = 5
    Public intCurrentSelectedSearchIndex As Integer
    Public strTReTdsEmailId As String = "tdspac@fastfacts.co"
    Public strFastFactWebsite As String = "http://www.fastfacts.co.in"
    Public strTaxAccountingWebsite As String = "fastfacts.co.in"
    Public strWebAppVersion As String
    Public Const strVersionHTML As String = "http://www.fastfacts.co.in/Tdspacsql.htm"
    Public blnIsVersionChecked As Boolean = False
    Public blnUsingOldeVersion As Boolean = False
    Public strConfigFile As String = Application.StartupPath + "\\ACF.dll"
    'ver 3.06-SoftLicense end

    'Ver 3.1.7-REQ687 start
    Public IsRunByActiveDirectory As Boolean = False
    Public strADUserID As String
    Public strADPassword As String
    Public conTdsPacStr As String
    Public IsAdminLoginClick As Boolean = False
    Public strADLoginUserID As String
    'Ver 3.1.7-REQ687 end 

    'Ver 3.1.7-REQ687 start
    Public bmpFile As String = Application.StartupPath + "\\Indus.bmp"  'This will check whether the Bmp file exists or not
    ''Ver 3.1.7-REQ687 end


    Public Function ComboFill(ByVal cboCombo As ComboBox, ByVal strField As String, ByVal drName As SqlClient.SqlDataReader, Optional ByVal blnClearCbo As Boolean = True)
        'This is generic function used to fill any combo box using dataReader
        If blnClearCbo = True Then cboCombo.Items.Clear() 'This will clear earlier contents of combobox if any
        While drName.Read      'Loop to fill combobox
            cboCombo.Items.Add(drName(strField))
        End While
        drName.Close()
        cboCombo.SelectedIndex = -1
    End Function
    'Added By Komal Shah
    'To fill User for SalTDS and TDSPac products
    Public Function ComboFill(ByVal cboCombo As ComboBox, ByVal strField As String, ByVal strValue As String, ByVal drName As DataTable, Optional ByVal blnClearCbo As Boolean = True)
        'This is generic function used to fill any combo box using dataReader
        If blnClearCbo = True Then cboCombo.Items.Clear() 'This will clear earlier contents of combobox if any
        'Loop to fill combobox

        cboCombo.DataSource = drName
        cboCombo.DisplayMember = strField
        cboCombo.ValueMember = strValue
        cboCombo.SelectedIndex = -1
    End Function



    Public Function ComboFill(ByVal cboCombo As ComboBox, ByVal strField As String, ByVal strSelect As String, Optional ByVal blnClearCbo As Boolean = True, Optional ByVal strDefault As String = "NoVal")
        Dim cmd As New SqlClient.SqlCommand(strSelect, conTdsPac)
        Dim drName As SqlClient.SqlDataReader = cmd.ExecuteReader
        'This is generic function used to fill any combo box using dataReader
        If blnClearCbo = True Then cboCombo.Items.Clear() 'This will clear earlier contents of combobox if any
        If strDefault <> "NoVal" Then cboCombo.Items.Add(strDefault)
        While drName.Read      'Loop to fill combobox
            cboCombo.Items.Add(drName(strField))
        End While
        drName.Close()
        cmd.Dispose()
        If strDefault <> "NoVal" Then
            cboCombo.SelectedIndex = 0
        Else
            cboCombo.SelectedIndex = -1
        End If

    End Function
    Public Sub FillQtr(ByVal cboCombo As ComboBox)
        cboCombo.Items.Clear()
        cboCombo.Items.Add("Q1")
        cboCombo.Items.Add("Q2")
        cboCombo.Items.Add("Q3")
        cboCombo.Items.Add("Q4")
    End Sub
    Public Function QuarterEndDate(ByVal mQt As String) As Date
        Select Case mQt
            Case "Q1"
                QuarterEndDate = "30/06/" & Year(dtOpDt)
            Case "Q2"
                QuarterEndDate = "30/09/" & Year(dtOpDt)
            Case "Q3"
                QuarterEndDate = "31/12/" & Year(dtOpDt)
            Case "Q4"
                QuarterEndDate = "31/03/" & Year(dtClDt)
        End Select
    End Function
    Public Function QuarterFromDate(ByVal mQt As String) As Date
        Select Case mQt
            Case "Q1"
                QuarterFromDate = "01/04/" & Year(dtOpDt)
            Case "Q2"
                QuarterFromDate = "01/07/" & Year(dtOpDt)
            Case "Q3"
                QuarterFromDate = "01/10/" & Year(dtOpDt)
            Case "Q4"
                QuarterFromDate = "01/01/" & Year(dtClDt)
        End Select
    End Function
    'Public Sub IniSettings(ByVal frm As System.Windows.Forms.Panel, Optional ByVal blnReadonly As Boolean = True, Optional ByVal blnEnable As Boolean = False)
    '    Dim mCtrl As System.Windows.Forms.Control
    '    Dim mCol As System.Drawing.Color
    '    Dim mCtrlTxt As System.Windows.Forms.TextBox
    '    Dim mCtrlNum As Tdspac.NumTextBox
    '    Dim mCtrlMsk As Tdspac.MskDate
    '    Dim mCtrlCode As Tdspac.CodeTextbox

    '    For Each mCtrl In frm.Controls
    '        If TypeOf (mCtrl) Is TextBox Then
    '            mCtrlTxt = mCtrl
    '            mCol = mCtrlTxt.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrlTxt.ReadOnly = blnReadonly
    '            Else
    '                mCtrlTxt.TabStop = False
    '            End If
    '        ElseIf TypeOf (mCtrl) Is NumTextBox Then
    '            mCtrlNum = mCtrl
    '            mCol = mCtrlNum.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrlNum.ReadOnly = blnReadonly
    '            Else
    '                mCtrlNum.TabStop = False
    '            End If
    '        ElseIf TypeOf (mCtrl) Is CodeTextbox Then
    '            mCtrlCode = mCtrl
    '            mCol = mCtrlCode.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrlCode.ReadOnly = blnReadonly
    '            Else
    '                mCtrlCode.ReadOnly = True
    '                mCtrlCode.TabStop = False
    '            End If

    '        ElseIf TypeOf (mCtrl) Is MskDate Then
    '            mCtrlMsk = mCtrl
    '            mCol = mCtrlMsk.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrlMsk.ReadOnly = blnReadonly
    '                'gajanan start
    '                mCtrlMsk.Enabled = blnEnable
    '                'gajanan end
    '            Else
    '                mCtrlMsk.TabStop = False
    '                mCtrlMsk.ReadOnly = True
    '                'gajanan start
    '                mCtrlMsk.Enabled = True
    '                'gajanan end
    '            End If
    '        ElseIf TypeOf (mCtrl) Is ComboBox Then
    '            mCol = mCtrl.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrl.Enabled = blnEnable
    '            Else
    '                mCtrl.Enabled = False
    '            End If
    '            'Ver 2.4.0-AXIS Bank Start
    '        ElseIf TypeOf (mCtrl) Is CheckBox Then
    '            mCol = mCtrl.BackColor
    '            If mCol.ToString <> System.Drawing.SystemColors.Info.ToString Then
    '                mCtrl.Enabled = blnEnable
    '            Else
    '                mCtrl.Enabled = False
    '            End If
    '            'Ver 2.4.0-AXIS Bank End
    '        End If
    '    Next
    'End Sub
    'Public Function fnIsValidPan(ByVal parampanvalue As String) As String
    '    Dim returnValue As String

    '    Try
    '        Using myConnection As New Data.SqlClient.SqlConnection
    '            Using myCommand As New Data.SqlClient.SqlCommand("Select dbo.IsPANVAlid(@PANValue)", conTdsPac)
    '                myCommand.CommandType = CommandType.Text
    '                myCommand.Parameters.Add(New Data.SqlClient.SqlParameter("@PANValue", parampanvalue))
    '                myCommand.CommandTimeout = 0
    '                Try
    '                    returnValue = myCommand.ExecuteScalar()
    '                Catch ex As Exception
    '                    Throw ex
    '                End Try
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    '    Return returnValue
    '    'Using myConnection As New Data.SqlClient.SqlConnection
    '    '    Using myCommand As New Data.SqlClient.SqlCommand("Select dbo.IsPANVAlid(@PANValue)", conTdsPac)
    '    '        myCommand.CommandType = CommandType.Text
    '    '        myCommand.Parameters.Add(New Data.SqlClient.SqlParameter("@PANValue", parampanvalue))
    '    '        myCommand.CommandTimeout = 0
    '    '        Try
    '    '            returnValue = myCommand.ExecuteScalar()
    '    '        Catch ex As Exception
    '    '            Throw ex
    '    '        End Try
    '    '    End Using
    '    'End Using


    'End Function
    Public Function ValidPAN(ByVal strPan As String, Optional ByVal blnFormat As Boolean = True) As String
        Dim strValidChr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim Str4thChr As String = "ABCFGHJLPT"
        Dim strValidNum As String = "1234567890"
        Dim AryPan() As Char = Nothing
        'Returns "VALID" or "V008" for other than 4 th chr validation fails
        'Returns "V009" for 4th character valiation fails
        'PAN validations - first 5 & last char must have alphabets 
        ' remaining should contain numeric


        'Dim returnValue As String
        'Dim parampanvalue As String

        'parampanvalue = strPan

        'Try

        '    Using myCommand As New Data.SqlClient.SqlCommand("Select dbo.IsPANVAlid(@PANValue)", conTdsPac)
        '        myCommand.CommandType = CommandType.Text
        '        myCommand.Parameters.Add(New Data.SqlClient.SqlParameter("@PANValue", parampanvalue))
        '        myCommand.CommandTimeout = 0
        '        returnValue = myCommand.ExecuteScalar()
        '        Return (returnValue)

        '    End Using

        'Catch ex As Exception
        '    Throw ex
        'End Try


        If strPan = Nothing Then
            Return "INVALID"
        ElseIf strPan.Length <> 10 Then
            If blnFormat = False Then
                Return "INVALID"
            Else
                Return "V008"
            End If

        ElseIf UCase(strPan) = "PANAPPLIED" Or UCase(strPan) = "PANNOTAVBL" Or UCase(strPan) = "PANINVALID" Then
            If blnFormat = False Then
                Return "INVALID"
            Else
                Return "VALID"
            End If

        Else
            AryPan = strPan.ToCharArray
            If InStr(strValidChr, AryPan(0).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryPan(1).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryPan(2).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryPan(3).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryPan(4).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryPan(9).ToString, CompareMethod.Text) = 0 Then
                If blnFormat = False Then
                    Return "INVALID"
                Else
                    Return "V008"
                End If
            End If
            If InStr(strValidNum, AryPan(5).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryPan(6).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryPan(7).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryPan(8).ToString, CompareMethod.Text) = 0 Then
                If blnFormat = False Then
                    Return "INVALID"
                Else
                    Return "V008"
                End If
            End If
            If InStr(Str4thChr, AryPan(3).ToString, CompareMethod.Text) = 0 Then
                If blnFormat = False Then
                    Return "INVALID"
                Else
                    Return "V009"
                End If
            End If
            Return "VALID"
        End If
    End Function
    Public Function IsEmpty(ByVal ctlName As Control, ByVal strValue As String, ByVal strMsg As String, ByVal blnFocus As Boolean) As Boolean
        'To check whether the control is empty  
        'For eg text incase of textbox or selectedItem in case of combobox
        If strValue = "" Or strValue = "//" Then
            MsgBox(strMsg, MsgBoxStyle.Information)
            If blnFocus = True Then
                ctlName.Focus()
            End If
            Return True
        End If
        Return False
    End Function


    Public Function NumCheck(ByVal chrValue As Char) As Boolean
        'This function will check whether input data is numeric 
        Dim intVal As Integer
        intVal = Asc(chrValue) 'This will return ascii value of chrValue
        If intVal = 8 Or intVal = 46 Or (intVal > 47 And intVal < 58) Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function RetrieveData(ByVal strField As Object, Optional ByVal strDefaultValue As String = "Null")
        'This function will check whether retrieved data is null
        If IsDBNull(strField) Then
            If strDefaultValue <> "Null" Then
                Return strDefaultValue
            End If
        Else
            Return strField
        End If
    End Function
    Public Function LabelFill(ByVal strTblName As String, ByVal strCodeField As String, ByVal strNameField As String, ByVal strCodeValue As String) As String
        Try
            Dim StrTCmd As String
            Dim dsTCmd As SqlClient.SqlCommand
            Dim drName As SqlClient.SqlDataReader
            StrTCmd = "select " & strNameField & " from " & strTblName & " where " & strCodeField & "='" & strCodeValue & "'"
            dsTCmd = New SqlClient.SqlCommand(StrTCmd, conTdsPac)
            dsTCmd.CommandType = CommandType.Text
            dsTCmd.CommandTimeout = 0
            drName = dsTCmd.ExecuteReader()
            If drName.HasRows Then
                drName.Read()
                StrTCmd = drName.Item(0)
                If drName.IsClosed = False Then drName.Close()
                Return StrTCmd
            Else
                If drName.IsClosed = False Then drName.Close()
                Return ""
            End If
            drName.Close()
            dsTCmd.Dispose()
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Function
    Public Sub AddToDataset(ByVal ds As DataSet, ByVal strQry As String)
        Try
            Dim strName() As String = Nothing
            Dim strValue() As String = Nothing
            Dim strNameString As String
            Dim strValueString As String
            Dim tableName As String
            Dim position1, position2 As Integer
            Dim i As Integer
            position1 = InStr(strQry, "(")
            position2 = InStr(strQry, ")")

            tableName = Mid(strQry, 13, InStr(strQry, "(") - 14) 'table name
            strNameString = Mid(strQry, position1 + 1, position2 - position1 - 1) 'fields name
            strName = strNameString.Split(",")
            strQry = strQry.Remove(position1 - 1, position2 - position1 + 1)

            position1 = InStr(strQry, "(")
            position2 = InStr(strQry, ")")
            strValueString = Mid(strQry, position1 + 1, position2 - position1 - 1)
            strValue = strValueString.Split(",")
            For i = 0 To strName.Length - 1
                strValue(i) = strValue(i).Trim()
            Next
            'removing quotes


            For i = 0 To strValue.Length - 1
                If (Left(strValue(i), 1) = "'") Then
                    strValue(i) = strValue(i).Remove(0, 1)
                End If
                If (Right(strValue(i), 1) = "'") Then
                    strValue(i) = strValue(i).Remove(strValue(i).Length - 1, 1)
                End If
                strValue(i) = strValue(i).Replace("''", "'")
                strValue(i) = strValue(i).Replace("|", ",")
                strValue(i) = strValue(i).Replace("^", "(")
                strValue(i) = strValue(i).Replace("$", ")")

            Next

            For i = 0 To strName.Length - 1
                strName(i) = strName(i).Trim
                strValue(i) = strValue(i).Trim()
            Next
            'Finding the data type
            Dim mFname As String

            For i = 0 To strValue.Length - 1
                mFname = strName(i).ToString
                If ds.Tables(0).Columns(mFname).DataType.ToString = "System.Boolean" Then
                    strValue(i) = IIf((strValue(i) = "1"), True, False)
                ElseIf ds.Tables(0).Columns(mFname).DataType.ToString = "System.DateTime" Then
                    Dim strDateVal As String
                    strDateVal = Trim((strValue(i)))
                    If strDateVal.ToUpper <> "NULL" Then
                        strValue(i) = (Mid(strDateVal, 4, 2) & "/" & Mid(strDateVal, 1, 2) & "/" & Mid(strDateVal, 7, 4))
                    Else
                        strValue(i) = "Null"
                    End If
                End If
            Next


            Dim insertRow As DataRow = ds.Tables(tableName).NewRow()
            For i = 0 To strValue.Length - 1
                mFname = strName(i).ToString
                If strValue(i) <> "Null" Then
                    If mFname = "Active" Then
                        insertRow(mFname) = IIf(Trim(strValue(i)) = "Y", "Yes", "No")
                    Else
                        insertRow(mFname) = strValue(i)
                    End If
                End If
            Next
            ds.Tables(tableName).Rows.Add(insertRow)
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Sub
    Public Sub DeleteFromDataset(ByVal dsName As DataSet, ByVal tableName As String, ByVal intCurPosition As Integer)
        dsName.Tables(tableName).Rows.RemoveAt(intCurPosition)
    End Sub
#Region " Update to Dataset"

    Public Sub UpdateToDataset(ByVal dsName As DataSet, ByVal strQueryName As String, ByVal mIntCurPosition As Integer)
        Try
            Dim strName(100) As String
            Dim strValue(100) As String
            Dim strFieldValue() As String = Nothing
            Dim strNameString As String
            Dim strValueString As String
            Dim tableName As String
            Dim position1, position2 As Integer
            Dim str As String
            Dim strCondition As String
            Dim strc() As String
            Dim intCount As Integer = 0
            Dim i As Integer = 0
            str = strQueryName
            position1 = InStr(strQueryName, " ")
            str = str.Remove(position1 - 1, 1)
            str = str.Insert(position1 - 1, "@")
            position2 = str.IndexOf(" ")
            tableName = str.Substring(position1, position2 - position1)
            str = str.Remove(0, position2 + 4)
            str = str.Remove(0, 1)

            position1 = str.IndexOf("where")
            strValueString = str.Substring(0, position1 - 1)
            strCondition = str.Remove(0, position1 + 6)
            strFieldValue = strValueString.Split(",")

            strCondition = strCondition.Replace("and", ",")
            strCondition = strCondition.Replace("'", " ")

            strc = strCondition.Split(",")

            For i = 0 To strc.Length - 1
                position1 = InStr(strc(i), "=")
                strName(intCount) = Mid(strc(i), 1, position1 - 1).Trim
                strValue(intCount) = Mid(strc(i), position1 + 1).Trim
                strValue(intCount) = Mid(strc(i), position1 + 1).Trim
                intCount = intCount + 1
            Next i

            For i = 0 To strFieldValue.Length - 1
                position1 = InStr(strFieldValue(i), "=")
                strName(intCount) = Mid(strFieldValue(i), 1, position1 - 1).Trim
                strValue(intCount) = Mid(strFieldValue(i), position1 + 1).Trim
                intCount = intCount + 1
            Next

            For i = 0 To intCount - 1
                If (Left(strValue(i), 1) = "'") Then
                    strValue(i) = strValue(i).Remove(0, 1)
                End If
                If (Right(strValue(i), 1) = "'") Then
                    strValue(i) = strValue(i).Remove(strValue(i).Length - 1, 1)
                End If
                strValue(i) = strValue(i).Replace("''", "'")
                strValue(i) = strValue(i).Replace("|", ",")
                strValue(i) = strValue(i).Replace("^", "(")
                strValue(i) = strValue(i).Replace("$", ")")
            Next
            For i = 0 To intCount - 1
                strName(i) = strName(i).Trim
                strValue(i) = strValue(i).Trim()
            Next

            Dim mFname As String
            dsName.Tables(tableName).Rows(mIntCurPosition).BeginEdit()
            For i = 0 To intCount - 1
                mFname = strName(i).ToString
                'Start 07-10-2010
                'If dsName.Tables(0).Columns(mFname).DataType.ToString <> Nothing Then
                'End 07-10-2010
                If dsName.Tables(0).Columns(mFname).DataType.ToString = "System.Boolean" Then
                    dsName.Tables(tableName).Rows(mIntCurPosition).Item(mFname) = IIf((strValue(i) = "1"), True, False)
                ElseIf dsName.Tables(0).Columns(mFname).DataType.ToString = "System.DateTime" Then
                    Dim strDateVal As String
                    strDateVal = Trim((strValue(i)))
                    If strDateVal.ToUpper <> "NULL" Then
                        dsName.Tables(tableName).Rows(mIntCurPosition).Item(mFname) = (Mid(strDateVal, 4, 2) & "/" & Mid(strDateVal, 1, 2) & "/" & Mid(strDateVal, 7, 4))
                    Else
                        strValue(i) = "Null"
                    End If
                ElseIf dsName.Tables(0).Columns(mFname).DataType.ToString = "System.Decimal" Then
                    dsName.Tables(tableName).Rows(mIntCurPosition).Item(mFname) = IIf((strValue(i).ToUpper = "NULL"), "NULL", strValue(i))
                Else
                    If mFname = "Active" Then
                        dsName.Tables(tableName).Rows(mIntCurPosition).Item(mFname) = IIf(Trim(strValue(i)) = "Y", "Yes", "No")
                    Else
                        dsName.Tables(tableName).Rows(mIntCurPosition).Item(mFname) = Trim(strValue(i))
                    End If
                End If
                'Start 07-10-2010
                'End If
                'End 07-10-2010
            Next i
            dsName.Tables(tableName).Rows(mIntCurPosition).EndEdit()
            dsName.Tables(tableName).AcceptChanges()

        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Sub
#End Region
    Public Function ValidEmailId(ByVal strEmailid As String) As Boolean
        Dim AryString As String() = Nothing
        Dim mEmailid As String
        Dim i As Integer
        Try
            'Check Printable chars except space & ^
            If InStr(strEmailid.Trim(), " ", CompareMethod.Text) > 0 Or InStr(strEmailid.Trim(), "^", CompareMethod.Text) > 0 Then Return False
            'Split with @ 
            AryString = strEmailid.Split(New [Char]() {"@"c})
            If (AryString.GetUpperBound(0)) = 1 Then
                ' @ should not preceded by dot without any chars
                If AryString(0).ToString.Trim() = "" Or AryString(0).ToString.Trim() = "." Then Return False
                ' @ preceded & succeeded by atleast one char
                For i = 1 To AryString.GetUpperBound(0)
                    If AryString(i).ToString.Trim() = "" Then
                        Return False
                    End If
                Next
                mEmailid = AryString(1).ToString
                AryString = Nothing
                'Split with . 
                AryString = mEmailid.Split(New [Char]() {"."c})
                If (AryString.GetUpperBound(0)) >= 1 Then
                    If AryString(0).ToString.Trim() = "" Then Return False
                    ' . preceded & succeeded by atleast one char & eMail id should not end with .
                    For i = 1 To AryString.GetUpperBound(0)
                        If AryString(i).ToString.Trim() = "" Then
                            Return False
                        End If
                    Next
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Function
    Public Function ValidTAN(ByVal strTan As String) As Boolean
        Dim strValidChr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim strValidNum As String = "1234567890"
        Dim AryTan() As Char = Nothing

        If strTan.Length <> 10 Then
            Return False
        Else
            AryTan = strTan.ToCharArray
            If InStr(strValidChr, AryTan(0).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryTan(1).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryTan(2).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryTan(3).ToString, CompareMethod.Text) = 0 Or InStr(strValidChr, AryTan(9).ToString, CompareMethod.Text) = 0 Then
                Return False
            End If
            If InStr(strValidNum, AryTan(4).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryTan(5).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryTan(6).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryTan(7).ToString, CompareMethod.Text) = 0 Or InStr(strValidNum, AryTan(8).ToString, CompareMethod.Text) = 0 Then
                Return False
            End If
            Return True
        End If
    End Function
#Region " State"


    Public Sub FillState(ByVal cboName As ComboBox)

        cboName.Items.Add("ANDAMAN AND NICOBAR ISLANDS")
        cboName.Items.Add("ANDHRA PRADESH")
        cboName.Items.Add("ARUNACHAL PRADESH")
        cboName.Items.Add("ASSAM")
        cboName.Items.Add("BIHAR")
        cboName.Items.Add("CHANDIGARH")
        cboName.Items.Add("DADRA & NAGAR HAVELI")
        cboName.Items.Add("DAMAN & DIU")
        cboName.Items.Add("DELHI")
        cboName.Items.Add("GOA")
        cboName.Items.Add("GUJARAT")
        cboName.Items.Add("HARYANA")
        cboName.Items.Add("HIMACHAL PRADESH")
        cboName.Items.Add("JAMMU & KASHMIR")
        cboName.Items.Add("KARNATAKA")
        cboName.Items.Add("KERALA")
        cboName.Items.Add("LAKSHADWEEP")
        cboName.Items.Add("MADHYA PRADESH")
        cboName.Items.Add("MAHARASHTRA")
        cboName.Items.Add("MANIPUR")
        cboName.Items.Add("MEGHALAYA")
        cboName.Items.Add("MIZORAM")
        cboName.Items.Add("NAGALAND")
        cboName.Items.Add("ORISSA")
        cboName.Items.Add("PONDICHERRY")
        cboName.Items.Add("PUNJAB")
        cboName.Items.Add("RAJASTHAN")
        cboName.Items.Add("SIKKIM")
        cboName.Items.Add("TAMILNADU")
        cboName.Items.Add("TRIPURA")
        cboName.Items.Add("UTTAR PRADESH")
        cboName.Items.Add("WEST BENGAL")
        'cboName.Items.Add("CHHATTISGARH") 'ver 2.1.22
        cboName.Items.Add("CHHATISHGARH")
        cboName.Items.Add("UTTARANCHAL")
        'Ver 2.3.6-1513 Start
        'cboName.Items.Add("JHARLHAND")
        cboName.Items.Add("JHARKHAND")
        'Ver 2.3.6-1513 End
        'Ver 3.0.2-REQ349 start
        cboName.Items.Add("TELANGANA")
        'Ver 3.0.2-REQ349 end 
        'Ver 3.2.93 Mentis 0164939 start
        cboName.Items.Add("LADAKH")
        'Ver 3.2.93 Mentis 0164939 end 
        cboName.Items.Add("OTHERS")
    End Sub
#End Region

    Public Sub FillStateSALTDS(ByVal cboBox As ComboBox, Optional ByVal blnClearCbo As Boolean = True)
        Dim stateTable As New DataTable
        Using da As New SqlDataAdapter(" select StateName , StateID from MDM.State where CountryID = 1 ", conSALTDS)
            da.Fill(stateTable)
        End Using
        If blnClearCbo = True Then cboBox.Items.Clear() 'This will clear earlier contents of combobox if any

        cboBox.DisplayMember = "StateName"
        cboBox.ValueMember = "StateID"
        'If Not cboBox.Items.Contains("--Select--") Then
        '    cboBox.Items.Insert(0, New ListViewItem("--Select--"))
        'End If
        Dim dr As DataRow = stateTable.NewRow()
        dr("StateName") = "--Select--"
        dr("StateID") = "0"
        stateTable.Rows.InsertAt(dr, 0)
        cboBox.DataSource = stateTable
        cboBox.SelectedIndex = 0

    End Sub

    Public Sub FillCity(ByVal cboCity As ComboBox, ByVal StateID As Integer, Optional ByVal blnClearCbo As Boolean = True)
        Dim cityTable As New DataTable
        Using da As New SqlDataAdapter(" select CityName, CityID from MDM.City where StateID = '" & StateID & "'", conSALTDS)
            da.Fill(cityTable)
        End Using
        'This will clear earlier contents of combobox if any

        cboCity.DisplayMember = "CityName"
        cboCity.ValueMember = "CityID"
        'If Not cboCity.Items.Contains("--Select--") Then
        '    cboCity.Items.Insert(0, New ListViewItem("--Select--"))
        'End If
        Dim dr As DataRow = cityTable.NewRow()
        dr("CityName") = "--Select--"
        dr("CityID") = "0"
        cityTable.Rows.InsertAt(dr, 0)
        cboCity.DataSource = cityTable
        cboCity.SelectedIndex = 0

    End Sub


    Public Sub ErrHandler(ByVal strErr As Exception)
        MsgBox(strErr.Message)
        MsgBox(strErr.StackTrace)
    End Sub
    Public Sub AddToTdsCode(ByVal mFileName As String, ByVal mTableName As String)
        Dim objReader As StreamReader
        Dim strFilePath As String
        Dim strLine As String
        Dim cmdCommand As New SqlClient.SqlCommand
        Try
            strFilePath = Application.StartupPath & "\" & mFileName
            If Not File.Exists(strFilePath) Then
                MessageBox.Show("Please contact FastFacts. " & mFileName & " file not found.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            cmdCommand.Connection = conTdsPac
            'ver 2.0
            'ver 2.0.3-15
            If mTableName = "SetDisplay" Then
                cmdCommand.CommandText = "select count(*) from " & mTableName & " where cocd='" & StrCocd & "'"
            Else
                cmdCommand.CommandText = "select count(*) from " & mTableName
            End If
            'ver 2.0.3-15
            If cmdCommand.ExecuteScalar = 0 Then

                objReader = New StreamReader(strFilePath)
                'objReader = New StreamReader(mFileName)
                strLine = Replace(Replace(Replace(objReader.ReadLine, """", "'"), "True", "1"), "False", "0")
                'ver 2.0.3-15
                If mTableName = "SetDisplay" Then
                    strLine = Replace(strLine, "DEMO", StrCocd)
                End If

                Do Until strLine = ""
                    cmdCommand.CommandText = "Insert Into " & mTableName & " values (" & strLine & ")"
                    cmdCommand.ExecuteNonQuery()
                    strLine = Replace(Replace(Replace(objReader.ReadLine, """", "'"), "True", "1"), "False", "0")
                    'ver 2.0.3-15
                    If mTableName = "SetDisplay" Then
                        strLine = Replace(strLine, "DEMO", StrCocd)
                    End If
                Loop
            End If
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Sub
    Public Function OpenNewConnection() As SqlClient.SqlConnection
        If StrAuthentication = "Windows" Then
            OpenNewConnection = New SqlClient.SqlConnection("user id=" & StrSQLUserId & ";Password=" & StrSQLPwd & ";integrated security=SSPI;initial catalog=" & StrDbName & ";data source=" & StrSvrName & ";MultipleActiveResultSets=True;")
        Else
            OpenNewConnection = New SqlClient.SqlConnection("user id=" & StrSQLUserId & ";Password=" & StrSQLPwd & ";initial catalog=" & StrDbName & ";data source=" & StrSvrName & ";MultipleActiveResultSets=True;")
        End If

    End Function
    Public Sub GetRights(ByVal strMenu As String)
        'admin will have all rights
        If UCase(StrUserId) = "ADMIN" Then
            blnCanAdd = True
            blnCanModify = True
            blnCanDelete = True
            blnCanPrint = True
            Exit Sub
        End If
        ' if not admin, check for the rights

        Dim cmdUR As New SqlClient.SqlCommand("Select * from UserRights where cocd ='" & StrCocd & "' and DBCode ='" & StrDbName & "' and UserId='" & StrUserId & "' and MenuItem ='" & strMenu & "'", conDestData)
        Dim dr As SqlClient.SqlDataReader

        dr = cmdUR.ExecuteReader

        If dr.Read = True Then
            blnCanAdd = dr("UpdateOptA")
            blnCanModify = dr("UpdateOptM")
            blnCanDelete = dr("UpdateOptD")
            blnCanPrint = dr("UpdateOptP")
        Else
            blnCanAdd = False
            blnCanModify = False
            blnCanDelete = False
            blnCanPrint = False
        End If
        dr.Close()
        cmdUR.Dispose()
    End Sub
    Public Function ChangeToDate(ByVal strDt As String) As String
        If strDt = "//" Then
            Return "Null"
        Else
            Return "'" & strDt.Substring(3, 2) & "/" & strDt.Substring(0, 2) & "/" & strDt.Substring(6, 4) & "'"
        End If

    End Function
#Region " Amount in Words"
    Public Function AmountInWords(ByVal value As Double) As String

        Try

            Dim amt As Double

            amt = value
            Dim amtword As String
            Dim words(100) As String
            If amt = 0 Then
                amtword = " "

            End If
            Dim temp_amt As Double
            Dim twrd As String
            temp_amt = 0
            twrd = ""

            words(1) = "One"
            words(2) = "Two"
            words(3) = "Three"
            words(4) = "Four"
            words(5) = "Five"
            words(6) = "Six"
            words(7) = "Seven"
            words(8) = "Eight"
            words(9) = "Nine"
            words(10) = "Ten"
            words(11) = "Eleven"
            words(12) = "Twelve"
            words(13) = "Thirteen"
            words(14) = "Fourteen"
            words(15) = "Fifteen"
            words(16) = "Sixteen"
            words(17) = "Seventeen"
            words(18) = "Eighteen"
            words(19) = "Nineteen"
            words(20) = "Twenty"
            words(30) = "Thirty"
            words(40) = "Forty"
            words(50) = "Fifty"
            words(60) = "Sixty"
            words(70) = "Seventy"
            words(80) = "Eighty"
            words(90) = "Ninety"

            temp_amt = amt

            amt = Int(amt)
            Do While amt <> 0
                If amt >= 1000000000 And (amt / 1000000000) <= 20 And amt < 100000000000.0# Then
                    twrd = twrd + words(Int(amt / 1000000000)) + " Million "
                    amt = amt Mod 1000000000
                End If
                If amt >= 1000000000 And (amt / 1000000000) > 20 And amt < 100000000000.0# Then
                    twrd = twrd + Words_fun(Int(amt / 1000000000)) + " Million "
                    amt = amt Mod 1000000000
                End If
                If amt >= 10000000 And (amt / 10000000) <= 20 And amt < 1000000000 Then
                    twrd = twrd + words(Int(amt / 10000000)) + " Crore "
                    amt = amt Mod 10000000
                End If
                If amt >= 10000000 And (amt / 10000000) > 20 And amt < 1000000000 Then
                    twrd = twrd + Words_fun(Int(amt / 10000000)) + " Crore "
                    amt = amt Mod 10000000
                End If
                If amt >= 100000 And (amt / 100000) <= 20 And amt < 10000000 Then
                    twrd = twrd + words(Int(amt / 100000)) + " Lakh "
                    amt = amt Mod 100000
                End If
                If amt >= 100000 And (amt / 100000) > 20 And amt < 10000000 Then
                    twrd = twrd + Words_fun(Int(amt / 100000)) + " Lakh "
                    amt = amt Mod 100000


                End If
                If amt >= 1000 And (amt / 1000) <= 20 And amt < 100000 Then
                    twrd = twrd + words(Int(amt / 1000)) + " Thousand "
                    amt = amt Mod 1000

                End If
                If amt >= 1000 And (amt / 1000) > 20 And amt < 100000 Then
                    twrd = twrd + Words_fun(Int(amt / 1000)) + " Thousand "
                    amt = amt Mod 1000

                End If
                If amt >= 100 And amt < 1000 Then
                    twrd = twrd + words(Int(amt / 100)) + " Hundred "
                    amt = amt Mod 100
                End If
                If amt >= 20 And amt < 100 Then
                    twrd = twrd + words(Int(amt / 10) * 10) + " "
                    amt = amt Mod 10
                End If
                If amt > 0 And amt < 20 Then
                    twrd = twrd + words(Int(amt)) + " "
                    amt = 0
                End If
                If amt >= 20 And amt <> 0 Then
                    twrd = twrd + words(Int(amt)) + " "
                    amt = 0
                End If
            Loop
            Dim deci_val As Integer
            deci_val = (temp_amt - Int(temp_amt)) * 100

            If deci_val = 0 Then
                amtword = twrd + " Only "
            Else
                amtword = twrd + " And Paise " + Words_fun(deci_val) + " Only "
            End If
            Return amtword
        Catch ex As Exception

        End Try
    End Function
    Public Function Words_fun(ByVal val As Integer)
        Dim words(100) As String
        words(1) = "One"
        words(2) = "Two"
        words(3) = "Three"
        words(4) = "Four"
        words(5) = "Five"
        words(6) = "Six"
        words(7) = "Seven"
        words(8) = "Eight"
        words(9) = "Nine"
        words(10) = "Ten"
        words(11) = "Eleven"
        words(12) = "Twelve"
        words(13) = "Thirteen"
        words(14) = "Fourteen"
        words(15) = "Fifteen"
        words(16) = "Sixteen"
        words(17) = "Seventeen"
        words(18) = "Eighteen"
        words(19) = "Nineteen"
        words(20) = "Twenty"
        words(30) = "Thirty"
        words(40) = "Forty"
        words(50) = "Fifty"
        words(60) = "Sixty"
        words(70) = "Seventy"
        words(80) = "Eighty"
        words(90) = "Ninety"

        Dim twrd1 As String
        twrd1 = ""

        If val < 100 And val > 20 Then
            twrd1 = twrd1 + words(Int(val / 10) * 10) + " "
            val = val Mod 10

        End If
        val = Int(val)

        If val <= 20 And val <> 0 Then
            twrd1 = twrd1 + words(Int(val)) + " "

        End If
        Words_fun = twrd1
    End Function

#End Region
    Public Function ConvertDataReaderToDataSet(ByVal reader As SqlClient.SqlDataReader) As DataSet
        Try
            Dim dataSet As DataSet = New DataSet
            Dim schemaTable As DataTable = reader.GetSchemaTable()
            Dim dataTable As DataTable = New DataTable("myTable")
            Dim intCounter As Integer
            For intCounter = 0 To schemaTable.Rows.Count - 1
                Dim dataRow As DataRow = schemaTable.Rows(intCounter)
                Dim columnName As String = dataRow("ColumnName")
                Dim column As DataColumn = New DataColumn(columnName)
                dataTable.Columns.Add(column)
            Next

            dataSet.Tables.Add(dataTable)
            While reader.Read()
                Dim dataRow As DataRow = dataTable.NewRow()
                For intCounter = 0 To reader.FieldCount - 1
                    Dim str As String = reader.GetValue(intCounter).ToString
                    If str = "" Then
                        dataRow(intCounter) = ""
                    Else
                        dataRow(intCounter) = reader.GetValue(intCounter)
                    End If
                Next
                dataTable.Rows.Add(dataRow)
            End While
            Return dataSet

        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Function



    Public Function EnCrypt(ByVal Text As String) As String
        Dim objWriter As StreamWriter
        Try
            Dim sr As StreamReader = File.OpenText(Text)
            Dim s, strWords, strTemp As String
            Dim intasc, intTemp As Integer
            Dim input As String
            Dim strEncMessage, strMessage As String
            strWords = sr.ReadToEnd
            sr.Close()

            Dim objPublicKeyGenerator = New clsEncrypt.PublicKeyGenerator
            Dim objMyKeyPair = objPublicKeyGenerator.MakeKeyPair
            objPublicKeyGenerator = Nothing
            objMyKeyPair.PublicKey.Key = strPublic
            objMyKeyPair.PrivateKey.Key = strPrivate
            objMyKeyPair.PrivateKey.Label = "My Private Key"
            objMyKeyPair.PublicKey.Label = "My Public Key"
            Dim objCryptographyFunctions = New clsEncrypt.CryptographyFunctions
            strPublic = objMyKeyPair.PublicKey.Key
            strPrivate = objMyKeyPair.PrivateKey.Key
            Dim objMyKeyRing = New clsEncrypt.KeyRing
            objMyKeyRing.AddKeyPair(objMyKeyPair)
            strMessage = strWords

            strEncMessage = objCryptographyFunctions.SignAndEncrypt(objMyKeyPair, strMessage)
            Dim objReader As StreamReader

            Dim strFileName As String
            objWriter = New StreamWriter(Text)
            objWriter.Write(strEncMessage)
            objWriter.Close()
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try

    End Function
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
            Call ErrHandler(ex)
        End Try
    End Function
    Function StateCode(ByVal StrState As String) As String
        Select Case StrState
            Case "ANDAMAN AND NICOBAR ISLANDS"
                StateCode = "01"
            Case "ANDHRA PRADESH"
                StateCode = "02"
            Case "ARUNACHAL PRADESH"
                StateCode = "03"
            Case "ASSAM"
                StateCode = "04"
            Case "BIHAR"
                StateCode = "05"
            Case "CHANDIGARH"
                StateCode = "06"
            Case "DADRA & NAGAR HAVELI"
                StateCode = "07"
            Case "DAMAN & DIU"
                StateCode = "08"
            Case "DELHI"
                StateCode = "09"
            Case "GOA"
                StateCode = "10"
            Case "GUJARAT"
                StateCode = "11"
            Case "HARYANA"
                StateCode = "12"
            Case "HIMACHAL PRADESH"
                StateCode = "13"
            Case "JAMMU & KASHMIR"
                StateCode = "14"
            Case "KARNATAKA"
                StateCode = "15"
            Case "KERALA"
                StateCode = "16"
            Case "LAKHSWADEEP"
                StateCode = "17"
            Case "MADHYA PRADESH"
                StateCode = "18"
            Case "MAHARASHTRA"
                StateCode = "19"
            Case "MANPUR"
                StateCode = "20"
            Case "MEGHALAYA"
                StateCode = "21"
            Case "MIZORAM"
                StateCode = "22"
            Case "NAGALAND"
                StateCode = "23"
            Case "ORISSA"
                StateCode = "24"
            Case "PONDICHERRY"
                StateCode = "25"
            Case "PUNJAB"
                StateCode = "26"
            Case "RAJASTHAN"
                StateCode = "27"
            Case "SIKKIM"
                StateCode = "28"
            Case "TAMILNADU"
                StateCode = "29"
            Case "TRIPURA"
                StateCode = "30"
            Case "UTTAR PRADESH"
                StateCode = "31"
            Case "WEST BENGAL"
                StateCode = "32"
            Case "CHHATISHGARH", "CHHATTISGARH" 'ver 2.1.22
                StateCode = "33"
            Case "UTTARANCHAL"
                StateCode = "34"
            Case "JHARKHAND"
                StateCode = "35"
                'Ver 3.0.2-REQ349 start
            Case "TELANGANA"
                StateCode = "36"
                'Ver 3.0.2-REQ349 end 
                'Ver 3.2.93 Mentis-0164939 start
            Case "LADAKH"
                StateCode = "37"
                'Ver 3.2.93 Mentis-0164939  end 
            Case "99-FOR NRI"
                StateCode = "99"
            Case Else
                StateCode = "00"
        End Select
    End Function
    Function GetNullVal(ByVal mVal As Object) As Double
        Try
            GetNullVal = (IIf(IsDBNull(mVal), 0, mVal))
            Exit Function
        Catch ex As Exception
            GetNullVal = 0
        End Try
    End Function
    'ra 1.3.109
    Sub ChangeDate()
        Try
            Dim lngLocale As Int32
            lngLocale = GetSystemDefaultLCID()
            If SetLocaleInfo(lngLocale, &H1F, "dd/MM/yyyy") = False Then
            End If
            MsgBox("Date Format Changed Successfully. Please run Tdspac again", vbInformation)
            End
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'ra 1.3.109
    'ra1.3.109
    Public Function AddColmnInTable(ByVal strTable As String, ByVal strField As String, ByVal strFieldType As String, ByVal conTable As SqlConnection)
        Dim cmdTable As New SqlCommand
        cmdTable.Connection = conTable
        Try
            cmdTable.CommandText = "select   top 1 " & strField & " from " & strTable & ""
            cmdTable.ExecuteNonQuery()
        Catch ex As Exception
            cmdTable.CommandText = "alter table " & strTable & " add " & strField & " " & strFieldType & ""
            cmdTable.ExecuteNonQuery()
        End Try
    End Function
    'ra1.3.109

    'uj 1.3.109
    Public Sub ExportToExcel(ByVal dsExport As DataSet)
        Try
            Dim IEProcess As Process = New Process

            Dim hWndDesk
            Dim params
            Dim mFile As File
            Dim rs As Long
            params = vbNullString
            hWndDesk = GetDesktopWindow()

            Dim ts As StreamWriter
            Dim i, j As Integer
            Dim mExpText As String
            Dim mStr As String
            Dim mPos As Integer

            Dim ExlAppln As New Object
            Dim ExlBook As New Object

            ExlAppln = CreateObject("Excel.Application")

            ts = New StreamWriter(Application.StartupPath & "\ExportP.xls")

            mExpText = ""
            mStr = ""
            'ver 2.0.3-23

            For k As Integer = 0 To dsExport.Tables(0).Columns.Count - 1
                mStr = dsExport.Tables(0).Columns(k).ColumnName.ToString
                If mExpText = "" Then
                    mExpText = IIf(mStr = "", Space(1), mStr)
                    'mExpText = mExpText & Chr(9) & mStr
                Else
                    mExpText = mExpText & Chr(9) & IIf(mStr = "", Space(1), mStr)
                End If

                'memptext=
                'MsgBox(dsExport.Tables(0).Columns(k).ColumnName)
            Next
            ts.WriteLine(mExpText)
            'ver 2.0.2
            mExpText = ""
            mStr = ""
            'ver 2.0.3-23

            For i = 0 To dsExport.Tables(0).Rows.Count - 1
                mStr = ""
                For j = 0 To dsExport.Tables(0).Columns.Count - 1
                    If mExpText = "" Then
                        'Ver 2.3.6-E27 Start
                        'mExpText = IIf((dsExport.Tables(0).Rows(i).Item(j) = ""), Space(1), dsExport.Tables(0).Rows(i).Item(j))
                        If IsDBNull(dsExport.Tables(0).Rows(i).Item(j)) = False Then
                            'Ver 2.5.5-E243  Start
                            'mExpText = IIf((dsExport.Tables(0).Rows(i).Item(j) = ""), Space(1), dsExport.Tables(0).Rows(i).Item(j))
                            mExpText = IIf((Convert.ToString(dsExport.Tables(0).Rows(i).Item(j)) = ""), Space(1), Convert.ToString(dsExport.Tables(0).Rows(i).Item(j)))
                            'Ver 2.5.5-E243  End
                        End If
                        'Ver 2.3.6-E27 End
                        mExpText = IIf(Mid(mExpText, 1, 1) = "0", "'" & mExpText, mExpText)
                    Else
                        'Ver 2.3.6-E27 Start
                        'mPos = InStr(1, dsExport.Tables(0).Rows(i).Item(j), Chr(13))
                        'mStr = dsExport.Tables(0).Rows(i).Item(j)
                        If IsDBNull(dsExport.Tables(0).Rows(i).Item(j)) = False Then
                            mPos = InStr(1, dsExport.Tables(0).Rows(i).Item(j), Chr(13))
                            'Ver 2.5.5-E243 Start
                            ' mStr = dsExport.Tables(0).Rows(i).Item(j)
                            mStr = Convert.ToString(dsExport.Tables(0).Rows(i).Item(j))
                            'Ver 2.5.5-E243 End
                        Else
                            mStr = vbNullString
                        End If
                        'Ver 2.3.6-E27 End

                        Do While mPos > 0
                            mStr = Mid(mStr, 1, Val(mPos) - 1) & " " & Mid(mStr, Val(mPos) + 2)
                            mPos = InStr(1, mStr, Chr(13))
                        Loop

                        mExpText = mExpText & Chr(9) & mStr
                    End If
                Next
                ts.WriteLine(mExpText)
                mExpText = ""
            Next
            ts.Close()

            If mFile.Exists(Application.StartupPath & "\ExportP.xls") = True Then
                IEProcess.Start(Application.StartupPath & "\ExportP.xls")
            End If
        Catch ex As Exception
            MsgBox("The process cannot access the file """ & Application.StartupPath & "\ExportP.xls"" because it is being used by another process.")
            If ex.Message = "The process cannot access the file """ & Application.StartupPath & "\ExportP.xls"" because it is being used by another process." Then
                MsgBox("Close file 'ExportP.xls' and then try again!")
            Else
                Call ErrHandler(ex)
            End If

        End Try
    End Sub
    'ver 2.0.3-22
    Public Function RQ(ByVal StrVal As String) As String
        'Remove Quote
        RQ = Replace(StrVal, "'", "''")
        Return RQ
    End Function
    'ver 2.0.3-22

    Public Sub AlterImportSetupTable() 'ver 2.03 - For Pune Client
        '1. Objective - To Add Surcharge field into ImportSetup Table
        'first to find out whether this field is exising or not in the 
        'database
        Try
            'Ver 2.4.1-1629 Start
            Dim cmdNewFld As New SqlCommand
            'Ver 2.4.1-1629 End

            Dim adpNewField As New SqlDataAdapter("Select * from ImportSetup", conTdsPac)
            Dim ds As New DataSet
            adpNewField.Fill(ds)

            'Ver 2.3.8-1719 Start
            'If ds.Tables(0).Columns.Contains("Surcharge") = True Then
            '    adpNewField.Dispose()
            '    Exit Sub
            'End If
            If ds.Tables(0).Columns.Contains("Surcharge") = False Then
                'Ver 2.3.8-1719 End
                Dim conNewField As New SqlCommand("Alter table ImportSetup add Surcharge nVarChar(50)", conTdsPac)
                conNewField.ExecuteNonQuery()
                Dim conNewField1 As New SqlCommand("Alter table LogImportSetup add Surcharge nVarChar(50)", conTdsPac)
                conNewField1.ExecuteNonQuery()
            End If


            'Ver 2.4.1-1629 Start
            If ds.Tables(0).Columns.Contains("State") = False Then
                cmdNewFld.Connection = conTdsPac
                cmdNewFld.CommandText = "Alter table ImportSetup add State nvarchar(50)"
                cmdNewFld.ExecuteNonQuery()
                cmdNewFld.CommandText = "Alter table LogImportSetup add State nvarchar(50)"
                cmdNewFld.ExecuteNonQuery()
            End If
            'Ver 2.4.1-1629 End

            'Ver 2.3.8-1719 Start
            adpNewField.Dispose()
            ds.Dispose()

            Dim Da As New SqlDataAdapter
            Dim DsTdsSubCode As New DataSet
            Dim cmd As New SqlCommand

            Da = New SqlDataAdapter("Select * from ImportSetup", conTdsPac)
            Da.Fill(DsTdsSubCode)
            If DsTdsSubCode.Tables(0).Columns.Contains("TranTdsStatus") = False Then
                cmd = New SqlCommand("Alter table ImportSetup add TranTdsStatus nVarchar(50)", conTdsPac)
                cmd.ExecuteNonQuery()
                cmd = New SqlCommand("Alter table LogImportSetup add TranTdsStatus nVarchar(50)", conTdsPac)
                cmd.ExecuteNonQuery()
                cmd = New SqlCommand("Alter table ImportSetup add TranTdsSubcode nVarchar(50)", conTdsPac)
                cmd.ExecuteNonQuery()
                cmd = New SqlCommand("Alter table LogImportSetup add TranTdsSubcode nVarchar(50)", conTdsPac)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
            End If
            DsTdsSubCode.Dispose()
            Da.Dispose()
            'Ver 2.3.8-1719 End
            'Ver 2.5.3-E225 Start
            Dim DaImpSetUp As New SqlClient.SqlDataAdapter("select * from ImportSetup", conTdsPac)
            Dim DsImpSetUp As New DataSet
            DaImpSetUp.Fill(DsImpSetUp, "ImportSetup")
            If DsImpSetUp.Tables("ImportSetup").Columns.Contains("DateOfPayment") = False Then
                Dim cmdAlter As New SqlClient.SqlCommand
                cmdAlter.Connection = conTdsPac

                cmdAlter.CommandText = "ALTER table ImportSetup add DateOfPayment nvarchar(50), Threshold nvarchar(50)"
                cmdAlter.ExecuteNonQuery()

                cmdAlter.CommandText = "ALTER table LogImportSetup add DateOfPayment nvarchar(50), Threshold nvarchar(50)"
                cmdAlter.ExecuteNonQuery()
            End If
            DsImpSetUp.Dispose()
            DaImpSetUp.Dispose()
            'Ver 2.5.3-E225 End

        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Sub

    Public Sub CreateTabeTdsFile()
        'form16a-new format changes
        Try
            'ver2.0.4-PRN Change
            Dim cmd As New SqlClient.SqlCommand
            cmd.Connection = conTdsPac
            cmd.CommandText = "SELECT count(Name) FROM sysobjects Where Name= 'eTdsFile' AND xType= 'U'"
            If cmd.ExecuteScalar = 0 Then
                cmd.CommandText = "create table eTdsFile(FormNo varchar(10),Quarter varchar(10),RDate smalldatetime,PRNNo varchar(15),Cocd varchar(15))"
            Else
                Dim adpNewField As New SqlDataAdapter("Select * from eTdsFile", conTdsPac)
                Dim ds As New DataSet
                adpNewField.Fill(ds)
                If ds.Tables(0).Columns.Contains("Cocd") = True Then
                    adpNewField.Dispose()
                    Exit Sub
                End If
                adpNewField.Dispose()
                'cmd.CommandText = "alter table eTdsFile add Cocd varchar(15)"
            End If
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            If ex.Message = "There is already an object named 'eTdsFile' in the database." Then
                Exit Sub
            ElseIf ex.Message = "Column names in each table must be unique. Column name 'Cocd' in table 'eTdsFile' is specified more than once." Then
                Exit Sub
            Else
                Call ErrHandler(ex)
            End If
            'ver2.0.4-PRN Change
        End Try
    End Sub

    Sub CreateSunMaster()
        'Ver 3.1.0-REQ592 start
        '        On Error GoTo errhead
        '        Dim strSql As String
        '        Dim cmdSun As New SqlCommand
        '        cmdSun.Connection = conTdsPac
        '        'Checking the table
        '        Dim cmd As New SqlClient.SqlCommand

        '        cmd.Connection = conTdsPac

        '        cmd.CommandText = "SELECT count(Name) FROM sysobjects Where Name= 'SunMaster' AND xType= 'U'"
        '        If cmd.ExecuteScalar > 0 Then
        '            'ver 2.3.6-1479 start
        '            'Ver 2.3.9-E60 start
        '            'If StrFinYear = "2008-2009" Then
        '            'Ver 2.4.3-E93 start
        '            'If StrFinYear = "2008-2009" Or StrFinYear = "2009-2010" Then
        '            'Ver 2.5.0-E200 start
        '            'If StrFinYear = "2008-2009" Or StrFinYear = "2009-2010" Or StrFinYear = "2010-2011" Then
        '            'Ver 2.5.7-E322 Start
        '            'If StrFinYear = "2008-2009" Or StrFinYear = "2009-2010" Or StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Then
        '            'Ver 2.6.2-E462 start
        '            'If StrFinYear = "2008-2009" Or StrFinYear = "2009-2010" Or StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Or StrFinYear = "2012-2013" Then
        '            If StrFinYear = "2008-2009" Or StrFinYear = "2009-2010" Or StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Or StrFinYear >= "2012-2013" Then
        '                'Ver 2.6.2-E462 end
        '                'Ver 2.5.7-E322 End
        '                'Ver 2.5.0-E200 end
        '                'Ver 2.4.3-E93 end
        '                'Dim cmdTemp As New SqlCommand("SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='8'", conTdsPac)
        '                Dim cmdTemp As New SqlClient.SqlCommand
        '                cmdTemp.Connection = conTdsPac
        '                If StrFinYear = "2008-2009" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='8'"
        '                    'Ver 2.4.3-E93 start
        '                    'Else
        '                ElseIf StrFinYear = "2009-2010" Then
        '                    'Ver 2.4.3-E93 end
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='9'"
        '                    'Ver 2.4.3-E93 start
        '                ElseIf StrFinYear = "2010-2011" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='J'"
        '                    'Ver 2.4.3-E93 end
        '                    'Ver 2.5.0-E200 start
        '                ElseIf StrFinYear = "2011-2012" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='K'"
        '                    'Ver 2.5.0-E200 end
        '                    'Ver 2.5.7-E322 start
        '                ElseIf StrFinYear = "2012-2013" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='L'"
        '                    'Ver 2.5.7-E322 end
        '                    'Ver 2.6.2-E569 start
        '                ElseIf StrFinYear = "2013-2014" Then
        '                    'Ver 2.6.4-3783 start
        '                    'cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='3'"
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='M'"
        '                    'Ver 2.6.4-3783 end
        '                    'Ver 2.6.2-E569 end
        '                    'Ver3.0.0-Req300 Start
        '                ElseIf StrFinYear = "2014-2015" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='N'"
        '                    'Ver3.0.0-Req300 end 
        '                    'Ver 3.0.5-REQ515 start
        '                ElseIf StrFinYear = "2015-2016" Then
        '                    cmdTemp.CommandText = "SELECT count(*) from SunMaster where CodeField ='First' and CodeValue ='O'"
        '                    'Ver 3.0.5-REQ515 end
        '                End If
        '                'Ver 2.4.1-E79 start
        '                Call UpdateSun("If not exists (select * from SunMaster where CodeField ='First' and CodeValue ='B')     Insert into SunMaster values ('First','B','wef 1st April','')", cmdSun)
        '                Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='I') Insert into SunMaster values ('Third','I','Interest from Banking Copmpany','94A')", cmdSun)
        '                Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='K') Insert into SunMaster values ('Third','K','Interest other than from Banking Company','94A')", cmdSun)
        '                Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='H') Insert into SunMaster values ('Third','H','Winning from Horse Race','4BB')", cmdSun)
        '                Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='O') Insert into SunMaster values ('Third','O','Commision on sale of lottery tickets','94B')", cmdSun)
        '                'Ver 2.4.1-E79 end
        '                'Ver 2.7.4-REQ418 start
        '                If StrFinYear >= "2014-2015" Then
        '                    Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='D') Insert into SunMaster values ('Third','D','Payment under Life Insurance Policy','4DA')", cmdSun)
        '                End If
        '                'Ver 2.7.4-REQ418 end
        '                'Ver 2.3.9-E60 end
        '                If cmdTemp.ExecuteScalar = 0 Then
        '                'Ver 2.3.9-E60 start
        '                If StrFinYear = "2008-2009" Then
        '                    'Ver 2.3.9-E60 end
        '                    Call UpdateSun("Insert into SunMaster values ('First','8','wef 1st April','')", cmdSun)
        '                    'Ver 2.3.9-E60 start
        '                    'Ver 2.4.3-E93 start
        '                    'Else
        '                ElseIf StrFinYear = "2009-2010" Then
        '                    'Ver 2.4.3-E93 end
        '                    Call UpdateSun("Insert into SunMaster values ('First','9','wef 1st April','')", cmdSun)
        '                    'Ver 2.4.3-E93 start
        '                ElseIf StrFinYear = "2010-2011" Then
        '                    Call UpdateSun("Insert into SunMaster values ('First','J','wef 1st April','')", cmdSun)
        '                    'Ver 2.4.3-E93 end
        '                    'Ver 2.5.0-E200 start
        '                ElseIf StrFinYear = "2011-2012" Then
        '                    Call UpdateSun("Insert into SunMaster values ('First','K','wef 1st April','')", cmdSun)
        '                    'Ver 2.5.0-E200 end
        '                    'Ver 2.5.7-E322 Start
        '                ElseIf StrFinYear = "2012-2013" Then
        '                    Call UpdateSun("Insert into SunMaster values ('First','L','wef 1st April','')", cmdSun)
        '                    'Ver 2.5.7-E322 End
        '                    'Ver 2.6.2-E569 start
        '                ElseIf StrFinYear = "2013-2014" Then
        '                    'Ver 2.6.4-3783 start
        '                    'Call UpdateSun("Insert into SunMaster values ('First','3','wef 1st April','')", cmdSun)
        '                    Call UpdateSun("Insert into SunMaster values ('First','M','wef 1st April','')", cmdSun)
        '                    'Ver 2.6.4-3783 end
        '                    'Ver 2.6.2-E569 end

        '                    'Ver3.0.0-Req300 Start
        '                ElseIf StrFinYear = "2014-2015" Then
        '                    Call UpdateSun("Insert into SunMaster values ('First','N','wef 1st April','')", cmdSun)
        '                    'Ver3.0.0-Req300 end
        '                    'Ver 3.0.5-REQ515 start
        '                ElseIf StrFinYear = "2015-2016" Then
        '                    Call UpdateSun("Insert into SunMaster values ('First','O','wef 1st April','')", cmdSun)
        '                    'Ver 3.0.5-REQ515 start
        '                End If
        '                'Ver 2.3.9-E60 end
        '                cmdTemp.Dispose()
        '            End If
        '            'ver 2.3.6-1479 end
        '            Exit Sub
        '        End If
        '        cmd.Dispose()

        '        strSql = "Create Table SunMaster (CodeField nVarChar(10),CodeValue nVarChar(20),Description nVarChar(50),TdsCode nVarChar(10))"
        '        cmdSun.CommandText = strSql
        '        cmdSun.ExecuteNonQuery()
        '        'Ver 2.3.9-E60 start
        '        'Call UpdateSun("Insert into SunMaster values ('First','A','Till 31st May','')", cmdSun)
        '        'Call UpdateSun("Insert into SunMaster values ('First','7','wef 1st June','')", cmdSun)
        '        If StrFinYear = "2007-2008" Then
        '            Call UpdateSun("Insert into SunMaster values ('First','A','Till 31st May','')", cmdSun)
        '            Call UpdateSun("Insert into SunMaster values ('First','7','wef 1st June','')", cmdSun)
        '        End If
        '        'Ver 2.3.9-E60 end
        '        'Ver 2.4.3-E93 start
        '        'Ver 2.5.0-E200 start
        '        'If StrFinYear = "2010-2011" Then
        '        'Ver 2.5.7-E322 Start
        '        'If StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Then
        '        'Ver 2.6.2-E569 start
        '        'If StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Or StrFinYear = "2012-2013" Then
        '        If StrFinYear = "2010-2011" Or StrFinYear = "2011-2012" Or StrFinYear >= "2012-2013" Then
        '            'Ver 2.6.2-E569 end
        '            'Ver 2.5.7-E322 End
        '            'Ver 2.5.0-E200 end
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='F') Insert into SunMaster values ('Third','F','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='G') Insert into SunMaster values ('Third','G','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='J') Insert into SunMaster values ('Third','J','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='Q') Insert into SunMaster values ('Third','Q','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='U') Insert into SunMaster values ('Third','U','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='V') Insert into SunMaster values ('Third','V','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='W') Insert into SunMaster values ('Third','W','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='X') Insert into SunMaster values ('Third','X','For 195','195')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='Z') Insert into SunMaster values ('Third','Z','For 195','195')", cmdSun)
        '        End If
        '        'Ver 2.4.3-E93 end
        '        Call UpdateSun("Insert into SunMaster values ('Second','A','Companies','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','B','Individual','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','C','Companies','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','F','Firm','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','H','HUF','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','J','Artificial Juricidical Person','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','L','Local Authority','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','N','Individual','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','O','Co-Operative Society','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Second','P','AOP/BOI','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','A','Advertisement','94C')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','B','Brokerage / Commission','94H')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','C','Contractor','94C')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','I','Interest','94A')", cmdSun)
        '        'Ver 2.6.2-E569 start
        '        'Ver 3.0.0-REQ300 Start
        '        'If StrFinYear = "2013-2014" Then
        '        If StrFinYear >= "2013-2014" Then
        '            'Ver 3.0.0-REQ300 end 
        '            Call UpdateSun("Insert into SunMaster values ('Third','K','Interest other than from Banking Company','94A')", cmdSun)
        '            'Ver 3.1.0-?? start
        '            'Call UpdateSun("Insert into SunMaster values ('Third','H','Directors Fees','94J')", cmdSun)
        '            'Call UpdateSun("Insert into SunMaster values ('Third','V','Immovable Property (Note 4)','4IA')", cmdSun)
        '            'Ver 3.1.0-?? end
        '        End If
        '        'Ver 2.6.2-E569 end
        '        Call UpdateSun("Insert into SunMaster values ('Third','N','Insurance Commission','94D')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','P','Professional Fees','94J')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','R','Rent','94I')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','S','Sub Contractor','94C')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','L','Winning from Lotteries','94B')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','Y','Royalty','195')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Third','M','Plant & Machinery Rent','94I')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Fourth','0','Zero Rate - Self declration','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Fourth','1','Normal Rate','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Fourth','A - Z','Special Rate','')", cmdSun)
        '        Call UpdateSun("Insert into SunMaster values ('Fourth','2','Surcharge Rate 10','')", cmdSun)
        '        'Ver 2.4.1-E72 start
        '        Call UpdateSun("Insert into SunMaster values ('Fourth','3','No SC No Cess','')", cmdSun)
        '        'Ver 2.4.1-E72 end
        '        'ver 2.3.6-1479 start
        '        If StrFinYear = "2008-2009" Then Call UpdateSun("Insert into SunMaster values ('First','8','wef 1st April','')", cmdSun)
        '        'ver 2.3.6-1479 End
        '        'Ver 2.3.9-E60 start
        '        If StrFinYear = "2009-2010" Then
        '            Call UpdateSun("Insert into SunMaster values ('Third','D','Dividend','194')", cmdSun)
        '            Call UpdateSun("Insert into SunMaster values ('Third','E','Deposit of NSS','4EE')", cmdSun)
        '            Call UpdateSun("Insert into SunMaster values ('Third','T','Interest on Securities','193')", cmdSun)
        '            'Ver 2.4.1-E79 start
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='First' and CodeValue ='B')                             (Insert into SunMaster values ('First','B','wef 1st April','')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='First' and CodeValue ='9')                             (Insert into SunMaster values ('First','9','wef 1st April','')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='I') Insert into SunMaster values ('Third','I','Interest from Banking Copmpany','94A')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='K') Insert into SunMaster values ('Third','K','Interest other than from Banking Company','94A')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='H') Insert into SunMaster values ('Third','H','Winning from Horse Race','4BB')", cmdSun)
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Third' and CodeValue ='O') Insert into SunMaster values ('Third','O','Commision on sale of lottery tickets','94B')", cmdSun)
        '            'Ver 2.4.1-E79 end
        '        End If
        '        'Ver 2.4.6-2564 start
        '        If StrFinYear = "2010-2011" Then
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Fourth' and CodeValue ='4') Insert into SunMaster values ('Fourth','4','Penal Rate','')", cmdSun)
        '            'Ver 2.50-E200 start
        '        ElseIf StrFinYear = "2011-2012" Then
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Fourth' and CodeValue ='5') Insert into SunMaster values ('Fourth','5','Threshold Entries','')", cmdSun)
        '            'Ver 2.50-E200 end
        '            'Ver 2.5.7-E322 Start
        '            'Ver 2.6.2-E569 start
        '            'ElseIf StrFinYear >= "2012-2013" Then
        '        ElseIf StrFinYear >= "2012-2013" Then
        '            'Ver 2.6.2-E569 end
        '            Call UpdateSun("If not exists (select * from SunMaster where CodeField ='Fourth' and CodeValue ='5') Insert into SunMaster values ('Fourth','5','Threshold Entries','')", cmdSun)
        '            'Ver 2.5.7-E322 End
        '        End If
        '        'Ver 2.4.6-2564 start
        '        'Ver 2.3.9-E60 end
        'errhead:

        '        Exit Sub

        Dim _textStreamReader As StreamReader
        Dim _Assembly As Reflection.[Assembly]
        Dim strSql As String
        Dim cmdSun As New SqlCommand
        Dim cmd As New SqlClient.SqlCommand

        Dim strStreamVal As String
        Dim arryVal() As String

        cmdSun.Connection = conTdsPac
        strStreamVal = ""

        Try

            cmd.Connection = conTdsPac
            'Ver 3.1.4-QC1047 start
            cmd.CommandText = "SELECT count(Name) FROM sysobjects Where Name= 'SunMaster' AND xType= 'U'"
            If cmd.ExecuteScalar = 0 Then
                strSql = "Create Table SunMaster (CodeField nVarChar(10),CodeValue nVarChar(20),Description nVarChar(50),TdsCode nVarChar(10))"
                cmdSun.CommandText = strSql
                cmdSun.ExecuteNonQuery()
            End If
            'Ver 3.1.4-QC1047 end 
            cmd.CommandText = "Delete from SunMaster"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "SELECT count(Name) FROM sysobjects Where Name= 'SunMaster' AND xType= 'U'"
            If cmd.ExecuteScalar > 0 Then
                _Assembly = Reflection.[Assembly].GetExecutingAssembly()

                If File.Exists(Application.StartupPath + "\Aastral.bmp") = False And File.Exists(Application.StartupPath + "\Sun.bmp") = False Then
                    Exit Sub
                End If

                If File.Exists(Application.StartupPath + "\Aastral.bmp") = True Then
                    _textStreamReader = New StreamReader(_Assembly.GetManifestResourceStream("Tdspac.AastralCode.txt"))
                Else
                    _textStreamReader = New StreamReader(_Assembly.GetManifestResourceStream("Tdspac.SunCode.txt"))
                End If

                'StrFinYear = "2021-2122"

                Do While _textStreamReader.Peek() <> -1
                    strStreamVal = _textStreamReader.ReadLine()
                    arryVal = strStreamVal.Split(",")
                    If arryVal(0) = "FinYear" And arryVal(1) = StrFinYear Then
                        Do While _textStreamReader.Peek() <> -1
                            strStreamVal = _textStreamReader.ReadLine()
                            arryVal = strStreamVal.Split(",")
                            If arryVal(0) <> "FinYear" Then
                                'Call UpdateSun("If not exists (select * from SunMaster where CodeField =" & arryVal(0) & " and CodeValue =" & arryVal(1) & ") Insert into SunMaster values (" & strStreamVal & ")", cmdSun)
                                Call UpdateSun("If not exists (select * from SunMaster where CodeField =" & arryVal(0) & " and CodeValue =" & arryVal(1) & " and TdsCode=" & arryVal(3) & ") Insert into SunMaster values (" & strStreamVal & ")", cmdSun)
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                Loop

            End If
            cmd.Dispose()
            cmdSun.Dispose()

        Catch ex As Exception
            Exit Sub
        End Try
        'Ver 3.1.0-REQ592 end
    End Sub

    Private Sub UpdateSun(ByVal strSql As String, ByVal cmdSun As SqlCommand)
        Try
            cmdSun.CommandText = strSql
            cmdSun.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function ValidPanPattern(ByVal mPan As String) As Boolean
        Dim j As Integer
        Dim strValidChr As String

        strValidChr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        ValidPanPattern = False

        'Ver 2.3.8-II Priority 1 start
        If Trim(mPan) = "" Or Len(Trim(mPan)) <> 10 Then
            ValidPanPattern = False
            Exit Function
        End If
        'Ver 2.3.8-II Priority 1 end

        For j = 1 To 5
            If j = 4 Then
                If InStr(1, "ABCFGHJLPT", UCase(Mid(mPan, 4, 1))) = 0 Then
                    ValidPanPattern = False
                    Exit Function
                End If
            End If

            If InStr(1, strValidChr, Mid(mPan, j, 1)) = 0 Then
                ValidPanPattern = False
                Exit Function
            End If
        Next
        For j = 6 To 9
            If InStr(1, "1234567890", Mid(mPan, j, 1)) = 0 Then
                ValidPanPattern = False
                Exit Function
            End If
        Next
        If InStr(1, strValidChr, Mid(mPan, 10, 1)) = 0 Then
            ValidPanPattern = False
            Exit Function
        End If
        ValidPanPattern = True

    End Function
    Function CreatePanFile(ByVal mFolder As String) As Boolean

        On Error GoTo errhead
        strPanFile = mFolder & "\InvalidPan.txt"
        tsPanFile = New System.IO.StreamWriter(strPanFile)
        tsPanFile.WriteLine("Sr.No " & "Party Name" & Space(65) & "Wrong Pan")
        CreatePanFile = True
        Exit Function

errhead:
        If Err.Number = 70 Then
            MsgBox("Please close file : " & strPanFile & "  before creating eTDS File", vbInformation, "eTDSWiz")
            CreatePanFile = False
            Exit Function
        Else

        End If
    End Function
    Function CalculatePercentage(ByVal blnSalary As Boolean) As Boolean
        Dim curPercentage As Decimal
        Dim curValidPercet As Decimal
        Dim strCorrect As Integer
        On Error Resume Next
        tsPanFile.Close()
        If blnSalary = True Then
            'ver 2.3.2-FVU2116 start
            'curValidPercet = 10
            'strCorrect = 90
            curValidPercet = 5
            strCorrect = 95
            'ver 2.3.2-FVU2116 end
        Else
            'ver 2.3.2-FVU2116 start
            'curValidPercet = 30
            'strCorrect = 70
            curValidPercet = 15
            strCorrect = 85
            'ver 2.3.2-FVU2116 end
        End If

        Dim r
        curPercentage = (intTotalWrongPan / intTotalRecordForPan) * 100
        If curPercentage > curValidPercet Then
            MsgBox("Total Deductee Records= " & intTotalRecordForPan & vbCrLf & "Recods with structurally invalid PANs = " & intTotalWrongPan & vbCrLf & "Percentage of wrong PAN = " & curPercentage & vbCrLf & "eTDS file cannot be created. Minimum " & strCorrect & " % PANs must be structurally valid", vbCritical)

            If MsgBox("Do you still want to Create eTds File?", MsgBoxStyle.YesNo) = vbYes Then
                CalculatePercentage = True
            Else
                Dim IEProcess As Process = New Process
                IEProcess.Start(strPanFile)
                'r = ShellExecute(hWndDesk, "Open", strPanFile, params, 0&, 1)
                CalculatePercentage = False
            End If           'End

            Exit Function
        End If
        CalculatePercentage = True
    End Function
    Sub UpdatePanFile(ByVal StrParty As String, ByVal strPAN As String)
        On Error Resume Next
        intTotalRecordForPan = intTotalRecordForPan + 1
        If ValidPanPattern(strPAN) = False Then
            intTotalWrongPan = intTotalWrongPan + 1
            tsPanFile.WriteLine(intTotalWrongPan & Space(7 - intTotalWrongPan.ToString.Length) & StrParty & Space(75 - StrParty.Length) & strPAN)
        End If

    End Sub

    'Ver 2.4.0-AXIS-BANK Start
    Function PrnUpdated(ByVal strTdsCode As String, ByVal strChlnNo As String) As Integer
        Dim strQuery As String
        Dim cmd As New SqlCommand

        If blnPrnRestrict Then
            If strTdsCode <> "195" Then
                strQuery = "select count(c.[Quarter]) from eTdsFile e, Challan C where e.cocd=c.cocd and e.[quarter]=c.[quarter] and e.FormNo='26Q' and c.challanno='" & strChlnNo & "' and e.Cocd='" & StrCocd & "'"
            Else
                strQuery = "select count(c.[Quarter]) from eTdsFile e, Challan c where e.cocd=c.cocd and e.[quarter]=c.[quarter] and e.FormNo='27Q' and c.challanno='" & strChlnNo & "' and e.Cocd='" & StrCocd & "' "
            End If
            cmd.Connection = conTdsPac
            cmd.CommandText = strQuery
            Return cmd.ExecuteScalar
        End If
    End Function
    'Ver 2.4.0-AXIS-BANK End

    'Ver 2.4.0-E62 Start
    Public Function DecryptFile(ByVal strFileName As String) As String
        Try
            Dim bytKey As Byte()
            Dim bytIV As Byte()
            Dim strOutput As String
            Dim sreader As StreamReader

            '--Send the password to the CreateKey function.
            bytKey = CreateKey(PagePass)
            '--Send the password to the CreateIV function.
            bytIV = CreateIV(PagePass)
            '--Start the decryption.
            strOutput = Application.StartupPath & "\Comp.txt"

            fsInput = New System.IO.FileStream(strFileName, FileMode.Open, FileAccess.Read)

            Dim csCryptoStream As CryptoStream
            '--Declare your CryptoServiceProvider.
            Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged
            csCryptoStream = New CryptoStream(fsInput, _
            cspRijndael.CreateDecryptor(bytKey, bytIV), _
            CryptoStreamMode.Read)

            sreader = New StreamReader(csCryptoStream)

            'If sreader = False Then
            Try
                Return (sreader.ReadLine)
            Catch ex As Exception
                Call ErrHandler(ex)
            End Try

            'End If

            If Not (csCryptoStream Is Nothing) Then
                csCryptoStream.Close()
            End If

        Catch ex As Exception
        Finally
            '--Close FileStreams and CryptoStream.
            If Not (fsInput Is Nothing) Then
                fsInput.Close()
            End If
        End Try
        DecryptFile = ""
    End Function
    'Ver 2.4.0-E62 End
    'Ver 2.4.1-E71 Start
    Public Sub AddToTdsRate(ByVal mFileName As String, ByVal mTableName As String)
        Dim objReader As StreamReader
        Dim strFilePath As String
        Dim strLine As String
        Dim cmdCommand As New SqlClient.SqlCommand
        Try

            strFilePath = Application.StartupPath & "\" & mFileName

            cmdCommand.Connection = conTdsPac

            objReader = New StreamReader(strFilePath)
            strLine = Replace(Replace(Replace(objReader.ReadLine, """", "'"), "True", "1"), "False", "0")
            Do Until strLine = ""
                cmdCommand.CommandText = "Insert Into " & mTableName & " values (" & strLine & ")"
                cmdCommand.ExecuteNonQuery()
                strLine = Replace(Replace(Replace(objReader.ReadLine, """", "'"), "True", "1"), "False", "0")
            Loop
        Catch ex As Exception
            Call ErrHandler(ex)
        End Try
    End Sub
    'Ver 2.4.1-E71 End

    'Ver 2.4.6-E104 Start
    Public Function IsStoredProExists() As Boolean
        Dim dr As SqlDataReader
        Dim cmd As New SqlCommand
        Try
            cmd.Connection = conTdsPac
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "select * from sysobjects where name='GenCertificate2010' and Type='P'"
            dr = cmd.ExecuteReader()
            Return dr.HasRows
        Catch ee As Exception
            Return False
        Finally
            dr.Close()
            cmd.Dispose()
        End Try
    End Function
    'Ver 2.4.6-E104 End

#Region "1. Global Variables "

    '*************************
    '** Global Variables
    '*************************

    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream

#End Region
#Region "2. Create A Key "

    '*************************
    '** Create A Key
    '*************************

    Private Function CreateKey(ByVal strPassword As String) As Byte()
        'Convert strPassword to an array and store in chrData.
        Dim chrData() As Char = strPassword.ToCharArray
        'Use intLength to get strPassword size.
        Dim intLength As Integer = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        'Declare bytKey(31).  It will hold 256 bits.
        Dim bytKey(31) As Byte

        'Use For Next to put a specific size (256 bits) of 
        'bytResult into bytKey. The 0 To 31 will put the first 256 bits
        'of 512 bits into bytKey.
        For i As Integer = 0 To 31
            bytKey(i) = bytResult(i)
        Next

        Return bytKey 'Return the key.
    End Function

#End Region
#Region "3. Create An IV "

    '*************************
    '** Create An IV
    '*************************

    Private Function CreateIV(ByVal strPassword As String) As Byte()
        'Convert strPassword to an array and store in chrData.
        Dim chrData() As Char = strPassword.ToCharArray
        'Use intLength to get strPassword size.
        Dim intLength As Integer = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        'Declare bytIV(15).  It will hold 128 bits.
        Dim bytIV(15) As Byte

        'Use For Next to put a specific size (128 bits) of 
        'bytResult into bytIV. The 0 To 30 for bytKey used the first 256 bits.
        'of the hashed password. The 32 To 47 will put the next 128 bits into bytIV.
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next

        Return bytIV 'return the IV
    End Function

#End Region
#Region "4.Decrypt File "

    '****************************
    '** Decrypt File
    '****************************

    Private Enum CryptoAction
        'Define the enumeration for CryptoAction.
        ActionEncrypt = 1
        ActionDecrypt = 2
    End Enum
#End Region
    'Ver 2.4.0-E62 End 

    'Ver 2.3.8-2678 Start    
    Public Function ValidateValueAsPerRegExp(ByVal pValue As String, ByVal pRegExp As String) As Boolean
        Dim obj As Match = System.Text.RegularExpressions.Regex.Match(pValue, pRegExp)
        If Not obj.Success Then
            Return False
        Else
            Return True
        End If
    End Function
    'Ver 2.3.8-2678 End




    'Ver 2.5.5-E92 Start Start
    Public Function VerifyChallan(ByVal CsiFilePath As String, ByVal BankBranchCode As String, ByVal DateOfChallanNo As String, ByVal BankChallanNo As String, ByVal DepositAmountAsPerChallan As Double) As Boolean
        VerifyChallan = False
        ' Dim MD5 As New MD5

        Dim header As String
        Dim footer As String
        Dim InputString As String
        Dim ChlnAmount As String

        header = "1qi5b63p"
        footer = "9rtio7lb"

        'TanOfDeductor
        DateOfChallanNo = Replace(DateOfChallanNo, "/", "")
        ChlnAmount = Format(DepositAmountAsPerChallan, ".00")
        InputString = header & BankBranchCode & DateOfChallanNo & BankChallanNo & TANDeductor & ChlnAmount & footer


        Dim tsFile As String
        Dim objReader As New System.IO.StreamReader(CsiFilePath)

        tsFile = objReader.ReadLine

        Dim strValue As String

        'Ver 3.0.0 REQ-258 start   ' changed the previous code as it was not working 28/03/2014 -Jitendra
        'Do While objReader.Peek <> -1
        '    strValue = objReader.ReadLine
        '    If (strValue = MD5.GetHashCode1(InputString)) Then
        '        VerifyChallan = True
        '        Exit Do
        '    End If
        'Loop

        Dim md5hash As MD5 = MD5.Create()

        Dim hash As String = GetMd5Hash(md5hash, InputString)
        Do While objReader.Peek <> -1
            strValue = objReader.ReadLine
            If (strValue = hash) Then
                Return True
                'Exit Do
            End If
        Loop
        'Ver 3.0.0 REQ-258 end 
        objReader.Close()

    End Function


    'Ver 2.5.5-E92 End

    'Ver 2.5.6-2770 start
    Public Function Qt(ByRef s As String) As String
        Qt = Chr(34) & s & Chr(34) & " "
    End Function
    'Ver 2.5.6-2770 end
    'Ver 2.5.9-E371 start
    Public Function UpdateSection()
        On Error GoTo errhead

        'Vishal done for execute salTdsQuerty for adding LADAKH start Ver 3.2.7.1-FastFacts-787 start-0164939 start
        Dim cmdSALTDS As New SqlCommand
        If IsACFFound = True Then
            cmdSALTDS.Connection = conSALTDS
        End If
        'Vishal done for execute salTdsQuerty for adding LADAKH end Ver 3.2.7.1-FastFacts-787 start-0164939 start

        Dim cmdTds As New SqlCommand
        cmdTds.Connection = conTdsPac
        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = conTdsPac

        If Val(Mid(StrFinYear, 1, 4)) >= "2011" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4LB') Insert into Tdscode values ('4LB','194LB','Interest on infrastructure debt bond','',0,0,'1','27Q','16A','03/01/2005 0:00','Admin','')", cmdTds)
        End If
        If Val(Mid(StrFinYear, 1, 4)) >= "2012" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4LC') Insert into Tdscode values ('4LC','194LC','Interest on Specified Company','',0,0,'1','27Q','16A','03/01/2005 0:00','Admin','')", cmdTds)
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='6CJ') Insert into Tdscode values ('6CJ','206CJ','Collection at source on sale of mineral','',0,0,'1','27EQ','27D','03/01/2005 0:00','Admin','')", cmdTds)
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='6CK') Insert into Tdscode values ('6CK','206CK','Collection at source on sale of Jewellary bullion','',0,0,'1','27EQ','27D','03/01/2005 0:00','Admin','')", cmdTds)
            'Ver 2.6.2-E462 start
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4IA') Insert into Tdscode values ('4IA','194IA','Tds On immovable property','',0,0,'1','26Q','16A','06/01/2013 0:00','Admin','')", cmdTds)
            'Ver 2.6.2-E462 end
        End If
        'Ver 2.6.6-E677 start
        If Val(Mid(StrFinYear, 1, 4)) >= "2013" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4LD') Insert into Tdscode values ('4LD','194LD','Income by way of interest on certain bonds and Govt securities.','',0,0,'1','27Q','16A','04/01/2013 0:00','Admin','')", cmdTds)
        End If
        'Ver 2.6.6-E677  end
        'Ver 2.7.4-REQ418 start
        If Val(Mid(StrFinYear, 1, 4)) >= "2014" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4DA') Insert into Tdscode values ('4DA','194DA','Life Insurance Policy maturity amount.','',0,0,'1','26Q','16A','10/01/2014 0:00','Admin','')", cmdTds)
            'Ver 3.0.4-REQ465 start
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4BA') Insert into Tdscode values ('4BA','194LBA','Income from units of a business trust','',0,0,'1','26Q','16A','10/01/2014 0:00','Admin','')", cmdTds)
            'Ver 3.0.4-REQ465 end
            'Ver 3.1.1-REQ608 start
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='9IA') Insert into Tdscode values ('9IA','194IA','TDS on sale of property','',0,0,'1','26Q','16A','2014-04-01 0:00','Admin','')", cmdTds)
            'Ver 3.1.1-REQ608 end
        End If
        'Ver 2.7.4-REQ418 end
        'Ver 3.19-REQ743 start
        If Val(Mid(StrFinYear, 1, 4)) >= "2017" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4IC') Insert into Tdscode values ('4IC','194IC','Payment under specified agreement.','',0,0,'1','26Q','16A','01/04/2017 0:00','Admin','')", cmdTds)
        End If
        'Ver 3.19-REQ743 end

        'Ver 3.2.7.1-FastFacts-787 start
        If Val(Mid(StrFinYear, 1, 4)) >= "2019" Then
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='94A' and ThresholdAmt=40000) Update Tdscode set ThresholdAmt=40000 where TdsCode ='94A'", cmdTds)
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='94I' and ThresholdAmt=240000) Update Tdscode set ThresholdAmt=240000 where TdsCode ='94I'", cmdTds)
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4IA' and ThresholdAmt=240000) Update Tdscode set ThresholdAmt=240000 where TdsCode ='4IA'", cmdTds)
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='4IB' and ThresholdAmt=240000) Update Tdscode set ThresholdAmt=240000 where TdsCode ='4IB'", cmdTds)

            'Ver 3.2.7.1-FastFacts-787 start-0164939 start
            Call UpdateSection("If not exists (select * from Tdscode where TdsCode ='94N') Insert into Tdscode values ('94N','194N','Cash payments in excess of INR 1 Crores','',0,0,'1','26Q','16A','09/01/2019 0:00','Admin','')", cmdTds)
            If IsACFFound = True Then
                Call UpdateSection("If not exists (select * from mdm.state where statename ='LADAKH') Insert into mdm.state (statename,StateCode,CountryID,createDby,CreatedDate) values('LADAKH',37,1,1,'2019-09-01')", cmdSALTDS)
                Call UpdateSection("If not exists (select * from mdm.city where cityname ='LADAKH') Insert into mdm.city (cityname,citycode,StateID,ismetro,CreatedBy,createddate) values ('LADAKH',(select max(citycode) from mdm.city)+1,(select mdm.state.stateid  from mdm.state where mdm.state.statecode='37'),0,1,'2019-09-01')", cmdSALTDS)
            End If
            'Ver 3.2.7.1-FastFacts-787 start-0164939 end
        End If


        'Ver 3.2.7.1-FastFacts-787 end

errhead:

        Exit Function
    End Function
    'Ver 2.5.9-E371 end
    'Ver 2.5.9-E371 start
    Private Sub UpdateSection(ByVal strSql As String, ByVal cmdTds As SqlCommand)
        Try
            cmdTds.CommandText = strSql
            cmdTds.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'Ver 2.5.9-E371 end 
    'Ver 2.6.0-E384 start
    Public Function CheckDuplicate(ByVal strField As String, ByVal strVal As String, ByVal strEmpCd As String) As Boolean
        Try
            Dim blnDuplicate As Boolean
            If strVal = "" Then Exit Function
            Dim cmd As New SqlCommand("Select " & strField & " from DeducteeMaster where deducteecode <> '" & strEmpCd & "' and " & strField & " ='" & strVal & "'", conTdsPac)
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            If dr.HasRows Then
                blnDuplicate = True
            End If
            dr.Close()
            cmd.Clone()
            Return blnDuplicate
            Exit Function
        Catch ex As Exception
            Return False
        End Try

    End Function
    'Ver 2.6.0-E384 end
    'ver 2.6.0-E422 Start
    Function EncryptVBOLD(ByVal InputTxt As String)
        Dim ctr As Double
        For ctr = Len(InputTxt) To 1 Step -1
            EncryptVBOLD = EncryptVBOLD & Chr(Asc(Mid(InputTxt, ctr, 1)) + (30))
        Next ctr
    End Function
    Public Sub PDFReportExportToExcel(ByVal dsExport As DataTable, ByVal fPath As String)
        Try
            Dim IEProcess As Process = New Process

            Dim hWndDesk
            Dim params
            Dim mFile As File
            Dim rs As Long
            params = vbNullString
            hWndDesk = GetDesktopWindow()

            Dim ts As StreamWriter
            Dim i, j As Integer
            Dim mExpText As String
            Dim mStr As String
            Dim mPos As Integer

            Dim ExlAppln As New Object
            Dim ExlBook As New Object

            ExlAppln = CreateObject("Excel.Application")

            ts = New StreamWriter(fPath)

            mExpText = ""
            mStr = ""
            'ver 2.0.3-23

            For k As Integer = 0 To dsExport.Columns.Count - 1
                mStr = dsExport.Columns(k).ColumnName.ToString
                If mExpText = "" Then
                    mExpText = IIf(mStr = "", Space(1), mStr)
                    'mExpText = mExpText & Chr(9) & mStr
                Else
                    mExpText = mExpText & Chr(9) & IIf(mStr = "", Space(1), mStr)
                End If

                'memptext=
                'MsgBox(dsExport.Tables(0).Columns(k).ColumnName)
            Next
            ts.WriteLine(mExpText)
            'ver 2.0.2
            mExpText = ""
            mStr = ""
            'ver 2.0.3-23

            For i = 0 To dsExport.Rows.Count - 1
                mStr = ""
                For j = 0 To dsExport.Columns.Count - 1
                    If mExpText = "" Then
                        'Ver 2.3.6-E27 Start
                        'mExpText = IIf((dsExport.Tables(0).Rows(i).Item(j) = ""), Space(1), dsExport.Tables(0).Rows(i).Item(j))
                        If IsDBNull(dsExport.Rows(i).Item(j)) = False Then
                            'Ver 2.5.5-E243  Start
                            'mExpText = IIf((dsExport.Tables(0).Rows(i).Item(j) = ""), Space(1), dsExport.Tables(0).Rows(i).Item(j))
                            mExpText = IIf((Convert.ToString(dsExport.Rows(i).Item(j)) = ""), Space(1), Convert.ToString(dsExport.Rows(i).Item(j)))
                            'Ver 2.5.5-E243  End
                        End If
                        'Ver 2.3.6-E27 End
                        mExpText = IIf(Mid(mExpText, 1, 1) = "0", "'" & mExpText, mExpText)
                    Else
                        'Ver 2.3.6-E27 Start
                        'mPos = InStr(1, dsExport.Tables(0).Rows(i).Item(j), Chr(13))
                        'mStr = dsExport.Tables(0).Rows(i).Item(j)
                        If IsDBNull(dsExport.Rows(i).Item(j)) = False Then
                            mPos = InStr(1, dsExport.Rows(i).Item(j), Chr(13))
                            'Ver 2.5.5-E243 Start
                            ' mStr = dsExport.Tables(0).Rows(i).Item(j)
                            mStr = Convert.ToString(dsExport.Rows(i).Item(j))
                            'Ver 2.5.5-E243 End
                        Else
                            mStr = vbNullString
                        End If
                        'Ver 2.3.6-E27 End

                        Do While mPos > 0
                            mStr = Mid(mStr, 1, Val(mPos) - 1) & " " & Mid(mStr, Val(mPos) + 2)
                            mPos = InStr(1, mStr, Chr(13))
                        Loop

                        mExpText = mExpText & Chr(9) & mStr
                    End If
                Next
                ts.WriteLine(mExpText)
                mExpText = ""
            Next
            ts.Close()

            If mFile.Exists(fPath) = True Then
                '    IEProcess.Start(fPath)
            End If
        Catch ex As Exception
            MsgBox("The process cannot access the file """ & fPath & """ because it is being used by another process.")
            If ex.Message = "The process cannot access the file """ & fPath & """ because it is being used by another process." Then
                MsgBox("Close file 'ExportPDF.xls' and then try again!")
            Else
                Call ErrHandler(ex)
            End If

        End Try
    End Sub
    'ver 2.6.0-E422 End
    'Ver 2.6.8-REQ232 start
    Public Sub ExportToExcelNew(ByVal dsExport As DataView)
        Try
            Dim IEProcess As Process = New Process

            Dim hWndDesk
            Dim params
            Dim mFile As File
            Dim rs As Long
            params = vbNullString
            hWndDesk = GetDesktopWindow()

            Dim ts As StreamWriter
            Dim i, j As Integer
            Dim mExpText As String
            Dim mStr As String
            Dim mPos As Integer

            Dim ExlAppln As New Object
            Dim ExlBook As New Object

            ExlAppln = CreateObject("Excel.Application")

            ts = New StreamWriter(Application.StartupPath & "\ExportP.xls")

            mExpText = ""
            mStr = ""

            For k As Integer = 0 To dsExport.Table.Columns.Count - 1
                mStr = dsExport.Table.Columns(k).ColumnName.ToString
                If mExpText = "" Then
                    mExpText = IIf(mStr = "", Space(1), mStr)

                Else
                    mExpText = mExpText & Chr(9) & IIf(mStr = "", Space(1), mStr)
                End If

            Next
            ts.WriteLine(mExpText)

            mExpText = ""
            mStr = ""


            For i = 0 To dsExport.Count - 1
                mStr = ""
                For j = 0 To dsExport.Table.Columns.Count - 1
                    If mExpText = "" Then

                        If IsDBNull(dsExport.Item(i)(j)) = False Then

                            mExpText = IIf((Convert.ToString(dsExport.Item(i)(j)) = ""), Space(1), Convert.ToString(dsExport.Item(i)(j)))

                        End If

                        mExpText = IIf(Mid(mExpText, 1, 1) = "0", "'" & mExpText, mExpText)
                    Else

                        If IsDBNull(dsExport.Item(i)(j)) = False Then
                            mPos = InStr(1, dsExport.Item(i)(j), Chr(13))

                            mStr = Convert.ToString(dsExport.Item(i)(j))

                        Else
                            mStr = vbNullString
                        End If


                        Do While mPos > 0
                            mStr = Mid(mStr, 1, Val(mPos) - 1) & " " & Mid(mStr, Val(mPos) + 2)
                            mPos = InStr(1, mStr, Chr(13))
                        Loop

                        mExpText = mExpText & Chr(9) & mStr
                    End If
                Next
                ts.WriteLine(mExpText)
                mExpText = ""
            Next
            ts.Close()

            If mFile.Exists(Application.StartupPath & "\ExportP.xls") = True Then
                IEProcess.Start(Application.StartupPath & "\ExportP.xls")
            End If
        Catch ex As Exception
            MsgBox("The process cannot access the file """ & Application.StartupPath & "\ExportP.xls"" because it is being used by another process.")
            If ex.Message = "The process cannot access the file """ & Application.StartupPath & "\ExportP.xls"" because it is being used by another process." Then
                MsgBox("Close file 'ExportP.xls' and then try again!")
            Else
                Call ErrHandler(ex)
            End If

        End Try
    End Sub
    Public Sub ExportToExcelNew(ByVal dsExport As DataView, StrExcel As String)
        Try
            Dim IEProcess As Process = New Process

            Dim hWndDesk
            Dim params
            Dim mFile As File
            Dim rs As Long
            params = vbNullString
            hWndDesk = GetDesktopWindow()

            Dim ts As StreamWriter
            Dim i, j As Integer
            Dim mExpText As String
            Dim mStr As String
            Dim mPos As Integer

            Dim ExlAppln As New Object
            Dim ExlBook As New Object

            ExlAppln = CreateObject("Excel.Application")

            ts = New StreamWriter(StrExcel)

            mExpText = ""
            mStr = ""

            For k As Integer = 0 To dsExport.Table.Columns.Count - 1
                mStr = dsExport.Table.Columns(k).ColumnName.ToString
                If mExpText = "" Then
                    mExpText = IIf(mStr = "", Space(1), mStr)

                Else
                    mExpText = mExpText & Chr(9) & IIf(mStr = "", Space(1), mStr)
                End If

            Next
            ts.WriteLine(mExpText)

            mExpText = ""
            mStr = ""


            For i = 0 To dsExport.Count - 1
                mStr = ""
                For j = 0 To dsExport.Table.Columns.Count - 1
                    If dsExport.Table.Columns(j).DataType.Name = "DataTable" Then Continue For
                    If mExpText = "" Then

                        If IsDBNull(dsExport.Item(i)(j)) = False Then

                            mExpText = IIf((Convert.ToString(dsExport.Item(i)(j)) = ""), Space(1), Convert.ToString(dsExport.Item(i)(j)))

                        End If

                        mExpText = IIf(Mid(mExpText, 1, 1) = "0", "'" & mExpText, mExpText)
                    Else

                        If IsDBNull(dsExport.Item(i)(j)) = False Then
                            mPos = InStr(1, dsExport.Item(i)(j), Chr(13))

                            mStr = Convert.ToString(dsExport.Item(i)(j))

                        Else
                            mStr = vbNullString
                        End If


                        Do While mPos > 0
                            mStr = Mid(mStr, 1, Val(mPos) - 1) & " " & Mid(mStr, Val(mPos) + 2)
                            mPos = InStr(1, mStr, Chr(13))
                        Loop

                        mExpText = mExpText & Chr(9) & mStr
                    End If
                Next
                ts.WriteLine(mExpText)
                mExpText = ""
            Next
            ts.Close()

            If mFile.Exists(StrExcel) = True Then
                IEProcess.Start(StrExcel)
            End If
        Catch ex As Exception
            MsgBox("The process cannot access the file " & StrExcel & " because it is being used by another process.")
            If ex.Message = "The process cannot access the file " & StrExcel & " because it is being used by another process." Then
                MsgBox("Close file 'ExportP.xls' and then try again!")
            Else
                Call ErrHandler(ex)
            End If

        End Try
    End Sub

    'Ver 2.6.8-REQ232 End
    'Ver 3.0.0 REQ-258 start '==>Jitendra
    Public Function GetMd5Hash(ByVal md5Hash As MD5, ByVal input As String) As String

        ' Convert the input string to a byte array and compute the hash.
        Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input))

        ' Create a new Stringbuilder to collect the bytes
        ' and create a string.
        Dim sBuilder As New StringBuilder()

        ' Loop through each byte of the hashed data 
        ' and format each one as a hexadecimal string.
        Dim i As Integer
        For i = 0 To data.Length - 1
            sBuilder.Append(data(i).ToString("x2"))
        Next i

        ' Return the hexadecimal string.
        Return sBuilder.ToString()

    End Function
    'Ver 3.0.0 REQ-258 end 

    'Ver 3.05-Softlicence start
    Public Sub ReadCustomerCode()
        Try

            If (File.Exists(strLICXMLFile)) Then

                Using dsLicense = New DataSet()
                    dsLicense.ReadXml(strLICXMLFile)
                    strCustomerCode = dsLicense.Tables(0).Rows(0)("CustomerCode").ToString()
                End Using
            Else
                strCustomerCode = ""
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "TDSPAC", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub OpenOutlook()
        Try

            'Process.Start(String.Format("mailto:{0}?subject={1}&cc={2}&bcc={3}&body={4}", strTReTdsEmailId, "", "", "", ""));
            Process.Start(String.Format("mailto:{0}", strTReTdsEmailId))

        Catch ex As Exception

            MessageBox.Show(ex.Message, "TDSPAC", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Ver 3.05-Softlicence end
    '30/06/2015 for JAVA utility Jitendra

    'Ver 3.0.7 REQ530 end  
    Public Function old_GetJavaPath() As String
        Try


            Dim strJavaPath As String = ""
            'If strJavaPath = "" Then
            '    If Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)) = True Then
            '        If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre6\bin\javaw.exe") = True Then
            '            strJavaPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre6\bin\javaw "
            '        Else
            '            If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre7\bin\javaw.exe") = True Then
            '                strJavaPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre7\bin\javaw "
            '            End If
            '        End If
            '    End If
            'End If

            If strJavaPath = "" Then
                If Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)) = True Then
                    If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre1.8.0_65\bin\javaw.exe") = True Then
                        strJavaPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\Java\jre1.8.0_65\bin\javaw "
                    End If
                End If
            End If

            Return strJavaPath
        Catch ex As Exception
            LogError(ex)
        End Try
    End Function

    Public Function GetJavaPath() As String
        Try
            Dim strJavaPath As String = ""
            Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim subkey As Microsoft.Win32.RegistryKey = rk.OpenSubKey("SOFTWARE\\Javasoft\\Java Runtime Environment\\")
            If subkey IsNot Nothing Then
                Dim strVersion As String = subkey.GetValue("CurrentVersion")
                Dim basekey As Microsoft.Win32.RegistryKey = subkey.OpenSubKey(strVersion)
                Dim strJavaInstalledpath As String = basekey.GetValue("JavaHome").ToString()
                If strVersion >= 1.8 Then
                    If strJavaPath = "" Then
                        If File.Exists(strJavaInstalledpath + "\bin\javaw.exe") = True Then
                            strJavaPath = strJavaInstalledpath + "\bin\javaw "
                        End If
                    End If
                End If
            End If
            Return strJavaPath
        Catch ex As Exception
            LogError(ex)
            Return ""
        End Try
    End Function


    'Ver 3.0.7 REQ530 end  
    ' End here 
    'Ver 2.7.8-REQ563 start
    Public Function ValidateAOCertificateNumber4B(ByVal AOCertificateNumber) As Boolean
        Dim strAlphabets As String
        Dim j As Integer

        strAlphabets = "GH"

        ValidateAOCertificateNumber4B = True

        If Trim(Len(AOCertificateNumber)) <> 10 Then
            ValidateAOCertificateNumber4B = False
            Exit Function
        End If

        For j = 1 To 1
            If InStr(1, strAlphabets, Mid(AOCertificateNumber, j, 1)) = 0 Then
                ValidateAOCertificateNumber4B = False
                Exit Function
            End If
        Next

        For j = 2 To 10
            If Not IsNumeric(Mid(AOCertificateNumber, j, 9)) Then
                ValidateAOCertificateNumber4B = False
                Exit Function
            End If
        Next

    End Function
    Private Sub LogError(ex As Exception)
        Dim message As String = String.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        message += Environment.NewLine
        message += "-----------------------------------------------------------"
        message += Environment.NewLine
        message += String.Format("Message: {0}", ex.Message)
        message += Environment.NewLine
        message += String.Format("StackTrace: {0}", ex.StackTrace)
        message += Environment.NewLine
        message += String.Format("Source: {0}", ex.Source)
        message += Environment.NewLine
        message += String.Format("TargetSite: {0}", ex.TargetSite.ToString())
        message += Environment.NewLine
        message += "-----------------------------------------------------------"
        message += Environment.NewLine

        If System.IO.File.Exists(Application.StartupPath + "\ErrorLog.txt") Then
            GoTo lblNextProcess
        Else
            System.IO.File.Create(Application.StartupPath + "\ErrorLog.txt").Close()
        End If
lblNextProcess:
        Dim path As String = (Application.StartupPath + "\ErrorLog.txt")
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(message)
            writer.Close()
        End Using
    End Sub
    'Ver 2.7.8-REQ563 end 
    'Ver 3.1.1-REQ608 start
    Public Function SpacialCharacterValidation(ByVal strVal As String, ByVal strMsgbox As String, Optional ByVal ControlName As Control = Nothing, Optional ByVal TABName As TabControl = Nothing, Optional ByVal TabPage As Integer = 0) As Boolean
        Dim strSpeCharacter As String
        Dim SpclCharValidation As Boolean
        Dim i As Integer

        strSpeCharacter = "`~!@#$%^&*( )_+,./?;:'""[{]}\|"
        SpclCharValidation = True

        If strVal.Length() > 0 Then
            For i = 1 To strVal.Length()
                If InStr(strSpeCharacter, Mid(strVal, i, 1)) = 0 Then
                    SpclCharValidation = False
                    Exit For
                End If
            Next

            If SpclCharValidation = True Then
                MsgBox(strMsgbox, MsgBoxStyle.Critical)
                If Not ControlName Is Nothing Then
                    TABName.SelectedTab = TABName.TabPages(TabPage)
                    ControlName.Focus()
                End If
            End If
        Else
            SpclCharValidation = False
        End If
        Return SpclCharValidation
    End Function
    'Ver 3.1.1-REQ608 end
    'Ver 3.1.5-REQ648 Start
    Public Function SpecialCharacterAndOnlyZeros(ByVal strVal As String) As Boolean
        Dim strSpeCharacter As String
        Dim SpclCharValidation As Boolean
        Dim i, j, sum As Integer

        strSpeCharacter = "`~!@#$%^&*( )_+,./?;:'""[{]}\|-"
        SpclCharValidation = False

        If strVal.Length() > 0 Then
            For i = 1 To strVal.Length()
                If InStr(strSpeCharacter, Mid(strVal, i, 1)) > 0 Then
                    SpclCharValidation = True
                    Exit For
                End If
            Next

            If SpclCharValidation = False And Not UCase(strVal) Like "*[A-Z]*" Then
                For j = 1 To strVal.Length()
                    If InStr(strSpeCharacter, Mid(strVal, j, 1)) = 0 Then
                        sum = sum + Mid(strVal, j, 1)
                    End If
                Next
                If sum = 0 Then
                    SpclCharValidation = True
                End If
            End If
        Else
            SpclCharValidation = True
        End If


        Return SpclCharValidation
    End Function
    'Ver 3.1.5-REQ648 end 

    'Ver 3.1.7-REQ687 start
    Public Function PaymentTypeRate()
        On Error GoTo ErrorHead

        Dim cmdTds As New SqlCommand
        cmdTds.Connection = conTdsPac

        Call UpdateSection("if not exists (select * from sys.all_objects where type ='U' and name ='PaymentTypeRate') Create Table PaymentTypeRate (FinYear varchar(10),PaymentType varchar(10),EduCessRate decimal(18,2),SurchargeRate decimal(18,2))", cmdTds)

        If Val(Mid(StrFinYear, 1, 4)) = "2015" Then
            Call UpdateSection("if not exists (select * from PaymentTypeRate where FinYear='" & StrFinYear & "' and PaymentType='PIS') Insert into PaymentTypeRate values('" & StrFinYear & "','PIS',3,12)", cmdTds)
            'Ver 3.18-REQ726 start
            'ElseIf Val(Mid(StrFinYear, 1, 4)) = "2016" Then
        ElseIf Val(Mid(StrFinYear, 1, 4)) >= "2016" Then
            'Ver 3.18-REQ726 end
            Call UpdateSection("if not exists (select * from PaymentTypeRate where FinYear='" & StrFinYear & "' and PaymentType='PIS') Insert into PaymentTypeRate values('" & StrFinYear & "','PIS',3,15)", cmdTds)
        End If


ErrorHead:
        Exit Function
    End Function
    'Ver 3.1.7-REQ687 end 

    'Ver 3.20-REQ761 start
    Public Function SpacialCharacterValide(ByVal strVal As String, ByVal strMsgbox As String, Optional ByVal ControlName As Control = Nothing, Optional ByVal TABName As TabControl = Nothing, Optional ByVal TabPage As Integer = 0) As Boolean
        Dim strSpeCharacter As String
        Dim SpclCharValidation As Boolean
        Dim i As Integer

        strSpeCharacter = "`~!@#$%^&*( )_+,./?;:'""[{]}\|"
        SpclCharValidation = False

        If strVal.Length() > 0 Then
            For i = 1 To strVal.Length()
                If InStr(strSpeCharacter, Mid(strVal, i, 1)) > 1 Then
                    SpclCharValidation = True
                    Exit For
                End If
            Next

            If SpclCharValidation = True Then
                MsgBox(strMsgbox, MsgBoxStyle.Critical)
                If Not ControlName Is Nothing Then
                    TABName.SelectedTab = TABName.TabPages(TabPage)
                    ControlName.Focus()
                    Return SpclCharValidation
                End If
            End If

            If strVal.Length() < 15 Then
                MsgBox("Enter valid GSTN !", MsgBoxStyle.Critical)
                If Not ControlName Is Nothing Then
                    TABName.SelectedTab = TABName.TabPages(TabPage)
                    ControlName.Focus()
                    SpclCharValidation = True
                    Return SpclCharValidation
                End If
            End If
        Else
            SpclCharValidation = False
        End If
        Return SpclCharValidation
    End Function
    'Ver 3.20-REQ761 end
End Module