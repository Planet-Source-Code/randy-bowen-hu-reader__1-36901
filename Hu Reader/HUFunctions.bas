Attribute VB_Name = "HUFunctions"
Option Explicit
'Alot of this code came from david brewstrer HUEditor
'some of the functions has changed and to add get card info
'Hope this helps
'Rattlesnake
'P.S i whiped this up quick in 15 minutes no error checking
'or spell checking so good luck
Dim TimeOutError As Boolean
Dim ReadTimeOut
Dim GetReturnFromLoader
Dim Temp
Declare Function GetTickCount Lib "kernel32.dll" () As Long
Dim tempReturn, Return1, Return2, Return3, Return4, Return5

Public Function ResetForATR()
        WriteHU ("0410019B00")
        tempReturn = ReadHUBytes(1)
        tempReturn = ReadHUBytes("&h" & tempReturn)
        Debug.Print "1 " & tempReturn
End Function

Public Function WriteHU(sData As String)
    Dim iByte As Integer
    Dim LoaderData As String
    Dim iPacketLen As Integer
    Dim x As Integer
    iPacketLen = Len(sData)
    LoaderData = ""
    For x = 1 To iPacketLen Step 2
        iByte = CDec("&h" & Mid(sData, x, 2))
        LoaderData = LoaderData & Chr(iByte)
    Next x
    'write to string to the loader
    Form1.MSComm.Output = LoaderData
End Function

Public Function ReadHUBytes(BytesToGet As Variant) As String
    Dim RXByte As String
    Dim RXString As String
    Dim RXReturnLength As Integer
    Dim z As Integer
    
    RXString = ""
    TimeOutError = False
    ReadTimeOut = GetTickCount + 80 'milliseconds
    Do While GetTickCount < ReadTimeOut And Form1.MSComm.InBufferCount < BytesToGet
        DoEvents
    Loop
    Form1.MSComm.InputLen = BytesToGet
    'read the return from the loader
    GetReturnFromLoader = Form1.MSComm.Input
    RXReturnLength = Len(GetReturnFromLoader)
    If RXReturnLength = 0 Then
        GetReturnFromLoader = Chr("0")
        RXReturnLength = 1
    End If
    'convert the retrun to txt format
    For z = 1 To RXReturnLength
        RXByte = Mid(GetReturnFromLoader, z, 1)
        Temp = HexString(Asc(RXByte), 1)
        If Len(Temp) < 2 Then Temp = "0" & Temp
        RXString = RXString & Temp
    Next z
    DoEvents
    ReadHUBytes = RXString
End Function

Function HexString(Number, Length)
    ' This function takes 2 arguments, a number and a length.  It converts the decimal
    ' number given by the first argument to a Hexidecimal string with its length
    ' equal to the number of digits given by the second argument
    Dim RetVal
    Dim CurLen
    RetVal = Hex(Number)
    CurLen = Len(RetVal)
    If CurLen < Length Then
        RetVal = String(Length - CurLen, "0") & RetVal
    End If
    HexString = RetVal
End Function

Function HexToDec(HexNumber)
    ' This function takes a string as input, assuming it to be a Hexidecimal string,
    ' and converts it to a decimal number.
    HexNumber = Replace(UCase(HexNumber), " ", "")
    HexToDec = CLng("&H" + HexNumber)
End Function

Public Function IsCardPresent() As Boolean
    WriteHU ("80")
    DelaySeconds (0.5)
    tempReturn = ReadHUBytes(1)
    Select Case tempReturn
        Case Is = "FF"
            IsCardPresent = True
        Case Else
            IsCardPresent = False
    End Select
End Function

Public Sub DelaySeconds(ByVal seconds As Single)
    Static start As Single
    start = Timer
    Do While Timer < start + seconds
        DoEvents
    Loop
End Sub

Public Sub GreenOn()
    WriteHU ("A1")
    Call DelaySeconds(0.25)
End Sub

Public Function GreenOff() As String
    WriteHU ("A0")
    Call DelaySeconds(0.25)
End Function

Public Sub RedOn()
    WriteHU ("A2")
    Call DelaySeconds(0.25)
End Sub

Public Function RedOff() As String
    WriteHU ("A3")
    Call DelaySeconds(0.25)
End Function

Public Function GetAtmelVersion() As String
On Error Resume Next
Call GreenOn
    WriteHU ("90")
    tempReturn = ReadHUBytes(4)
    GetAtmelVersion = Chr(CDec("&h" & Left(tempReturn, 2)))
    If tempReturn <> "00" Then
        GetAtmelVersion = GetAtmelVersion & Chr(CDec("&h" & Mid(tempReturn, 3, 2)))
        GetAtmelVersion = GetAtmelVersion & Chr(CDec("&h" & Mid(tempReturn, 5, 2)))
        GetAtmelVersion = GetAtmelVersion & Chr(CDec("&h" & Mid(tempReturn, 7, 2)))
    End If
Call GreenOff
End Function
Public Function GetATR() As String
Dim x
    '06 = 6 bytes to follow
    '10 = Set ATR Baud Rate
    '0E 10 = set timeout timer
    '01 = RESET CARD
    '93 = read 20 bytes from card
    '00 = Execute above inst
    Call GreenOn
    WriteHU ("A106100E10019300")
    'Call DelaySeconds(0.5)
    tempReturn = ReadHUBytes(1)
    'Call DelaySeconds(0.5)
    tempReturn = ReadHUBytes(1)
    'Call DelaySeconds(0.5)
    tempReturn = ReadHUBytes(CDec("&h" & tempReturn))
    GetATR = ""
    If Len(tempReturn) > 2 Then
        For x = 1 To Len(tempReturn) Step 2
            GetATR = GetATR & " " & Mid(tempReturn, x, 2)
        Next x
    Else
        GetATR = "BAD ATR"
    End If
    tempReturn = ReadHUBytes(14)
    Call GreenOff
End Function

Public Function GetSTARTUPINFO() As String
Dim PostATR, PostAtmelVersion, ATRHeader, Line1, CAMIDHex, _
CAMIDDec, USWHex, USWDec, DSWHex, DSWDec, GuideHex, GuideDec, _
TimeHex, TimeDec, RatingHex, RatingDec, SpendingLimitHex, _
SpendingLimitDec, actyear, actmonth, ActDateHex, ActDateDec, _
PPVLimitHex, PPVLimitDec, SpentHex, SpentDec, Fuse, IRDHex, _
IRDDec, RetValue, FuseHexString1, FuseDecString2, _
IRDDecString1, IRDDecString2, IRDDecString3, IRDDecString4, _
IRDDecString5, IRDDecString6, IRDDecString7, IRDDecString8, _
temp1, temp2, actday As String
Dim x As Integer

GetSTARTUPINFO = ""
PostATR = Form1.Label2.Caption
        '06 = Number of bytes to follow
        '482A000080 = the 2A Command
        '00 = Execute
        Call GreenOn
        WriteHU ("0A1A0E10C4482A000082BF00")
        Return1 = ReadHUBytes(1)
        Return1 = ReadHUBytes("&H02" & Return1)
        Debug.Print "1 " & Return1
        WriteHU ("02BF00")
        Return2 = ReadHUBytes(1)
        Return2 = ReadHUBytes("&H02" & Return2)
        Debug.Print "2 " & Return2
        WriteHU ("028200")
        Return3 = ReadHUBytes(1)
        Return3 = ReadHUBytes("&H02" & Return3)
        Debug.Print "3 " & Return3
        WriteHU ("08c448580000179B00")
        Return4 = ReadHUBytes(1)
        Return4 = ReadHUBytes("&H02" & Return4)
        Debug.Print "4 " & Return4
        ' assign important parts of returned from Ins2A data here

        tempReturn = Mid$(Return1, 63, 4)
        USWHex = tempReturn
        USWDec = HexToDec(USWHex)
        tempReturn = Mid$(Return1, 71, 4)
        DSWHex = tempReturn
        DSWDec = HexToDec(DSWHex)
        
        
        tempReturn = Mid$(Return1, 27, 2)
        RatingHex = tempReturn

   Select Case RatingHex
    Case Is = "00"
     RatingHex = tempReturn & "  All Locked"
    Case Is = "01"
     RatingHex = tempReturn & "  NR"
    Case Is = "02"
     RatingHex = tempReturn & "  G"
    Case Is = "04"
     RatingHex = tempReturn & "  PG"
    Case Is = "06"
     RatingHex = tempReturn & "  PG13"
    Case Is = "07"
     RatingHex = tempReturn & "  NR Content"
    Case Is = "09"
     RatingHex = tempReturn & "  R"
    Case Is = "0B"
     RatingHex = tempReturn & "  NR Mature"
    Case Is = "0D"
     RatingHex = tempReturn & "  NC17"
    Case Else
     RatingHex = tempReturn
  End Select

        tempReturn = Mid$(Return1, 29, 4)
        SpendingLimitHex = tempReturn
        SpendingLimitDec = "$" & Int(HexToDec(SpendingLimitHex) / 100) & ".00"

        tempReturn = Mid$(Return2, 49, 4)
        PPVLimitHex = tempReturn
        PPVLimitDec = "$" & Int(HexToDec(PPVLimitHex) / 100) & ".00"

        tempReturn = Mid$(Return2, 45, 4)
        SpentHex = tempReturn
        SpentDec = "$" & Int(HexToDec(SpentHex) / 100) & ".00"

  temp1 = HexToDec(Mid$(Return1, 33, 2))
  temp2 = HexToDec(Mid$(Return1, 35, 2))
  ActDateHex = HexString(temp1, 2) & HexString(temp2, 2)
    If ActDateHex = "0000" Then
        ActDateDec = "Not Active"
    Else
        ActDateHex = HexString(temp1, 2) & HexString(temp2, 2)
        actyear = 2000 + Left(((temp1 - 95) / 12 * 100), 1)
        actday = temp2
        actmonth = Round((Right(Int((((temp1 - 95) / 12)) * 100), 2) / 100) * 12)
        For x = 1 To 9
        Next x
        If actmonth Or actday = x Then
        actmonth = "0" & actmonth
        actday = "0" & actday
        Else
        actmonth = actmonth
        actday = actday
        End If
        ActDateDec = actmonth & " / " & actday & " / " & actyear
    End If

    FuseHexString1 = Mid$(Return1, 7, 2)
    FuseDecString2 = HexToDec(Mid$(Return1, 7, 2))
       Fuse = FuseHexString1 & HexString(FuseDecString2 Xor &HFF, 2)
 
    IRDDecString1 = HexToDec(Mid$(Return1, 47, 2))
    IRDDecString2 = HexToDec(Mid$(Return1, 55, 2))
    IRDDecString3 = HexToDec(Mid$(Return1, 49, 2))
    IRDDecString4 = HexToDec(Mid$(Return1, 57, 2))
    IRDDecString5 = HexToDec(Mid$(Return1, 51, 2))
    IRDDecString6 = HexToDec(Mid$(Return1, 59, 2))
    IRDDecString7 = HexToDec(Mid$(Return1, 53, 2))
    IRDDecString8 = HexToDec(Mid$(Return1, 61, 2))
    
    IRDHex = HexString(IRDDecString1 Xor IRDDecString2, 2) _
           & HexString(IRDDecString3 Xor IRDDecString4, 2) _
           & HexString(IRDDecString5 Xor IRDDecString6, 2) _
           & HexString(IRDDecString7 Xor IRDDecString8, 2)
    If Len(IRDHex) <= 4 Then
        IRDHex = "0000" & IRDHex
    End If

    IRDDec = HexToDec(IRDHex)
    If IRDDec <= 1 Then
        IRDDec = "0000000" & IRDDec
    End If
        tempReturn = Mid$(Return1, 47, 8)
        CAMIDHex = tempReturn
        CAMIDDec = "000" & HexToDec(CAMIDHex) & "x"
        tempReturn = Mid$(Return4, 29, 2)
        GuideHex = tempReturn
        
        tempReturn = Mid$(Return4, 25, 2)
        TimeHex = tempReturn
   Select Case TimeHex
    Case Is = "A0"
     TimeHex = tempReturn & "  Pacific (DST)"
    Case Is = "A2"
     TimeHex = tempReturn & "  Mountain (DST)"
    Case Is = "A4"
     TimeHex = tempReturn & "  Central (DST)"
    Case Is = "A6"
     TimeHex = tempReturn & "  Eastern (DST)"
    Case Is = "A8"
     TimeHex = tempReturn & "  Atlantic (DST)"
    Case Is = "A9"
     TimeHex = tempReturn & "  Newfoundland (DST)"
    Case Is = "20"
     TimeHex = tempReturn & "  Pacific"
    Case Is = "22"
     TimeHex = tempReturn & "  Mountain"
    Case Is = "24"
     TimeHex = tempReturn & "  Central"
    Case Is = "26"
     TimeHex = tempReturn & "  Eastern"
    Case Is = "28"
     TimeHex = tempReturn & "  Atlantic"
    Case Is = "29"
     TimeHex = tempReturn & "  Newfoundland"
    Case Else
     TimeHex = tempReturn
  End Select
  
        Call GreenOff
        ' Print all info here
        Line1 = vbCrLf _
        & " Card Information" & vbCr _
        & "-----------------------------------------------------------------------------------------------------" & vbCr _
        & " Card ID(CAM)      " & CAMIDDec & vbCr _
        & " IRD Number        " & IRDDec & vbCr _
        & " Guide Byte          " & GuideHex & vbCr _
        & " Fuse Byte           " & Fuse & vbCr _
        & " Time Zone          " & TimeHex & vbCr _
        & " Rating                " & RatingHex & vbCr _
        & " Spending Limit    " & SpendingLimitDec & vbCr _
        & " PPV Limit           " & PPVLimitDec & vbCr _
        & " Amount Spent     " & SpentDec & vbCr _
        & " USW                 " & USWHex & vbCr _
        & " DSW                 " & DSWHex & vbCr _
        & " Activation Date   " & ActDateDec & vbCr _
        & "-----------------------------------------------------------------------------------------------------" & vbCrLf
        GetSTARTUPINFO = " Reset Successful!" & vbCrLf & Line1
End Function

