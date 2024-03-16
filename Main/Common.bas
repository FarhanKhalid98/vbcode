Attribute VB_Name = "Common"
Option Explicit
'--------------------------
'Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
'Public Const LOCALE_SSHORTDATE = &H1F

Public Rs As New Recordset
Public DataFile As Integer, Fl As Long, Chunks As Integer
Public Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Public Const ChunkSize As Integer = 16384
Public Const conChunkSize = 100
Public vBinDataBase As String
Public strFileNm As String
Public strsql As String
Dim X As Integer
Public vTmp As String
Public vPurchaseID As String, vPurchaseDate As String

Public Enum EntryMode
  eLogIn = 1
  eReLogin = 2
  eLogOut = 3
  eAdd = 4
  eAddNewRowByEdit = 5
  eDelete = 6
  eEdit = 7
  eEditUnSaved = 8
  eRemoveRow = 9
  eRemoveRowUnSaved = 10
  eClearSavedRecord = 11
  eClearUnSavedRecord = 12
  eCloseSavedRecord = 13
  eCloseUnSavedRecord = 14
  eAddTempRecord = 15
  eReStoreRecord = 16
End Enum
Public Enum EntryForm
  eFrmLogin = 1
  eFrmLogOut = 2
  eFrmSaleInvoicePOS = 3
  eFrmSaleInvoiceDIS = 4
  eFrmSaleReturnInvoicePOS = 5
  eFrmSaleReturnInvoiceDIS = 6
  eFrmReplacementInvoice = 7
End Enum


Public Enum FormMode
  NewMode = 1
  ChangeMode = 2
  OpenMode = 3
  SelectionMode = 4
  PostedMode = 5
End Enum
'----------------------------
Public Enum SelectAccountCaller
    ssValidate = 1
    ssFunctionKey = 2
    ssButton = 3
End Enum

Public Enum LabelEffects
    lblEffectShadow = 0
    lblEffectBorder = 1
End Enum

Public Sub HelpLocation(Frm As Form)
   Frm.FraHelp.Top = 57
   Frm.FraHelp.Left = 492
   Frm.LblHelp.Top = 39
   Frm.LblHelp.Left = 744
End Sub

Public Sub AddLabelEffect(Frm As Form, _
    iEffectSize As Integer, _
    cEffectColor As ColorConstants, _
    cTextColor As ColorConstants, _
    Optional LabelEffect As LabelEffects = lblEffectShadow)
    
    Dim iMoveRight As Integer
    Dim iMoveDown As Integer
    Dim iCounter As Integer
    Dim iIterator As Integer
    Dim iStep As Integer
    
    'Unload all existing occurences of the Label
    For iIterator = 1 To Frm.LblCaption.Count - 1
        Unload Frm.LblCaption(iIterator)
    Next iIterator
    
    If LabelEffect = lblEffectShadow Then
        'Add Shadow (i.e. New labels to the top and left
        For iIterator = 1 To iEffectSize
            Load Frm.LblCaption(iIterator)
            Frm.LblCaption(iIterator).ForeColor = cEffectColor
            Frm.LblCaption(iIterator).Left = Frm.LblCaption(iIterator - 1).Left - 1
            Frm.LblCaption(iIterator).Top = Frm.LblCaption(iIterator - 1).Top - 1
            Frm.LblCaption(iIterator).Visible = True
        Next iIterator
    ElseIf LabelEffect = lblEffectBorder Then
        'Add a border (i.e. new labels all around the existing label.
        iCounter = 1
        'Me.Show
        For iMoveRight = -1 To 1
            For iMoveDown = -1 To 1
                For iIterator = 1 To iEffectSize
                    Load Frm.LblCaption(iCounter)
                    Frm.LblCaption(iCounter).Left = Frm.LblCaption(0).Left + (iIterator * iMoveRight)
                    X = Frm.LblCaption(0).Left
                    X = Frm.LblCaption(iCounter).Left
                    Frm.LblCaption(iCounter).Top = Frm.LblCaption(0).Top + (iIterator * iMoveDown)
                    Frm.LblCaption(iCounter).ForeColor = cEffectColor
                    Frm.LblCaption(iCounter).Visible = True
                    iCounter = iCounter + 1
                Next iIterator
            Next iMoveDown
        Next iMoveRight
    End If
    Frm.LblCaption(0).ForeColor = cTextColor
End Sub

Public Function SelfRound(n As Double) As Double
   On Error GoTo ErrorHandler
   Dim a As Double
   a = n - Int(n)
   If a >= 0.5 Then
      SelfRound = Int(n) + 1
   Else
      SelfRound = Int(n)
   End If
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

Public Sub ShowErrorMessage()
  MsgBox Err.Description, vbCritical, "Error occured"
End Sub

Public Function ChildDataExists(TableToCheck As String, KeyToCheck As String, Optional ExcludeTables As String, Optional AdvancedKey As String) As String
  On Error GoTo ErrorHandler
  KeyToCheck = Replace(KeyToCheck, " ", "")
  'Calling convention of this function:
  '  ChildDataExists("Parties","PartyID='62102'","ChartOfAccounts")
  '  ChildDataExists("DebitVouchers","VoucherNo=3 AND VoucherDate='2/2/2004'","DebitVouchersBody,VoucherReconciliation")
  Dim RsFKeys As ADODB.Recordset
  Dim RsDummy As ADODB.Recordset
  Dim vExc
  Dim vExcFilter As String
  Dim vCounter As Integer
  Set RsFKeys = CN.Execute("EXEC sp_fkeys @pktable_name = N'" & TableToCheck & "'")
  Set RsDummy = RsFKeys.Clone
  If Trim(ExcludeTables) <> "" Then vExc = Split(ExcludeTables, ",")
  If Not IsEmpty(vExc) Then
  For vCounter = LBound(vExc) To UBound(vExc)
    If vExcFilter <> "" Then vExcFilter = vExcFilter & " AND "
    vExcFilter = vExcFilter & "FKTABLE_NAME <> '" & vExc(vCounter) & "'"
  Next
  End If
  'All the excluded tables will be filtered
  Dim modifiedKeyToCheck As String
  RsFKeys.Filter = vExcFilter
  RsDummy.Filter = vExcFilter
  If vExcFilter = "" Then vExcFilter = "FKTABLE_NAME <> '' "
  Do Until RsFKeys.EOF
    modifiedKeyToCheck = KeyToCheck
    RsDummy.Filter = vExcFilter & " AND (FKTABLE_NAME = '" & RsFKeys!FKTABLE_NAME & "')"
    Do Until RsDummy.EOF
      If InStr(1, UCase(KeyToCheck), UCase(RsDummy!PKColumn_Name) & "=") > 0 Then
        modifiedKeyToCheck = Replace(UCase(KeyToCheck), UCase(RsDummy!PKColumn_Name), RsDummy!FKColumn_Name)
      End If
      RsDummy.MoveNext
    Loop
    '
    If CN.Execute("Select TOP 1 * From " & RsFKeys!FKTABLE_NAME & " Where " & modifiedKeyToCheck).RecordCount > 0 Then
      ChildDataExists = RsFKeys!FKTABLE_NAME
      Exit Function
    End If
    RsFKeys.MoveNext
  Loop
  Exit Function
ErrorHandler:
  Call ShowErrorMessage
End Function

Public Sub NonNumeric(iAscii As Integer, txt As TextBox, Optional IsDotAllowed As Boolean = False)
    If (iAscii > 57 Or iAscii < 48) And iAscii <> 8 Then
        If iAscii = 46 And IsDotAllowed = True Then
            If InStr(1, txt.Text, ".") <> 0 Then iAscii = 0
        Else
            iAscii = 0
        End If
    End If
End Sub

Public Sub ShowPicture(Frm As Form, j As Integer)
    On Error GoTo ErrorHandler
    strsql = "SELECT * FROM Pictures where Selection = 1"
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount = 0 Then Exit Sub
    DataFile = 1
    
    Open vTmp & "\PicTemp" For Binary Access Write As DataFile
        Fl = Rs(j).ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = Rs(j).GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = Rs(j).GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = vTmp & "\PicTemp"
    Frm.Picture = LoadPicture(FileName)
    Rs.Close
    Set Rs = Nothing
    Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Public Sub SavePicture1()
   strsql = "SELECT * FROM Pictures"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.AddNew
   strFileNm = App.Path & "\Form.jpg"
   DataFile = 1
   Close DataFile
   Open strFileNm For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs!pic.AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs!pic.AppendChunk Chunk()
       Next i
   Close DataFile
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub

Public Function FunSecurityCheck() As Boolean
'   On Error GoTo ErrorHandler
   Dim vCount As String
   Dim vMin As String
   Dim vCurrentDate As Date
   Dim vPreviousDate As Date
   Dim vExpiryDate As Date
   Dim vDate As Date
   Dim vCDate As Date
   Dim vCTime As Date
   Dim vSTime As Date
   
   CN.CursorLocation = adUseClient
   FunSecurityCheck = True
   vEncryptionString = CN.Execute("Select CompanyName from company").Fields(0).Value
   vEncryptionString = Left(vEncryptionString, InStr(1, vEncryptionString, " ") - 1)
   With CN.Execute("select * from Counter")
      If .RecordCount > 0 Then
         vCount = .Fields(0).Value
         vMin = .Fields(1).Value
         vCurrentDate = RC4(FromHexDump(EStr(.Fields(2).Value, False)), vEncryptionString)
         vPreviousDate = RC4(FromHexDump(EStr(.Fields(3).Value, False)), vEncryptionString)
         vExpiryDate = RC4(FromHexDump(EStr(.Fields(4).Value, False)), vEncryptionString)
         vSTime = Format(.Fields("CurrentDate").Value, "hh:mm")
         vDate = CN.Execute("select GetDate()").Fields(0).Value
         vCDate = DateValue(vDate)
         vCTime = Format(vDate, "hh:mm")
         'If Val(vCount) = 0 Then
      End If
   End With
'   CN.Execute "Insert into Watch(ErrorFrom,Narration) values ('Security Computer = " & LocalComputerName & "','vCDate = " & vCDate & ", vCTime = " & vCTime & " vMin = " & vMin & " vCurrentDate = " & vCurrentDate & "' )"
   vCount = Val(vCount) - 1
   If vCurrentDate <> vCDate Then
      If vCurrentDate < vCDate Then
         If DateDiff("d", vCurrentDate, vCDate) <= 500 Then
            vCurrentDate = vCDate 'After Review
            vCount = 0
            vMin = 0
         ElseIf DateDiff("d", vCDate, vExpiryDate) > 0 Then
            MsgBox "Please Contact the vendor.", vbCritical + vbOKOnly, "Soft Inn"
            vMin = 0
            FunSecurityCheck = False
         ElseIf vCDate > CDate("06-30-2016") Then
            MsgBox "Please Contact the vendor.", vbCritical + vbOKOnly, "Soft Inn"
            vMin = 0
            FunSecurityCheck = False
         End If
      ElseIf vCurrentDate = vCDate Then
         If vCTime <> vSTime Then
            vMin = Val(vMin) + 1
         End If
         If Val(vMin) > 1440 Then
            vCurrentDate = DateAdd("d", 1, vCurrentDate)
            vMin = 0
         End If
      Else
         If vPreviousDate <> vCDate Then
            vPreviousDate = vCDate
            vMin = 0
         Else
            vMin = Val(vMin) + 1
            If Val(vMin) > 1440 Then
               vCurrentDate = DateAdd("d", 1, vCurrentDate)
               vMin = 0
            End If
         End If
      End If
   Else
      If vCTime <> vSTime Then
         vMin = Val(vMin) + 1
      End If
      If Val(vMin) > 1440 Then
         vCurrentDate = DateAdd("d", 1, vCurrentDate)
         vMin = 0
      End If
   End If
   If DateDiff("d", vCurrentDate, vExpiryDate) <= 0 Then
'      CN.Execute "Insert into Watch(ErrorFrom,Narration) values ('Security','vCurrentDate = " & vCurrentDate & ", vExpiryDate = " & vExpiryDate & "')"
      CN.Execute "update Court set Type = '" & EStr(ToHexDump(RC4(EStr("¹Ïßäàâ", False), EStr("àÝÕäÚàá", False))), True) & "'" ' where SID = '" & b & "'"
      FunSecurityCheck = False
      'Timer2.Enabled = False
      Exit Function
   End If
   
   CN.Execute "update counter set Counter = '" & vCount & "'"
   CN.Execute "update counter set CMin = '" & vMin & "'"
   CN.Execute "update Counter set CLog = '" & EStr(ToHexDump(RC4(CStr(Month(vCurrentDate) & "/" & Day(vCurrentDate) & "/" & Year(vCurrentDate)), vEncryptionString)), True) & "'"
   CN.Execute "update Counter set PLog = '" & EStr(ToHexDump(RC4(CStr(Month(vPreviousDate) & "/" & Day(vPreviousDate) & "/" & Year(vPreviousDate)), vEncryptionString)), True) & "'"
   CN.Execute "update Counter set CurrentDate = '" & vDate & "'"
   Exit Function
ErrorHandler:
   FunSecurityCheck = False
   Call ShowErrorMessage
End Function

Public Function EncryptStr(myString As String, EncryPt As Boolean) As String
   Dim i As Integer
   Dim myPwd As String
   For i = 1 To Len(myString)
       If EncryPt = True Then
           myPwd = Chr(Asc(Mid(myString, i, i)) + 70 + Asc(i))
       ElseIf EncryPt = False Then
           myPwd = Chr(Asc(Mid(myString, i, i)) - 70 - Asc(i))
       End If
       EncryptStr = EncryptStr & myPwd
   Next
End Function

Public Sub Triggers(Flag As Boolean)
   On Error GoTo ErrorHandler
   Dim sql As String
   Dim i As Integer
   sql = "select Name from sysobjects where type = 'TR' order by name"
   With CN.Execute(sql)
      If .RecordCount > 0 Then
         sql = ""
         While Not .EOF
            If Flag = True Then
               sql = sql + "ALTER TABLE " + Right(!Name, Len(!Name) - InStr(1, !Name, "_")) + " ENABLE TRIGGER " + !Name + " GO "
            Else
               sql = sql + "ALTER TABLE " + Right(!Name, Len(!Name) - InStr(1, !Name, "_")) + " DISABLE TRIGGER " + !Name + " GO "
            End If
            .MoveNext
         Wend
      End If
   End With
   Dim vCommands() As String
   vCommands = Split(sql, "GO")
   For i = 0 To UBound(vCommands)
       CN.Execute vCommands(i)
   Next i
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

' Return a string of words to represent this
' currency value in Rupees.
Public Function Words_Money_Only(ByVal num As Currency) As _
    String
Dim dollars As Currency
Dim cents As Integer
Dim dollars_result As String
Dim cents_result As String

    ' Rupees.
    dollars = Int(num)
    dollars_result = Words_1_all(dollars)
    If Len(dollars_result) = 0 Then dollars_result = "zero"
Words_Money_Only = dollars_result
End Function
' Return a string of words to represent this
' currency value in Rupees and Paisas.
Public Function Words_Money_Paisa(ByVal num As Currency) As _
    String
Dim dollars As Currency
Dim cents As Integer
Dim dollars_result As String
Dim cents_result As String

    ' Rupees.
    dollars = Int(num)
    dollars_result = Words_1_all(dollars)
    If Len(dollars_result) = 0 Then dollars_result = "zero"

    If dollars_result = "one" Then
        dollars_result = dollars_result & " Rupya "
    Else
        dollars_result = dollars_result & " Rupees"
    End If

    ' Paisa.
    cents = CInt((num - dollars) * 100#)
    cents_result = Words_1_all(cents)
    If Len(cents_result) = 0 Then cents_result = "zero"

    If cents_result = "one" Then
        cents_result = cents_result & " Paisa"
    Else
        cents_result = cents_result & " Paisa"
    End If

     'Combine the results with cent
      Words_Money_Paisa = dollars_result & _
           " and " & cents_result
End Function

' Return a string of words to represent the
' integer part of this value.
Public Function Words_1_all(ByVal num As Currency) As _
    String
Dim power_value(1 To 5) As Currency
Dim power_name(1 To 5) As String
Dim digits As Integer
Dim result As String
Dim i As Integer

    ' Initialize the power names and values.
    power_name(1) = "trillion": power_value(1) = _
        1000000000000#
    power_name(2) = "billion":  power_value(2) = 1000000000
    power_name(3) = "million":  power_value(3) = 1000000
    power_name(4) = "thousand": power_value(4) = 1000
    power_name(5) = "":         power_value(5) = 1

    For i = 1 To 5
        ' See if we have digits in this range.
        If num >= power_value(i) Then
            ' Get the digits.
            digits = Int(num / power_value(i))

            ' Add the digits to the result.
            If Len(result) > 0 Then result = result & ", "
            result = result & _
                Words_1_999(digits) & _
                " " & power_name(i)

            ' Get the number without these digits.
            num = num - digits * power_value(i)
        End If
    Next i

    Words_1_all = Trim$(result)
End Function

Public Function Words_1_999(ByVal num As Integer) As String
Dim hundreds As Integer
Dim remainder As Integer
Dim result As String
    hundreds = num \ 100
    remainder = num - hundreds * 100

    If hundreds > 0 Then
        result = Words_1_19(hundreds) & " hundred "
    End If

    If remainder > 0 Then
        result = result & Words_1_99(remainder)
    End If

    Words_1_999 = Trim(result)
End Function

' Return a word for this value between 1 and 19.
Public Function Words_1_19(ByVal num As Integer) As String
    Select Case num
        Case 1
            Words_1_19 = "one"
        Case 2
            Words_1_19 = "two"
        Case 3
            Words_1_19 = "three"
        Case 4
            Words_1_19 = "four"
        Case 5
            Words_1_19 = "five"
        Case 6
            Words_1_19 = "six"
        Case 7
            Words_1_19 = "seven"
        Case 8
            Words_1_19 = "eight"
        Case 9
            Words_1_19 = "nine"
        Case 10
            Words_1_19 = "ten"
        Case 11
            Words_1_19 = "eleven"
        Case 12
            Words_1_19 = "twelve"
        Case 13
            Words_1_19 = "thirteen"
        Case 14
            Words_1_19 = "fourteen"
        Case 15
            Words_1_19 = "fifteen"
        Case 16
            Words_1_19 = "sixteen"
        Case 17
            Words_1_19 = "seventeen"
        Case 18
            Words_1_19 = "eightteen"
        Case 19
            Words_1_19 = "nineteen"
    End Select
End Function

Public Function Words_1_99(ByVal num As Integer) As String
Dim result As String
Dim tens As Integer

    tens = num \ 10

    If tens <= 1 Then
        ' 1 <= num <= 19
        result = result & " " & Words_1_19(num)
    Else
        ' 20 <= num
        ' Get the tens digit word.
        Select Case tens
            Case 2
                result = "twenty"
            Case 3
                result = "thirty"
            Case 4
                result = "forty"
            Case 5
                result = "fifty"
            Case 6
                result = "sixty"
            Case 7
                result = "seventy"
            Case 8
                result = "eighty"
            Case 9
                result = "ninety"
        End Select

        ' Add the ones digit number.
        result = result & " " & Words_1_19(num - tens * 10)
    End If

    Words_1_99 = Trim$(result)
End Function

