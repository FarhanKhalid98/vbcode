VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Attendance"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   3375
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   765
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   540
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   585
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs As DPFPFeatureExtraction
Dim ConvertSample As DPFPSampleConversion
Dim Verify As DPFPVerification
Dim Templ As Object
Dim Rs As New Recordset

Private Sub BtnClear_Click()
   Picture1.Picture = Nothing
End Sub

Private Sub BtnClose_Click()
 Capture.StopCapture
 ' Unload form.
 Unload Me
End Sub

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
 ' Must use hidden PictureBox to easily resize picture.
 Set HiddenPict.Picture = Pict
 Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
 Picture1.Picture = Picture1.Image
End Sub

Private Sub Form_Load()
   ' Create capture operation.
   Set Capture = New DPFPCapture
   ' Start capture operation.
   Capture.StartCapture
   ' Create DPFPFeatureExtraction object.
   Set CreateFtrs = New DPFPFeatureExtraction
   ' Create DPFPVerification object.
   Set Verify = New DPFPVerification
   ' Create DPFPSampleConversion object.
   Set ConvertSample = New DPFPSampleConversion
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open "select * from employees where BiometricPattern is not null", CN, adOpenStatic, adLockOptimistic
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
   Dim Feedback As DPFPCaptureFeedbackEnum
   Dim Res As DPFPVerificationResult
   ' Dim Templ As Object
   DrawPicture ConvertSample.ConvertToPicture(Sample)
   ' Process sample and create feature set for purpose of verification.
   Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeVerification)
   ' Quality of sample is not good enough to produce feature set.
   If Feedback = CaptureFeedbackGood Then
      Rs.MoveFirst
      For i = 1 To Rs.RecordCount
         With CN.Execute("Select BiometricPattern from Employees where EmpID = '" & Rs!EmpID & "'")
            If .RecordCount > 0 Then
               Set Templ = New DPFPTemplate
               ' Import binary data to template.
               Templ.Deserialize .Fields(0).GetChunk(.Fields(0).ActualSize)
               ' Compare feature set with template.
               Set Res = Verify.Verify(CreateFtrs.FeatureSet, Templ)
               If Res.Verified = True Then
                  Call SubAttendance(Rs!EmpID)
                  'MsgBox "The fingerprint was verified. And EmpID = " & Rs!EmpID
                  Exit For
               End If
            End If
         .Close
         End With
         Rs.MoveNext
      Next i
   End If
End Sub

Private Sub SubAttendance(ByVal EmpID As String)
   Dim vSQL As String
   With CN.Execute("select * from EmpAttendance where EmpID = '" & EmpID & "' And AttendDate = '" & Date & "'")
      If .RecordCount = 0 Then
         vSQL = "Insert into EmpAttendance (AttendID, EmpID, AttendDate, TimeIn, TimeUpdated, UserNo) values (" _
         & FunGetMaxID & ",'" & EmpID & "','" & Date & "','" & Now & "',0," & ObjUserSecurity.UserNo & ")"
         CN.Execute vSQL
         MsgBox Rs.Fields("EmpName").Value & " - " & EmpID & " is Attend in"
      Else
         If IsNull(!DateOut) Then
            vSQL = "Update EmpAttendance set DateOut = '" & Date & "', TimeOut = '" & Now & "'" & _
            " where EmpID = '" & !EmpID & "' And AttendDate = '" & !AttendDate & "'"
            CN.Execute vSQL
            vSQL = "Update EmpAttendance set WorkingTime =  dateDiff(Minute,timein,timeout) " & _
            " where EmpID = '" & !EmpID & "' And AttendDate = '" & !AttendDate & "'"
            CN.Execute vSQL
            MsgBox Rs.Fields("EmpName").Value & " is Attend Out"
         Else
            MsgBox "This Employee Already done his attendance.", vbExclamation, Me.Caption
         End If
      End If
   End With
End Sub

Private Function FunGetMaxID() As Long
   On Error GoTo ErrorHandler
   FunGetMaxID = CN.Execute("Select isnull(max(AttendID),0)+1 from EmpAttendance").Fields(0)
   Exit Function
ErrorHandler:
   Call ShowErrorMessage
End Function

