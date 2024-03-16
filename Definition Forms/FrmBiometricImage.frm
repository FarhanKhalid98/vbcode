VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmBiometricImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Biometric Image"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   270
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   135
      Width           =   2775
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   3105
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   615
   End
   Begin JeweledBut.JeweledButton BtnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   450
      TabIndex        =   2
      Top             =   3285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmBiometricImage.frx":0000
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   1770
      TabIndex        =   3
      Top             =   3285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmBiometricImage.frx":001C
      BC              =   12632256
      FC              =   0
   End
End
Attribute VB_Name = "FrmBiometricImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs As DPFPFeatureExtraction
Dim CreateTempl As DPFPEnrollment
Dim ConvertSample As DPFPSampleConversion
Dim Templ As Object

Private Sub BtnClear_Click()
   On Error GoTo ErrorHandler
   Picture1.Picture = Nothing
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnClose_Click()
   On Error GoTo ErrorHandler
   Capture.StopCapture
   ' Unload form.
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
   On Error GoTo ErrorHandler
   ' Must use hidden PictureBox to easily resize picture.
   Set HiddenPict.Picture = Pict
   Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
   Picture1.Picture = Picture1.Image
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   ' Create capture operation.
   Set Capture = New DPFPCapture
   ' Start capture operation.
   Capture.StartCapture
   ' Create DPFPFeatureExtraction object.
   Set CreateFtrs = New DPFPFeatureExtraction
   ' Create DPFPEnrollment object.
   Set CreateTempl = New DPFPEnrollment
   ' Create DPFPSampleConversion object.
   Set ConvertSample = New DPFPSampleConversion
   Set Templ = DefEmpolyees.GetTemplate
   If Templ Is Nothing Then Exit Sub
   ' Draw fingerprint image.
   'DrawPicture ConvertSample.ConvertToPicture(Templ)
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
   On Error GoTo ErrorHandler
   Dim Feedback As DPFPCaptureFeedbackEnum
   Dim i As Integer
   ' Draw fingerprint image.
   DrawPicture ConvertSample.ConvertToPicture(Sample)
   For i = 1 To 4
      ' Process sample and create feature set for purpose of enrollment.
      Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeEnrollment)
      ' Quality of sample is not good enough to produce feature set.
      If Feedback <> CaptureFeedbackGood Then Exit Sub
      ' Add feature set to template.
      CreateTempl.AddFeatures CreateFtrs.FeatureSet
   Next i
   If CreateTempl.TemplateStatus = TemplateStatusTemplateReady Then
      DefEmpolyees.SetTemplete CreateTempl.Template
      ' Template has been created, so stop capturing samples.
'      Capture.StopCapture
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

