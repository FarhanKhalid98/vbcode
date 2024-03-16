VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmSkinSelection 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4088
      TabIndex        =   0
      Top             =   8985
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Select"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSkinSelection.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   5393
      TabIndex        =   1
      Top             =   8985
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSkinSelection.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnNext 
      Height          =   435
      Left            =   5378
      TabIndex        =   3
      Top             =   8415
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      TX              =   "Next"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSkinSelection.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnPrevious 
      Height          =   435
      Left            =   4088
      TabIndex        =   4
      Top             =   8415
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      TX              =   "Previous"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "FrmSkinSelection.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Image ImgForm 
      Height          =   2265
      Left            =   7808
      Stretch         =   -1  'True
      Top             =   5565
      Width           =   2775
   End
   Begin VB.Image ImgDesktop 
      Height          =   1875
      Left            =   7778
      Stretch         =   -1  'True
      Top             =   3525
      Width           =   2775
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   270
      Width           =   2010
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11003
      Top             =   2505
      Width           =   330
   End
End
Attribute VB_Name = "FrmSkinSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New Recordset
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String
Dim strsql As String
Dim X As Integer

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub ShowPic()
   ' strsql = "SELECT * FROM Pictures"
   ' If Rs.State = adStateOpen Then Rs.Close
   ' Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
    
   ' Rs.AbsolutePosition = i
    
    DataFile = 1
    
    Open vTmp & "\PicTemp" For Binary Access Write As DataFile
        Fl = Rs("Desktop").ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = Rs("Desktop").GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = Rs("Desktop").GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = vTmp & "\PicTemp"
    ImgDesktop.Picture = LoadPicture(FileName)
    
    DataFile = 1
    
    Open vTmp & "\PicTemp" For Binary Access Write As DataFile
        Fl = Rs("Form").ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = Rs("Form").GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = Rs("Form").GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = vTmp & "\PicTemp"
    ImgForm.Picture = LoadPicture(FileName)
    Me.Picture = ImgForm.Picture
    'Rs.Close
    'Set Rs = Nothing
End Sub

Private Sub BtnNext_Click()
   Rs.MoveNext
   If Rs.EOF = True Then Rs.MoveLast: BtnNext.Enabled = False: Exit Sub
   ShowPic
   BtnPrevious.Enabled = True
End Sub

Private Sub BtnPrevious_Click()
   Rs.MovePrevious
   If Rs.BOF = True Then Rs.MoveFirst: BtnPrevious.Enabled = False: Exit Sub
   ShowPic
   BtnNext.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event 9, 1, 1, 1
    KeyCode = 0
  End If
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   'If FunValidation = False Then Exit Sub
   CN.Execute ("UPDATE Pictures Set Selection = 0 ")
   CN.Execute ("UPDATE Pictures Set Selection = 1 where ID = " & Rs.Fields("ID").Value)
   MsgBox "Your Default Skin has been Selected Successfully", vbInformation, "Information"
   ShowPicture Desktop, 1
   Unload Me
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

'Private Function FunValidation() As Boolean
'   On Error GoTo ErrorHandler
'   If Trim(TxtName.Text) = "" Then
'      MsgBox "Please specify a Company Name", vbExclamation, "Alert"
'      If TxtName.Enabled And TxtName.Visible Then TxtName.SetFocus
'      Exit Function
'   End If
'   'All Ok, now validation is success
'   FunValidation = True
'   Exit Function
'ErrorHandler:
'   Call ShowErrorMessage
'End Function

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "Skin Selection"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   strsql = "SELECT * FROM Pictures"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.Find "Selection = 1"
   ShowPic
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

