VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Begin VB.Form FrmAddSkin 
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
   Begin JeweledBut.JeweledButton BtnSave 
      Height          =   420
      Left            =   5993
      TabIndex        =   0
      Top             =   8985
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Save"
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
      MICON           =   "FrmAddSkin.frx":0000
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   7298
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
      MICON           =   "FrmAddSkin.frx":001C
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnForm 
      Height          =   315
      Left            =   8813
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6345
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      TX              =   "..."
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
      MICON           =   "FrmAddSkin.frx":0038
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtForm 
      Height          =   315
      Left            =   3488
      TabIndex        =   4
      Top             =   6345
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   500
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8333
      Top             =   3135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin JeweledBut.JeweledButton BtnDesktop 
      Height          =   315
      Left            =   8873
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4275
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      TX              =   "..."
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
      MICON           =   "FrmAddSkin.frx":0054
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtDesktop 
      Height          =   315
      Left            =   3548
      TabIndex        =   7
      Top             =   4275
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   500
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mandatory       =   1
   End
   Begin VB.Image ImgDesktop 
      Height          =   1875
      Left            =   9683
      Stretch         =   -1  'True
      Top             =   3585
      Width           =   2775
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Skin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2183
      TabIndex        =   8
      Top             =   4335
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Skin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2453
      TabIndex        =   5
      Top             =   6405
      Width           =   975
   End
   Begin VB.Image ImgForm 
      Height          =   2265
      Left            =   9713
      Stretch         =   -1  'True
      Top             =   5565
      Width           =   2775
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Skin"
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
      Width           =   1275
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   12908
      Top             =   2505
      Width           =   330
   End
End
Attribute VB_Name = "FrmAddSkin"
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

Private Sub SavePic()
   strsql = "SELECT * FROM Pictures"
   If Rs.State = adStateOpen Then Rs.Close
   Rs.Open strsql, CN, adOpenStatic, adLockOptimistic
   Rs.AddNew
   Rs("ID") = CN.Execute("SELECT isnull(MAX(ID),0)+1 FROM Pictures").Fields(0).Value
   
   ' Update Desktop Picture
   DataFile = 1
   Close DataFile
   Open TxtDesktop.Text For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs("Desktop").AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs("Desktop").AppendChunk Chunk()
       Next i
   Close DataFile
   
   ' Update Form Picture
   DataFile = 1
   Close DataFile
   Open TxtForm.Text For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       Rs("Form").AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           Rs("Form").AppendChunk Chunk()
       Next i
   Close DataFile
   Rs("Selection") = 1
   Rs.Update
   Rs.Close
   Set Rs = Nothing
End Sub

Private Sub BtnDesktop_Click()
   CD1.FileName = ""
   CD1.DialogTitle = "Enter Path to take Desktop Picture"
'   CD1.InitDir = App.Path
   CD1.Filter = "(Image Files)|*.bmp;*.jpg"
   CD1.ShowSave
   If CD1.FileName <> "" Then
      TxtDesktop.Text = CD1.FileName
      ImgDesktop.Picture = LoadPicture(CD1.FileName)
   Else
      CD1.FileName = ""
      ImgDesktop.Picture = Nothing
   End If
End Sub

Private Sub BtnForm_Click()
   CD1.FileName = ""
   CD1.DialogTitle = "Enter Path to take Form Picture"
'   CD1.InitDir = App.Path
   CD1.Filter = "(Image Files)|*.bmp;*.jpg"
   CD1.ShowSave
   If CD1.FileName <> "" Then
      TxtForm.Text = CD1.FileName
      ImgForm.Picture = LoadPicture(CD1.FileName)
   Else
      CD1.FileName = ""
      ImgForm.Picture = Nothing
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event 9, 1, 1, 1
    KeyCode = 0
  End If
End Sub

Private Sub BtnSave_Click()
   On Error GoTo ErrorHandler
   'If FunValidation = False Then Exit Sub
   SavePic
   MsgBox "Skin has been Added Successfully", vbInformation, "Information"
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
   SetWindowText Me.hWnd, "Add Skin"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
