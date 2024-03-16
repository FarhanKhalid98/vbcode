VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   495
      Left            =   4170
      TabIndex        =   3
      Top             =   1410
      Width           =   1350
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1965
      TabIndex        =   2
      Top             =   2340
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   540
      TabIndex        =   1
      Top             =   2385
      Width           =   1350
   End
   Begin VB.PictureBox Picture1 
      Height          =   2040
      Left            =   525
      ScaleHeight     =   1980
      ScaleWidth      =   3060
      TabIndex        =   0
      Top             =   135
      Width           =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String

Private Sub Command1_Click()
   ShowPic
End Sub

'------------Vb Picture Reading-----------------------

Private Sub Form_Load()
    cn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Pubs;Data Source=(local)"
    Dim strsql As String
    Set rs = Nothing
    Set rs = New Recordset
    Set rs1 = Nothing
    Set rs1 = New Recordset
    cn.Execute "delete from temp"
    strsql = "SELECT * FROM pub_info"
    rs.Open strsql, cn, adOpenStatic, adLockOptimistic
    strsql = "SELECT * FROM temp"
    rs1.Open strsql, cn, adOpenStatic, adLockPessimistic
End Sub

Private Sub ShowPic()
    DataFile = 1
    Open "pictemp" For Binary Access Write As DataFile
        Fl = rs!logo.ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = rs!logo.GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = rs!logo.GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = "pictemp"
    Picture1.Picture = LoadPicture(FileName)
    If Not rs.EOF Then rs.MoveNext
End Sub

Private Sub cmdSave_Click()
Dim vx As Integer, vy As Integer
    For vx = 1 To 20
      rs1.AddNew
      For vy = 1 To 6
         strFileNm = "G:\main barcode\EAN-0101000100018.bmp" '"c:\temp.bmp"
         SavePicture1 vy
      Next vy
      rs1.Update
    Next vx
    MsgBox "ok"
End Sub

Private Sub SavePicture1(row As Integer)
    
    DataFile = 1
    
    
    'If Dir(strFileNm) <> "" Then Kill strFileNm     'If file exists
    'SavePicture Picture1.Image, strFileNm  'spath
    Select Case row
    Case 1:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo1.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo1.AppendChunk Chunk()
          Next i
      Close DataFile
    Case 2:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo2.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo2.AppendChunk Chunk()
          Next i
      Close DataFile
   Case 3:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo3.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo3.AppendChunk Chunk()
          Next i
      Close DataFile
   Case 4:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo4.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo4.AppendChunk Chunk()
          Next i
      Close DataFile
   Case 5:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo5.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo5.AppendChunk Chunk()
          Next i
      Close DataFile
   Case 6:
      Open strFileNm For Binary Access Read As DataFile
          Fl = LOF(DataFile)   ' Length of data in file
          If Fl = 0 Then Close DataFile: Exit Sub
          Chunks = Fl \ ChunkSize
          Fragment = Fl Mod ChunkSize
          ReDim Chunk(Fragment)
          Get DataFile, , Chunk()
          rs1!logo6.AppendChunk Chunk()
          ReDim Chunk(ChunkSize)
          For i = 1 To Chunks
              Get DataFile, , Chunk()
              rs1!logo6.AppendChunk Chunk()
          Next i
      Close DataFile
   End Select
End Sub
