Attribute VB_Name = "PictureModule"
Option Explicit
Dim DataFile As Integer, Fl As Long, Chunks As Integer
Dim Fragment As Integer, Chunk() As Byte, i As Integer, FileName As String
Const ChunkSize As Integer = 16384
Const conChunkSize = 100
Dim strFileNm As String

Public Sub ShowPicture(Frm As Form)
    strsql = "SELECT * FROM Pictures"
    If rs.State = adStateOpen Then rs.Close
    rs.Open strsql, CN, adOpenStatic, adLockOptimistic
    DataFile = 1
    Open "pictemp" For Binary Access Write As DataFile
        Fl = rs!logo.ActualSize ' Length of data in file
        If Fl = 0 Then Close DataFile: Exit Sub
        Chunks = Fl \ ChunkSize
        Fragment = Fl Mod ChunkSize
        ReDim Chunk(Fragment)
        Chunk() = rs!Pic.GetChunk(Fragment)
        Put DataFile, , Chunk()
        For i = 1 To Chunks
            ReDim Buffer(ChunkSize)
            Chunk() = rs!Pic.GetChunk(ChunkSize)
            Put DataFile, , Chunk()
        Next i
    Close DataFile
    FileName = "pictemp"
    Frm.Picture = LoadPicture(FileName)
    If Not rs.EOF Then rs.MoveNext
End Sub

Public Sub SavePicture()
   rs.AddNew
   strFileNm = "E:\Soft Inn\Working Projects\Super Soft\Testing\Super Soft\Pictures\Form.jpg"
   DataFile = 1
   Open strFileNm For Binary Access Read As DataFile
       Fl = LOF(DataFile)   ' Length of data in file
       If Fl = 0 Then Close DataFile: Exit Sub
       Chunks = Fl \ ChunkSize
       Fragment = Fl Mod ChunkSize
       ReDim Chunk(Fragment)
       Get DataFile, , Chunk()
       rs!Pic.AppendChunk Chunk()
       ReDim Chunk(ChunkSize)
       For i = 1 To Chunks
           Get DataFile, , Chunk()
           rs!Pic.AppendChunk Chunk()
       Next i
   Close DataFile
   rs.Update
End Sub
