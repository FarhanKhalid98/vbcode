VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   1755
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643843
      CurrentDate     =   38745
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   510
      Left            =   390
      TabIndex        =   2
      Top             =   6975
      Width           =   2610
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "hhamd"
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5670
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim a(32) As String * 40
Dim n As String

Private Sub Form_Click()
   'form1.
   'Form1.PrintForm
   'LowLevelPrinting
   ObjectPrinting
   'Dim rs As New Recordset
   'rs.Open "select * from table1", CN, adOpenStatic, adLockPessimistic
   'rs.AddNew
   'rs!a = Now
   'rs.Update
   'MsgBox Format(rs!a, "dd/mm/yyyy") & "   " & Format(rs!a, "h:mm:ss tt")
   'rs.Close
End Sub

Private Sub LowLevelPrinting()
   n = "786 Self Store"
   a(0) = Space((40 - Len(n)) / 2) & n
   a(2) = "Chowk Ahla-e-Haddis, Khanewal."
   n = " Ph: 065-2558457"
   a(3) = Space((40 - Len(n))) & n
   a(4) = "------------------------------------------"
   n = "Time :" & Time
   a(5) = Space((40 - Len(n))) & n
   n = "Date: " & Format(Date, "ddd, MMM, dd, yyyy")
   a(6) = "------------------------------------------"
   a(7) = Space((40 - Len(n))) & n
   a(8) = "=========================================="
   'a(29) = Space(3) & n
   'Format(date,ddd, MMM, dd, yyyy)
   'time
   Open "lpt1" For Output As #1
   For i = 0 To 8
      Print #1, a(i)
   Next i
   Print #1, ""
   Print #1, ""
   Print #1, ""
   Close #1
End Sub

Private Sub ObjectPrinting()
   n = "Bismillah Self Store"
   With Printer
       .ColorMode = vbPRCMMonochrome
       .PrintQuality = vbPRPQDraft        'Low quality
       .DriverName = ""
       '.Zoom = 50
       '.CurrentY = 0
       '.CurrentX = 0
       '.Height = 5
       '.Width = 10
       '.TrackDefault = False
       .PaperSize = 5
       '.Zoom = 50
       .Font = "Arial"
       .FontBold = False
       .FontSize = 8
       For i = 0 To 27
         Printer.Print n & n
      'a(0) = Space(3) & n
      'a(i - 1) = Mid(n, i, 1)
      Next i
       '.Print "EAN Code: " & m_sBarcode
       '.FontBold = False
       'Printer.PaintPicture picEan.Image, 200, 1500
       Printer.Print
       Printer.Print
       Printer.Print
       .EndDoc
   End With
End Sub
