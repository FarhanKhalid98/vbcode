VERSION 5.00
Begin VB.Form FrmBarcodeViewer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EAN Generator"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2925
   ForeColor       =   &H8000000C&
   Icon            =   "FrmBarcodeViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   195
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   225
      ScaleHeight     =   0.365
      ScaleMode       =   5  'Inch
      ScaleWidth      =   1.104
      TabIndex        =   2
      Top             =   540
      Width           =   1590
   End
   Begin VB.TextBox txtBarcode 
      Height          =   315
      Left            =   120
      MaxLength       =   13
      TabIndex        =   1
      Tag             =   "Enter 7+ digits"
      ToolTipText     =   "Enter 7+ digits"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdEANCreate 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "Click here to generate a barcode"
      Top             =   1935
      Width           =   810
   End
   Begin VB.Menu zMenu1 
      Caption         =   "&Ean"
      Begin VB.Menu zEan 
         Caption         =   "&Generate"
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu zEan 
         Caption         =   "&Save Ean"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu zEan 
         Caption         =   "&Print"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu zEan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu zEan 
         Caption         =   "&Quit"
         Index           =   4
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmBarcodeViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As String * 1
    lfUnderline As String * 1
    lfStrikeOut As String * 1
    lfCharSet As String * 1
    lfOutPrecision As String * 1
    lfClipPrecision As String * 1
    lfQuality As String * 1
    lfPitchAndFamily As String * 1
    lfFaceName As String * 32
End Type

Public ParaInProductName As String
Public ParaInRate As Double
Public ParaInRate2 As Double
Public ParaInCompany As String
Public ParaInProductID As String
'Original spiel by original author:
'==================================
'Well, not many functions implemented here. I tried to make it as simple
'as possible even if this app loses any practical usage due it's simplicity,
'exept for the check number calculation and that's just what I needed.
'Be patient with my english,- I am not from this world.
'Any questions? Then mail me: davidsmejkal@hellada.cz

'Updated spiel by some guy ;) :
'==============================
'I've tried to make this as close to this VB naming convention file that i found
' as possible. Can somebody tell me why using short names is bad? Is the file
'size affected by the size of variable name that you use? Any comments welcome.
'
'New features:
'Includes a different checkdigit function than the original.
'Has been updated to include support for EAN 8 barcodes. SO FAR all generated
'barcodes have been good, any problems please email!
'Tried and tested at my local Somerfields :D
'Please email the updater: cmeister2@hotmail.com
'Or: cmeister2@btinternet.com (more likely to be replied to, since hotmail's
'filter goes mad on real emails to me :D
Dim m_sBarcode As String, m_lBarcodeLength As Long

Private Sub RotateText(PBCtrl As PictureBox, disptxt As String, CX, CY)
Dim Font As LOGFONT
Dim hFont As Long, ret As Long
Const FONTSIZE = 12  ' Desired point size of font

Font.lfEscapement = 900    ' 180-degree rotation
Font.lfFaceName = "Arial" + Chr$(0)
Font.lfWeight = 50

' Windows expects the font size to be in pixels and to be negative if you are specifying the character height you want.

Font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
hFont = CreateFontIndirect(Font)
SelectObject PBCtrl.hdc, hFont

PBCtrl.CurrentX = 1 ' CX
PBCtrl.CurrentY = 50 'CY
PBCtrl.Print disptxt

' Clean up by restoring original font.
ret = DeleteObject(hFont)
End Sub

Private Sub cmdEANCreate_Click()
On Error GoTo errHandler                            'Error Handling function

Dim bytBarcodeType As Byte, sTemp As String         'Initiate variables
With TxtBarCode
Select Case Len(.Text)
    Case 0 To 6:
        Alert "Enter 7+ numbers into the text box": Exit Sub    '6 or less numbers entered
    Case 7 To 11:
        bytBarcodeType = 7                                      'EAN 8 barcode
        m_lBarcodeLength = 8
    Case 12 To 20:
        bytBarcodeType = 12                                     'EAN 13 barcode
        m_lBarcodeLength = 13
End Select

'm_sBarcode = MakeBarcode(Left(.Text, bytBarcodeType))           'Puts correct checkdigit on barcode root.
'.Text = m_sBarcode                                              'Full EAN code
DrawEan                                                      'Draw the barcode!

End With
Exit Sub

errHandler:
Select Case Err.Number
    Case 13: Alert "Enter only numbers into text box!"   'In case someone puts other characters then numbers into textbox
    Case Else: Alert "Error occurred: " & Err.Description   'Any other error, die nicely
End Select
End Sub

Private Sub Form_Load()
Init                            'Initializes Mdl array - this holds the lines info!
TxtBarCode.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set FrmBarcodeViewer = Nothing        'Important!!! = remove form
End Sub

Private Sub DrawEan()
Dim sPath As String

   
BarEAN13 Picture1, 2, TxtBarCode, True

'   Picture1.Font = "Arial"  '"Verdana"
'   Picture1.FONTSIZE = 7
''   Picture1.FontBold = True
'   Picture1.CurrentX = 10
'   Picture1.CurrentY = 3 ' + BarH
'   Picture1.Print ParaInCompany
''   Picture1.CurrentY = 5
''   Picture1.CurrentX = Picture1.CurrentX + ((15 - (Len("Rs." & ParaInRate))) * 8)
''   Picture1.FontBold = False
''   If ParaInRate <> 0 Then Picture1.Print "Rs." & ParaInRate
'
'
'   Picture1.Font = "Arial"
'   Picture1.FONTSIZE = 7
'   Picture1.CurrentX = 10
'   Picture1.CurrentY = 46
'   Picture1.FontBold = False
'   'Picture1.Print ParaInProductName
'
'   Dim length, pos, a
'   length = Len(ParaInProductName)
'   If length > 23 Then
'      a = Left(ParaInProductName, 23)
'      pos = InStrRev(a, " ")
'      Picture1.Print Left(ParaInProductName, pos)
'      Picture1.CurrentX = 10
'      Picture1.CurrentY = 56
'      Picture1.Print Mid(ParaInProductName, pos + 1, 13)
'   Else
'      Picture1.Print ParaInProductName
'   End If
'   Picture1.Font = "Arial"
'   Picture1.FONTSIZE = 9
'   Picture1.FontBold = True
'   Picture1.CurrentX = Picture1.CurrentX + ((15 - (Len("Rs." & ParaInRate))) * 8)
'   Picture1.CurrentY = 60
'   If ParaInRate <> 0 Then Picture1.Print "Rs." & ParaInRate
'
'  'Picture1.Print (ParaInCompany) & Space(28 - (Len(Space(1) & ParaInCompany) + Len("Rs." & ParaInRate))) & "Rs." & ParaInRate
  
If m_lBarcodeLength <> 0 Then            'Only if EAn is drawn
    sPath = "c:\" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ParaInProductName, "/", "-"), """", "-"), "\", "-"), ".", "-"), "*", "-"), "?", "-"), "&", "-"), ":", "-") & " " & ParaInProductID & ".bmp"
    If Dir(sPath) <> "" Then Kill sPath     'If file exists
    SavePicture Picture1.Image, sPath
    'MsgBox "Ean saved as: " & Chr(34) & sPath & Chr(34)
End If
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
    Case 13: cmdEANCreate_Click: Exit Sub
    Case 8, 48 To 57: Exit Sub              'Allows only numbers to be typed
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub zEan_Click(Index As Integer)

Select Case Index
    Case 0: cmdEANCreate_Click
    Case 1
        
    Case 2
        'If m_lBarcodeLength <> 0 Then PrintEan Else Alert "No bar code to print!"
    Case 4: Unload Me
End Select
End Sub
