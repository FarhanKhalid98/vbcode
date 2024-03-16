VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmMergeProducts 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMergeProducts.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4853
      TabIndex        =   1
      Top             =   4223
      Width           =   2490
   End
   Begin VB.TextBox TxtMerge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4853
      TabIndex        =   0
      Top             =   2888
      Width           =   2490
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Height          =   420
      Left            =   5963
      TabIndex        =   3
      Top             =   5918
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
      MICON           =   "FrmMergeProducts.frx":6971
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Height          =   420
      Left            =   4658
      TabIndex        =   2
      Top             =   5918
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Ok"
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
      MICON           =   "FrmMergeProducts.frx":698D
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Into This Code"
      Height          =   195
      Left            =   4853
      TabIndex        =   6
      Top             =   3998
      Width           =   1035
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merge Products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1935
      TabIndex        =   5
      Top             =   165
      Width           =   2235
   End
   Begin VB.Image ImgExit 
      Height          =   315
      Left            =   11625
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge This Code"
      Height          =   195
      Left            =   4853
      TabIndex        =   4
      Top             =   2663
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMergeProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vStrSQL As String

Private Sub BtnClose_Click()
  Unload Me
End Sub

Private Sub BtnSelect_Click()
   On Error GoTo ErrorHandler
   TxtSource.Text = Right("00000" + CStr(Val(TxtSource.Text)), 5)
   TxtMerge.Text = Right("00000" + CStr(Val(TxtMerge.Text)), 5)
   CN.Execute "update salebody set productid = '" & TxtSource.Text & "' where productid = '" & TxtMerge.Text & "'"
   CN.Execute "update salereturnbody set productid = '" & TxtSource.Text & "' where productid = '" & TxtMerge.Text & "'"
   CN.Execute "update purchasebody set productid = '" & TxtSource.Text & "' where productid = '" & TxtMerge.Text & "'"
   CN.Execute "update purchasereturnbody set productid = '" & TxtSource.Text & "' where productid = '" & TxtMerge.Text & "'"
   CN.Execute "delete from products where productid = '" & TxtMerge.Text & "'"
   MsgBox "Products Has Been Merged"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo ErrorHandler
   If KeyCode = vbKeyReturn Then
     keybd_event 9, 1, 1, 1
     KeyCode = 0
   End If
   If Shift = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyS
            If BtnSelect.Enabled Then BtnSelect_Click
            KeyCode = 0
         Case vbKeyQ
            If BtnClose.Enabled Then BtnClose_Click
            KeyCode = 0
      End Select
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    SetWindowText Me.hwnd, "Merge Products"
    Exit Sub
ErrorHandler:
  Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   Dim frmObj As Object
   For Each frmObj In Forms
       Set frmObj = Nothing
   Next
Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub
