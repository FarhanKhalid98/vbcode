VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip SSTAB 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Control As TextBox
Attribute Control.VB_VarHelpID = -1

Dim WithEvents Ch_Delete_Row As CheckBox
Attribute Ch_Delete_Row.VB_VarHelpID = -1
Dim WithEvents Ch_SR_NO As Label
Attribute Ch_SR_NO.VB_VarHelpID = -1
Dim WithEvents Ch_Name As TextBox
Attribute Ch_Name.VB_VarHelpID = -1
Dim WithEvents Ch_Type As ComboBox
Attribute Ch_Type.VB_VarHelpID = -1

Dim WithEvents Extended_Control As VBControlExtender
Attribute Extended_Control.VB_VarHelpID = -1


Private Sub cmdAddCharacteristics_Click()

    Module1.SR_NO = Module1.SR_NO + 1
    Set Ch_Delete_Row = frmCharacteristics.Controls.Add("VB.CheckBox", "Ch_Delete_Row" & (Module1.SR_NO), tabDisplay)
    Ch_Delete_Row.Visible = True
    Ch_Delete_Row.Top = Module1.Top_Position + 100
    Ch_Delete_Row.Width = 1000
    Ch_Delete_Row.Left = 500
    Ch_Delete_Row.Caption = ""
    Ch_Delete_Row.Height = 315
    'MsgBox Ch_Delete_Row.Name

    Set Ch_SR_NO = frmCharacteristics.Controls.Add("VB.Label", "Ch_SR_NO" & (Module1.SR_NO), tabDisplay)
    Ch_SR_NO.Visible = True
    Ch_SR_NO.Top = Module1.Top_Position + 200
    Ch_SR_NO.Width = 750
    Ch_SR_NO.Left = Ch_Delete_Row.Left + Ch_Delete_Row.Width + 400
    Ch_SR_NO.Caption = Module1.SR_NO
    Ch_SR_NO.Height = 315

    Set Ch_Name = frmCharacteristics.Controls.Add("VB.TextBox", "Ch_Name" & (Module1.SR_NO), tabDisplay)
    Ch_Name.Visible = True
    Ch_Name.Top = Module1.Top_Position + 100
    Ch_Name.Width = 2000
    Ch_Name.Left = Ch_SR_NO.Left + Ch_SR_NO.Width + 200
    Ch_Name.Text = ""
    Ch_Name.Height = 315

    Set Ch_Type = frmCharacteristics.Controls.Add("VB.ComboBox", "Ch_Type" & (Module1.SR_NO), tabDisplay)
    Ch_Type.Visible = True
    Ch_Type.Top = Module1.Top_Position + 100
    Ch_Type.Width = 1500
    Ch_Type.Left = Ch_Name.Left + Ch_Name.Width + 50
    Ch_Type.Text = ""
    'Ch_Type.Height = 315
    Ch_Type.AddItem "Service"
    Ch_Type.AddItem "Special"
    Ch_Type.AddItem "Option"

    Module1.Top_Position = Module1.Top_Position + 400
End Sub

Private Sub Form_Load()
    Module1.SR_NO = 0
    Dim Test_Line As Control
    Set Test_Line = frmCharacteristics.Controls.Add("VB.Line", "LINE", frmCharacteristics)
    Test_Line.Visible = True
    Test_Line.X1 = 100
    Test_Line.Y1 = 600
    Test_Line.X2 = frmCharacteristics.Width
    Test_Line.Y2 = 600
    Top_Position = Test_Line.Y1
    frmCharacteristics.Show
    tabDisplay.Width = frmCharacteristics.Width - 1000
    tabDisplay.Height = frmCharacteristics.Height - 1500
    tabDisplay.Left = frmCharacteristics.Left + 200
    Call set_labels
End Sub


Sub set_labels()

    Dim Label_SR_NO As Control
    Dim Label_Name As Control
    Dim Label_Delete_Row As Control
    Dim Label_Type As Control

    Set Label_Delete_Row = frmCharacteristics.Controls.Add("VB.Label", "Label_Delete_Row" & (Module1.SR_NO), tabDisplay)
    Label_Delete_Row.Visible = True
    Label_Delete_Row.Top = Module1.Top_Position + 100
    Label_Delete_Row.Width = 1000
    Label_Delete_Row.Left = 300
    Label_Delete_Row.Caption = "Delete(Y/N)"
    Label_Delete_Row.Height = 315

    Set Label_SR_NO = frmCharacteristics.Controls.Add("VB.Label", "Label_SR_NO" & (Module1.SR_NO), tabDisplay)
    Label_SR_NO.Visible = True
    Label_SR_NO.Top = Module1.Top_Position + 100
    Label_SR_NO.Width = 750
    Label_SR_NO.Left = Label_Delete_Row.Left + Label_Delete_Row.Width + 400
    Label_SR_NO.Caption = "SR_NO"
    Label_SR_NO.Height = 315

    Set Label_Name = frmCharacteristics.Controls.Add("VB.Label", "Label_Name" & (Module1.SR_NO), tabDisplay)
    Label_Name.Visible = True
    Label_Name.Top = Module1.Top_Position + 100
    Label_Name.Width = 2000
    Label_Name.Left = Label_SR_NO.Left + Label_SR_NO.Width + 400
    Label_Name.Caption = "Characteristics Name"
    Label_Name.Height = 315

    Set Label_Type = frmCharacteristics.Controls.Add("VB.Label", "Label_Type" & (Module1.SR_NO), tabDisplay)
    Label_Type.Visible = True
    Label_Type.Top = Module1.Top_Position + 100
    Label_Type.Width = 1500
    Label_Type.Left = Label_Name.Left + Label_Name.Width + 50
    Label_Type.Caption = "Charac. Type"
    Label_Type.Height = 315

    Module1.Top_Position = Module1.Top_Position + 400
End Sub

Private Sub Control_GotFocus()
    Control.SelStart = 0
    Control.SelLength = Len(Control.Text)
End Sub

Private Sub Control_LostFocus()
    Control.SelLength = 0
End Sub
Public Sub InitialiseTextBoxExtenders(ByRef myForm As Form, ByRef extenderCollection As Collection)
    Dim formControl As Control
    Dim oTBXExtender As clsTextBoxExtender
    For Each formControl In myForm.Controls
        If TypeOf formControl Is TextBox Then
            Set oTBXExtender = New clsTextBoxExtender
            Set oTBXExtender.Control = formControl
            extenderCollection.Add oTBXExtender
        End If
     Next
End Sub
Private textBoxExtenderCollection As New Collection

Private Sub Form1_Load()
    Module1.InitialiseTextBoxExtenders Me, textBoxExtenderCollection
End Sub


