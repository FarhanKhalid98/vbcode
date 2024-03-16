VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Begin VB.Form FrmConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating new Configuration for Database"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtServer 
      Height          =   315
      Left            =   735
      TabIndex        =   4
      Text            =   "(Local)"
      Top             =   2535
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What kind of machine is this?"
      Height          =   645
      Left            =   705
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
      Begin VB.OptionButton OptServer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server"
         Height          =   240
         Left            =   2505
         TabIndex        =   3
         Top             =   300
         Width           =   1185
      End
      Begin VB.OptionButton OptClient 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client"
         Height          =   240
         Left            =   1005
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin JeweledBut.JeweledButton BtnSelect 
      Default         =   -1  'True
      Height          =   420
      Left            =   3120
      TabIndex        =   6
      Top             =   3510
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "OK"
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
      MICON           =   "FrmConfig.frx":0000
      BC              =   12632256
      FC              =   0
   End
   Begin JeweledBut.JeweledButton BtnClose 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   4440
      TabIndex        =   7
      Top             =   3510
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Cancel"
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
      MICON           =   "FrmConfig.frx":001C
      BC              =   12632256
      FC              =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For more information, contact the program vendor"
      Height          =   300
      Left            =   765
      TabIndex        =   8
      Top             =   3090
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name:"
      Height          =   210
      Left            =   735
      TabIndex        =   5
      Top             =   2325
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmConfig.frx":0038
      Height          =   660
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClose_Click()
    End
End Sub

Private Sub BtnSelect_Click()
    On Error GoTo ErrorHandler
    Dim objFSO As New Scripting.FileSystemObject
    Dim objFile As File
    Screen.MousePointer = vbHourglass
    If OptServer.Value And Trim(TxtServer.Text) = "" Then
        MsgBox "Please specify a valid server name.", vbExclamation, "Alert"
        Exit Sub
    End If
    Open App.Path & "\config.ini" For Output As #1
    Print #1, "SuperSoftv1;Data Source=" & IIf(OptServer.Value, "(Local)", TxtServer.Text)
    Close #1
    If CN.State = adStateOpen Then CN.Close
    'CN.Open "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=master;Data Source=" & TxtServer.Text
    CN.Open "Driver=SQL Server;Server=" & TxtServer.Text & ";database=master"
    CN.CursorLocation = adUseClient
    If Not DBExists Then
'        Dim vCommands() As String, vTotalText As String, vTempLine
'        Dim i As Integer
'        '---------------- creating database,tables,triggers,views,storedprocedures---------
'        Open App.Path & "\DBScript.script" For Input As #1
'        vTotalText = ""
'        While Not EOF(1)
'            Line Input #1, vTempLine
'            If Trim(vTempLine) <> "" Then vTotalText = vTotalText & vTempLine & vbCrLf
'        Wend
'        vCommands = Split(vTotalText, "GO")
'        For i = 0 To UBound(vCommands)
'            CN.Execute (Replace(Replace(vCommands(i), Chr(255), ""), Chr(254), ""))
'        Next
'        'MsgBox i & " statements executed"
'        Close #1
'        '--------------inserting defualt accounts in chartofaccounts--------------
'        Open App.Path & "\AccChart.script" For Input As #1
'        vTotalText = ""
'        While Not EOF(1)
'            Line Input #1, vTempLine
'            If Trim(vTempLine) <> "" Then vTotalText = vTotalText & vTempLine & vbCrLf
'        Wend
'        vCommands = Split(vTotalText, "GO")
'        For i = 0 To UBound(vCommands)
'            CN.Execute (Replace(Replace(vCommands(i), Chr(255), ""), Chr(254), ""))
'        Next
'        'CN.Execute vTotalText
'        Close #1
'        '-------------inserting registry values,tasks------------------
'        Open App.Path & "\Registry.script" For Input As #1
'        vTotalText = ""
'        While Not EOF(1)
'            Line Input #1, vTempLine
'            If Trim(vTempLine) <> "" Then vTotalText = vTotalText & vTempLine & vbCrLf
'        Wend
'        vCommands = Split(vTotalText, "GO")
'        For i = 0 To UBound(vCommands)
'            CN.Execute (Replace(Replace(vCommands(i), Chr(255), ""), Chr(254), ""))
'        Next
'        'CN.Execute vTotalText
'        Close #1
   'if dir("d:\database\Department_Data.mdf") <>"" then
   If objFSO.FileExists("d:\database\SuperSoftv1_Data.mdf") Then
      CN.Execute "EXEC sp_attach_db 'SuperSoftv1','d:\database\SuperSoftv1_Data.MDF','d:\database\SuperSoftv1_Log.LDF'"
   ElseIf objFSO.FileExists(App.Path & "\database\SuperSoftv1_Data.mdf") Then
      CN.Execute "EXEC sp_attach_db 'SuperSoftv1','" & App.Path & "\DataBase\SuperSoftv1_Data.MDF','" & App.Path & "\DataBase\SuperSoftv1_Log.LDF'"
   Else
      MsgBox "Database Not Found", vbInformation + vbOKOnly, "Error"
   End If
    'CN.Execute "exec sp_addlogin 'Tra', '' exec sp_grantlogin 'Tra' EXEC sp_addsrvrolemember 'Tra', 'sysadmin' EXEC sp_addsrvrolemember 'Tra', 'securityadmin' EXEC sp_addsrvrolemember 'Tra', 'serveradmin' EXEC sp_addsrvrolemember 'Tra', 'setupadmin' EXEC sp_addsrvrolemember 'Tra', 'processadmin' EXEC sp_addsrvrolemember 'Tra', 'diskadmin' EXEC sp_addsrvrolemember 'Tra', 'dbcreator' EXEC sp_addsrvrolemember 'Tra', 'bulkadmin'"
    End If
    Screen.MousePointer = vbDefault
    
    MsgBox "The configuration was completed successfully", vbInformation
    Unload Me
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    Call ShowErrorMessage
End Sub
