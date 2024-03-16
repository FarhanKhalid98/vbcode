VERSION 5.00
Object = "{48CD706D-45EF-498E-A692-A47A77914F95}#1.0#0"; "JwldButn.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{437AC048-61C7-4A9E-B976-DF5D12A9A884}#17.0#0"; "SITextBox.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAutoBackupDefulatDB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15420
   Icon            =   "FrmAutoBackupDefulatDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAutoBackupDefulatDB.frx":0ECA
   ScaleHeight     =   794
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1028
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5322
      TabIndex        =   0
      Top             =   3923
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   5322
      TabIndex        =   1
      Top             =   4238
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   2430
      Left            =   8922
      TabIndex        =   2
      Top             =   3923
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3829
      Top             =   1598
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   375
      Left            =   7024
      TabIndex        =   7
      Top             =   8768
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   123142147
      CurrentDate     =   39225
   End
   Begin VB.CheckBox ChkSchedule 
      Height          =   195
      Left            =   9454
      TabIndex        =   12
      Top             =   10118
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSComCtl2.DTPicker DTPHHmm 
      Height          =   315
      Left            =   7024
      TabIndex        =   8
      Top             =   9173
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "mm"
      Format          =   123142147
      UpDown          =   -1  'True
      CurrentDate     =   39224.9826388889
   End
   Begin VB.ComboBox CmbHHmm 
      Height          =   315
      ItemData        =   "FrmAutoBackupDefulatDB.frx":783B
      Left            =   7834
      List            =   "FrmAutoBackupDefulatDB.frx":7845
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   9173
      Visible         =   0   'False
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker DTPStartingTime 
      Height          =   315
      Left            =   9994
      TabIndex        =   10
      Top             =   9173
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "HH:mm:ss "
      Format          =   123142147
      UpDown          =   -1  'True
      CurrentDate     =   39224.0416666667
   End
   Begin SITextBox.Txt TxtJobName 
      Height          =   315
      Left            =   5104
      TabIndex        =   5
      Top             =   2558
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      Appearance      =   0
      MaxLength       =   40
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
   Begin MSComCtl2.DTPicker DTPEndingTime 
      Height          =   315
      Left            =   9994
      TabIndex        =   11
      Top             =   9578
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "HH:mm:ss"
      Format          =   123142147
      UpDown          =   -1  'True
      CurrentDate     =   39224.9583333333
   End
   Begin JeweledBut.JeweledButton btnClose 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8232
      TabIndex        =   4
      Top             =   7478
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":785D
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnSave 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   8149
      TabIndex        =   13
      Top             =   7883
      Visible         =   0   'False
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":7879
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnOpen 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   5494
      TabIndex        =   15
      Top             =   7883
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Open"
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":7895
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btndelete 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   9529
      TabIndex        =   16
      Top             =   7883
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Remove"
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":78B1
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnBackup 
      Height          =   315
      Left            =   8089
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3053
      Visible         =   0   'False
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":78CD
      BC              =   14737632
      FC              =   0
   End
   Begin JeweledBut.JeweledButton btnClear 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6844
      TabIndex        =   14
      Top             =   7883
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   741
      TX              =   "Clear"
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":78E9
      BC              =   14737632
      FC              =   0
   End
   Begin SITextBox.Txt TxtBackupName 
      Height          =   315
      Left            =   5119
      TabIndex        =   6
      Top             =   3053
      Visible         =   0   'False
      Width           =   2985
      _ExtentX        =   5265
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
   Begin JeweledBut.JeweledButton BtnSaveDefault 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   7879
      TabIndex        =   3
      Top             =   6758
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      TX              =   "Create Auto Backup"
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
      MICON           =   "FrmAutoBackupDefulatDB.frx":7905
      BC              =   14737632
      FC              =   0
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atuo Backup Files"
      Height          =   195
      Left            =   8922
      TabIndex        =   30
      Top             =   3638
      Width           =   1290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atuo Backup Path"
      Height          =   195
      Left            =   5322
      TabIndex        =   29
      Top             =   3638
      Width           =   1305
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DataBase Backup Name:"
      Height          =   195
      Left            =   3004
      TabIndex        =   27
      Top             =   3143
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Job / Schedule"
      Height          =   195
      Left            =   9724
      TabIndex        =   26
      Top             =   10118
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AutoBackup Schedule"
      Height          =   195
      Left            =   5764
      TabIndex        =   25
      Top             =   8363
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Height          =   60
      Left            =   5764
      TabIndex        =   24
      Top             =   9983
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Height          =   60
      Left            =   5764
      TabIndex        =   23
      Top             =   8633
      Visible         =   0   'False
      Width           =   5550
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occure every:"
      Height          =   195
      Left            =   5854
      TabIndex        =   22
      Top             =   9218
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date:"
      Height          =   195
      Left            =   5854
      TabIndex        =   21
      Top             =   8858
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending at: "
      Height          =   195
      Left            =   9184
      TabIndex        =   20
      Top             =   9623
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting at:"
      Height          =   195
      Left            =   9184
      TabIndex        =   19
      Top             =   9218
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00DEAB97&
      BackStyle       =   0  'Transparent
      Caption         =   "AtuoBackup Job's Name:"
      Height          =   195
      Left            =   3004
      TabIndex        =   18
      Top             =   2603
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AutoBackup DataBase"
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
      TabIndex        =   17
      Top             =   270
      Width           =   3090
   End
   Begin VB.Image ImgExit 
      Height          =   360
      Left            =   11625
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmAutoBackupDefulatDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMode As FormMode
Dim sSql As String
Dim Jobid As String
Dim TypeHHmm As Integer
Dim IntervalHHmm As Integer
Dim DeviceName As String
Dim DatabaseName As String
Dim CompanyName As String
Dim BackupMessage As String
Dim vConnStr As String

Private Sub btnBackup_Click()
   CD1.FileName = ""
   CD1.DialogTitle = "Enter Path to take Auto Backup"
   CD1.InitDir = App.Path
   CD1.Filter = "(DataBase Backup)|*.bak"
   CD1.ShowSave
   If CD1.FileName <> "" Then
      TxtBackupName.Text = CD1.FileName
      DeviceName = Left(CD1.FileTitle, Len(CD1.FileTitle) - 4)
   Else
      CD1.FileName = ""
   End If
End Sub

Private Sub BtnClear_Click()
FormStatus = NewMode
End Sub

Private Sub BtnClose_Click()
Unload Me
End Sub

Private Sub BtnDelete_Click()
On Error GoTo ErrorHandler
    Open App.Path & "\Config.ini" For Input As #1
        Input #1, vConnStr
        Close #1
        DatabaseName = Left(vConnStr, InStr(1, vConnStr, ";") - 1)
    If MsgBox("Do you want to remove the Job?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
        CN.BeginTrans

        CN.DefaultDatabase = "master"
            DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
            DeviceName = Left(DeviceName, Len(DeviceName) - 4)
   
            CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
            + " Begin" & vbCrLf _
            + " exec sp_dropdevice " & DeviceName & vbCrLf _
            + " End")
   
    
   
   
   
        CN.DefaultDatabase = "msdb"
   
        With CN.Execute("Select Job_ID from SysJobs where name = '" & TxtJobName.Text & "'")
            If .EOF = True Then Exit Sub
            Jobid = !Job_ID
        End With
        CN.Execute ("Use msdb Exec sp_delete_job @job_name = '" & TxtJobName.Text & "'")
        CN.Execute ("Use msdb Exec sp_delete_jobstep @job_ID = '" & Jobid & "', @Step_ID = 1")
        CN.Execute ("Use msdb Exec sp_delete_jobschedule @job_ID = '" & Jobid & "', @name = 'Daily Backup'")
        CN.Execute ("Use msdb Exec sp_dropdevice @LogicalName = '" & DeviceName & "'")
        CN.CommitTrans
   
    CN.DefaultDatabase = DatabaseName
   
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
    
   Call ShowErrorMessage
   CN.DefaultDatabase = DatabaseName
End Sub

Private Sub BtnOpen_Click()
   SchAutoBackup.Show vbModal, Me
   If SchAutoBackup.ParaOutID <> "" Then
      TxtJobName.Text = SchAutoBackup.ParaOutID
      GetCompeleteInfo
      Else: Exit Sub
   End If
   TxtBackupName.SetFocus
End Sub

Private Sub BtnSave_Click()
On Error GoTo ErrorHandler

   
If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

If CmbHHmm.Text = "Hour(s)" Then
   IntervalHHmm = DTPHHmm.Hour
Else
   IntervalHHmm = DTPHHmm.Minute
End If
Dim vConnStr As String
Open App.Path & "\Config.ini" For Input As #1
Input #1, vConnStr
Close #1

DatabaseName = Left(vConnStr, InStr(1, vConnStr, ";") - 1)
DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
DeviceName = Left(DeviceName, Len(DeviceName) - 4)

' step 1  add device before check existence
CN.DefaultDatabase = "master"

CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")

CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


CN.Execute sSql

CN.DefaultDatabase = DatabaseName

''' step 1  add device before check existence
''CN.DefaultDatabase = "master"
''
''CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
''   + " Begin" & vbCrLf _
''   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
''   + " End")
''
''' step 2  add new job
''CN.DefaultDatabase = "msdb"
''CN.Execute ("If not Exists (SELECT  job_ID, name FROM  msdb.dbo.sysjobs WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
''   + " Begin" & vbCrLf _
''   + " EXEC sp_add_job @job_name = '" & TxtJobName.Text & "',@enabled = " & Abs(ChkSchedule.Value) & ",@description = '" & TxtBackupName.Text & "'" & vbCrLf _
''   + " End")
''
'''***************************Getting Job ID ***************************'
''With CN.Execute("Select Job_ID from SysJobs where name = '" & TxtJobName.Text & "'")
''   If .EOF = True Then Exit Sub
''      Jobid = !job_ID
''   End With
''
''' step 3  add new jobStep
''CN.Execute ("If not Exists (SELECT  job_ID FROM   msdb.dbo.sysjobsteps WHERE Job_ID = '" & Jobid & "')" & vbCrLf _
''   + " Begin" & vbCrLf _
''   + " Exec sp_add_jobstep  @job_ID = '" & Jobid & "', @step_id = 1, @step_name =  'Backup Processing ', @subsystem = 'TSQL'," & vbCrLf _
''   + " @Command = 'USE master" & vbCrLf _
''   + " BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT '" & vbCrLf _
''   + " End")
''
''
''' step 4  add new jobschedule
''CN.Execute ("If not Exists (SELECT  job_ID FROM   msdb.dbo.sysjobschedules WHERE Job_ID = '" & Jobid & "')" & vbCrLf _
''   + " Begin" & vbCrLf _
''   + " Exec sp_add_jobschedule  @job_ID = '" & Jobid & "', @name = 'Daily Backup'," & vbCrLf _
''   + " @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4, @freq_interval = 1," & vbCrLf _
''   + " @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
''   + " @active_start_date =" & Format(dtpStartDate, "yyyymmdd") & ", @active_End_date =" & Format(DateAdd("yyyy", 5, dtpStartDate.Value), "yyyymmdd") & "," & vbCrLf _
''   + " @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
''   + " End")

FormStatus = NewMode

 Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   CN.DefaultDatabase = DatabaseName
End Sub



Private Sub CmbHHmm_Click()
If CmbHHmm.Text = "Hour(s)" Then
   DTPHHmm.CustomFormat = "HH"
   TypeHHmm = 8
Else
   DTPHHmm.CustomFormat = "mm"
   TypeHHmm = 4
End If
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
      If KeyCode = vbKeyReturn Then
         keybd_event 9, 1, 1, 1
         KeyCode = 0
      ElseIf Shift = vbCtrlMask Then
         Select Case KeyCode
            Case vbKeyS
               If BtnSave.Enabled = True Then BtnSave_Click
               KeyCode = 0
            Case vbKeyW
               If BtnClear.Enabled = True Then BtnClear_Click
               KeyCode = 0
            Case vbKeyQ
               If BtnClose.Enabled = True Then BtnClose_Click
               KeyCode = 0
            Case vbKeyO
               If btnOpen.Enabled = True Then BtnOpen_Click
               KeyCode = 0
            Case vbKeyR
               If btndelete.Enabled = True Then BtnDelete_Click
               KeyCode = 0
            End Select
   End If
   Exit Sub
ErrorHandler:
     Call ShowErrorMessage
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Not (UCase(ActiveControl.Name) Like UCase("txt*")) Then Exit Sub
   If BtnSave.Enabled = False Then FormStatus = ChangeMode
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler
   SetWindowText Me.hWnd, "AutoBackup DataBase"
   ShowPicture Me, 2
   AddLabelEffect Me, 2, vbWhite, vbBlack, lblEffectBorder
   CmbHHmm.Text = "Minute(s)"
   DTPHHmm.CustomFormat = "mm"
   TypeHHmm = 4
   IntervalHHmm = DTPHHmm.Minute
   FormStatus = NewMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo ErrorHandler
   If BtnSave.Enabled = True Then
      If MsgBox("Do you want to close without save?", vbQuestion + vbYesNo + vbDefaultButton2, "Alert") = vbNo Then Cancel = True
   Else
      Dim frmObj As Object
      For Each frmObj In Forms
         Set frmObj = Nothing
      Next
         Set FrmAutoBackupDefulatDB = Nothing
   End If
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub ImgExit_Click()
   Unload Me
End Sub

Private Property Get FormStatus() As FormMode
  FormStatus = vMode
End Property

Private Property Let FormStatus(ByVal vNewValue As FormMode)
  On Error GoTo ErrorHandler
   Select Case vNewValue
   Case Is = NewMode
     Call SubClearFields
      btnOpen.Enabled = True
      btndelete.Enabled = False
      BtnSave.Enabled = False
      BtnClear.Enabled = True
      TxtJobName.Enabled = True
   Case Is = OpenMode
      btnOpen.Enabled = True
      btndelete.Enabled = True
      BtnClear.Enabled = True
      BtnSave.Enabled = False
   Case Is = ChangeMode
      btnOpen.Enabled = False
      btndelete.Enabled = False
      BtnSave.Enabled = True
   Case Is = SelectionMode
   End Select
   Exit Property
ErrorHandler:
   Call ShowErrorMessage
End Property

Private Sub SubClearFields()
   On Error GoTo ErrorHandler
   Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is TextBox Then
         ctl.Text = ""
      ElseIf TypeOf ctl Is SITextBox.txt Then
         ctl.Text = ""
      End If
   Next
   dtpStartDate.Value = Date
   DTPHHmm.Minute = 1
   DTPStartingTime = "00:00:00"
   DTPEndingTime = "23:59:00"
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub GetCompeleteInfo()
   On Error GoTo ErrorHandler
   sSql = "Select J.name, j.enabled,  j.Description, JSch.Freq_Subday_interval as time, JSch.Freq_subday_Type as MinuteOrHour, " & vbCrLf _
   + "JSch.active_Start_date as StartDate, JSch.active_start_time as StartTime, " & vbCrLf _
   + "JSch.active_end_time as EndTime from sysjobs J" & vbCrLf _
   + "Inner Join sysjobSchedules JSch on J.Job_ID = Jsch.Job_ID where j.name = '" & TxtJobName.Text & "'"
   With CN.Execute(sSql)
      If Not .EOF Then
         ChkSchedule.Value = !Enabled
         TxtBackupName.Text = !Description
         
         dtpStartDate.Year = Left(!startDate, 4)
         dtpStartDate.Month = Left(Right(!startDate, 4), 2)
         dtpStartDate.Day = Right(!startDate, 2)
         
         If !MinuteOrHour = 4 Then
            DTPHHmm.Minute = !Time
            CmbHHmm.Text = "Minute(s)"
         Else
            DTPHHmm.Hour = !Time
            CmbHHmm.Text = "Hour(s)"
         End If
         If Len(!StartTime) = 6 Then
            DTPStartingTime.Hour = Left(!StartTime, 2)
         Else
            DTPStartingTime.Hour = Left(!StartTime, 1)
         End If
         
         DTPStartingTime.Minute = Val(Left(Right(!StartTime, 4), 2))
         DTPStartingTime.Second = Val(Right(!StartTime, 2))
         
         If Len(!EndTime) = 6 Then
            DTPEndingTime.Hour = Left(!EndTime, 2)
         Else
            DTPEndingTime.Hour = Left(!EndTime, 1)
         End If
    
         DTPEndingTime.Minute = Val(Left(Right(!EndTime, 4), 2))
         DTPEndingTime.Second = Val(Right(!EndTime, 2))
         
      End If
      .Close
   End With
   
   FormStatus = OpenMode
   Exit Sub
ErrorHandler:
   Call ShowErrorMessage
End Sub

Private Sub BtnSaveDefault_Click()
    On Error GoTo ErrorHandler
    
    Me.MousePointer = vbHourglass
    
    BtnSaveDefault.Enabled = False
        
        dtpStartDate.Value = Date
        Open App.Path & "\Config.ini" For Input As #1
        Input #1, vConnStr
        Close #1
        DatabaseName = Left(vConnStr, InStr(1, vConnStr, ";") - 1)
        CN.DefaultDatabase = DatabaseName

       CompanyName = CN.Execute("select CompanyName from Company").Fields(0)

        CN.DefaultDatabase = "master"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''First Job Starts Daily at 9:00 AM''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "1st Backup Occurs Daily at 09:00 AM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_09_00_AM.bak"

    
    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 9
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 9
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''2nd Job Starts Daily at 12:00 PM'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "2nd Backup Occurs Daily at 12:00 PM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_12_00_PM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 12
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 12
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''3rd Job Starts Daily at 03:00 PM'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "3rd Backup Occurs Daily at 03:00 PM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_03_00_PM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 15
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 15
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''4th Job Starts Daily at 06:00 PM'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "4th Backup Occurs Daily at 06:00 PM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_06_00_PM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 18
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 18
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''5th Job Starts Daily at 09:00 PM'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "5th Backup Occurs Daily at 09:00 PM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_09_00_PM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 21
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 21
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''6th Job Starts Daily at 12:00 AM'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "6th Backup Occurs Daily at 12:00 AM"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_12_00_AM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 0
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 0
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''7th Job Starts at 09:30 AM and Repeate after each 01 Hour Till 08:30''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "7th Backup Occurs at 09:30 AM And Repeat After Each 1 Hour"
    TxtBackupName.Text = Dir1.Path + "\" & DatabaseName & "_0930AM_to_0830AM.bak"

    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 9
    DTPStartingTime.Minute = 30
    DTPEndingTime.Hour = 8
    DTPEndingTime.Minute = 30
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
   

    
    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  add device before check existence
    
    CN.DefaultDatabase = "master"

    CN.Execute ("If  Exists (Select * from sysdevices WHERE name = '" & DeviceName & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_dropdevice " & DeviceName & vbCrLf _
   + " End")
    
    CN.Execute ("If not Exists (Select * from sysdevices WHERE name = '" & TxtJobName.Text & "')" & vbCrLf _
   + " Begin" & vbCrLf _
   + " exec sp_addumpdevice 'disk', '" & DeviceName & "', '" & TxtBackupName.Text & "'" & vbCrLf _
   + " End")

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"
      
    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "
      
    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master BACKUP DATABASE " & DatabaseName & " TO " & DeviceName & " With INIT ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "
      
    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql
    
    
    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''' 8th New Backup Occurs Daily at 04:00 PM''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    TxtJobName.Text = "8th New Backup Occurs Daily at 04:00 PM"
    TxtBackupName.Text = Dir1.Path + "\"

    
    
    CmbHHmm.Text = "Hour(s)"
    DTPHHmm.Hour = 1
    DTPStartingTime.Hour = 16
    DTPStartingTime.Minute = 0
    DTPEndingTime.Hour = 16
    DTPEndingTime.Minute = 59
    
   
    
    If TxtJobName.Text = "" Or TxtBackupName.Text = "" Then Exit Sub

    If CmbHHmm.Text = "Hour(s)" Then
        IntervalHHmm = DTPHHmm.Hour
    Else
        IntervalHHmm = DTPHHmm.Minute
    End If
    
    
'    DeviceName = Mid(TxtBackupName.Text, InStrRev(TxtBackupName.Text, "\") + 1)
'    DeviceName = Left(DeviceName, Len(DeviceName) - 4)


' step 1  Create Procdure for Add Backup Device
    
    CN.DefaultDatabase = "master"
    
    sSql = "Drop Procedure Add_Device"

    CN.Execute ("If  Exists (Select * from SYSobjects where name = 'Add_Device' )" & vbCrLf _
   + " Begin" & vbCrLf _
   + sSql & vbCrLf _
   + " End")
    
    sSql = " Create Procedure Add_Device " & vbCrLf _
            + " as " & vbCrLf _
            + " Declare @DeviceName as varchar(40)" & vbCrLf _
            + " Declare @PathName as Varchar(100)" & vbCrLf _
            + " Select @DeviceName = Convert(Varchar(20),Getdate(),112) + ' ' + '" & CompanyName & "'" & vbCrLf _
            + " Set @PathName = '" + TxtBackupName.Text + "' +  @DeviceName + '.bak'" & vbCrLf _
            + " exec sp_addumpdevice 'disk', @DeviceName, @PathName"
   
    CN.Execute (sSql)
    
    ''' Create Backup
    sSql = "Drop Procedure Create_Backup"

    CN.Execute ("If  Exists (Select * from SYSobjects where name = 'Create_Backup' )" & vbCrLf _
   + " Begin" & vbCrLf _
   + sSql & vbCrLf _
   + " End")
   
     sSql = " Create Procedure Create_Backup " & vbCrLf _
            + " as " & vbCrLf _
            + " Declare @DeviceName as varchar(40)" & vbCrLf _
            + " Declare @PathName as Varchar(100)" & vbCrLf _
            + " Select @DeviceName = Convert(Varchar(20),Getdate(),112) + ' ' + '" & CompanyName & "'" & vbCrLf _
            + " BACKUP DATABASE " & DatabaseName & " to  @DeviceName WITH INIT"
   
    CN.Execute (sSql)
    
    ''' Drop Device
    sSql = "Drop Procedure Drop_Device"

    CN.Execute ("If  Exists (Select * from SYSobjects where name = 'Drop_Device' )" & vbCrLf _
   + " Begin" & vbCrLf _
   + sSql & vbCrLf _
   + " End")
   
     sSql = " Create Procedure Drop_Device " & vbCrLf _
            + " as " & vbCrLf _
            + " Declare @DeviceName as varchar(40)" & vbCrLf _
            + " Select @DeviceName = Convert(Varchar(20),Getdate(),112) + ' ' + '" & CompanyName & "'" & vbCrLf _
            + " exec sp_dropdevice @DeviceName"
   
    CN.Execute (sSql)

' step 2  add job
    sSql = " BEGIN TRANSACTION            " & vbCrLf _
      + "   DECLARE @JobID BINARY(16)  " & vbCrLf _
      + "   DECLARE @ReturnCode INT    " & vbCrLf _
      + "   SELECT @ReturnCode = 0     " & vbCrLf _
      + " IF (SELECT COUNT(*) FROM msdb.dbo.syscategories WHERE name = N'[Uncategorized (Local)]') < 1 " & vbCrLf _
      + "   EXECUTE msdb.dbo.sp_add_category @name = N'[Uncategorized (Local)]'"

    sSql = sSql + "   -- Delete the job with the same name (if it exists)" & vbCrLf _
      + " SELECT @JobID = job_id     " & vbCrLf _
      + "   FROM   msdb.dbo.sysjobs    " & vbCrLf _
      + "   WHERE (name = N'" & TxtJobName.Text & "')       " & vbCrLf _
      + "   IF (@JobID IS NOT NULL)    " & vbCrLf _
      + "   BEGIN  " & vbCrLf _
      + "   -- Check if the job is a multi-server job  " & vbCrLf _
      + "   IF (EXISTS (SELECT  * " & vbCrLf _
      + "               FROM    msdb.dbo.sysjobservers " & vbCrLf _
      + "               WHERE   (job_id = @JobID) AND (server_id <> 0))) " & vbCrLf _
      + "   BEGIN " & vbCrLf _
      + "     -- There is, so abort the script " & vbCrLf _
      + "     RAISERROR (N'Unable to import job ''" & TxtJobName.Text & "'' since there is already a multi-server job with this name.', 16, 1) " & vbCrLf _
      + "     GOTO QuitWithRollback  " & vbCrLf _
      + "   END " & vbCrLf _
      + "   ELSE " & vbCrLf _
      + "     -- Delete the [local] job " & vbCrLf _
      + "     EXECUTE msdb.dbo.sp_delete_job @job_name = N'" & TxtJobName.Text & "' " & vbCrLf _
      + "     SELECT @JobID = NULL" & vbCrLf _
      + "   END " & vbCrLf _
      + " BEGIN "

    sSql = sSql + "  -- Add the job " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id = @JobID OUTPUT , @job_name = N'" & TxtJobName.Text & "', @owner_login_name = N'sa'," & vbCrLf _
      + " @description = N'" & TxtBackupName.Text & "', @category_name = N'[Uncategorized (Local)]', @enabled = " & Abs(ChkSchedule.Value) & "," & vbCrLf _
      + " @notify_level_email = 0, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 2, @delete_level= 0" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the job steps" & vbCrLf _
      + "  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id = @JobID, @step_id = 1, @step_name = N'Backup Processing', " & vbCrLf _
      + " @command = N'USE master exec Add_Device exec Create_Backup exec Drop_Device ', @database_name = N'master', @server = N'', @database_user_name = N'', " & vbCrLf _
      + " @subsystem = N'TSQL', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'', " & vbCrLf _
      + " @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2" & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id = @JobID, @start_step_id = 1 " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback "

    sSql = sSql + "   -- Add the job schedules " & vbCrLf _
      + "   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id = @JobID, @name = N'" & TxtBackupName.Text & "', @enabled = " & Abs(ChkSchedule.Value) & ", @freq_type = 4," & vbCrLf _
      + " @active_start_date = " & Format(dtpStartDate, "yyyymmdd") & ", @active_start_time = " & Format(DTPStartingTime, "HHmmss") & ", @freq_interval = 1, @freq_subday_type = " & TypeHHmm & ", @freq_subday_interval = " & IntervalHHmm & "," & vbCrLf _
      + " @freq_relative_interval = 0, @freq_recurrence_factor = 0, @active_end_date = " & Format(DateAdd("yyyy", 10, dtpStartDate.Value), "yyyymmdd") & ", @active_end_time = " & Format(DTPEndingTime, "HHmmss") & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + "   -- Add the Target Servers " & vbCrLf _
      + " EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @JobID, @server_name = N'(local)' " & vbCrLf _
      + "   IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback " & vbCrLf _
      + " END" & vbCrLf _
      + " COMMIT TRANSACTION          " & vbCrLf _
      + " GOTO   EndSave              " & vbCrLf _
      + " QuitWithRollback:" & vbCrLf _
      + "   IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION " & vbCrLf _
      + " EndSave: "


    CN.Execute sSql

    BackupMessage = BackupMessage & vbCrLf & TxtJobName.Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    
Shell ("net start SQLSERVERAGENT")
Shell ("SCM -Action 7 -Service sqlServeragent -SvcStartType 2")
MsgBox "Seven Different Backups Created :-" & vbCrLf & vbCrLf & BackupMessage, vbInformation
CN.DefaultDatabase = DatabaseName
Me.MousePointer = vbDefault
BtnSaveDefault.Enabled = True

Exit Sub
ErrorHandler:
   Call ShowErrorMessage
   CN.DefaultDatabase = DatabaseName
   BtnSaveDefault.Enabled = True
   Me.MousePointer = vbDefault

End Sub
