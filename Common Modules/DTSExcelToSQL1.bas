Attribute VB_Name = "DTSExcelToSQL"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\Documents and Settings\Sodtinn\My Documents\DTS.bas
'Package Name: New Package
'Package Description: DTS package description
'Generated Date: 2/10/2008
'Generated Time: 2:44:39 PM
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2

Public Sub DTSXLSToSQL(ExcelFileName As String, ServerName As String)
        Set goPackage = goPackageOld
        goPackage.Name = "New Package"
        goPackage.Description = "DTS package description"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")

        oConnection.ConnectionProperties("Data Source") = ExcelFileName
        oConnection.ConnectionProperties("Extended Properties") = "Excel 8.0;HDR=YES;"
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ExcelFileName
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = "sa"
        oConnection.ConnectionProperties("Initial Catalog") = "tempdb"
        oConnection.ConnectionProperties("Data Source") = ServerName
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ServerName
        oConnection.UserID = "sa"
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "tempdb"
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")

        oConnection.ConnectionProperties("Data Source") = ExcelFileName
        oConnection.ConnectionProperties("Extended Properties") = "Excel 8.0;HDR=YES;"
        
        oConnection.Name = "Connection 3"
        oConnection.ID = 3
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ExcelFileName
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = "sa"
        oConnection.ConnectionProperties("Initial Catalog") = "tempdb"
        oConnection.ConnectionProperties("Data Source") = ServerName
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 4"
        oConnection.ID = 4
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ServerName
        oConnection.UserID = "sa"
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "tempdb"
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [tempdb].[dbo].[Companies] Step"
        oStep.Description = "Create Table [tempdb].[dbo].[Companies] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [tempdb].[dbo].[Companies] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from Companies to [tempdb].[dbo].[Companies] Step"
        oStep.Description = "Copy Data from Companies to [tempdb].[dbo].[Companies] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from Companies to [tempdb].[dbo].[Companies] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = True
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [tempdb].[dbo].[Groups] Step"
        oStep.Description = "Create Table [tempdb].[dbo].[Groups] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [tempdb].[dbo].[Groups] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from Groups to [tempdb].[dbo].[Groups] Step"
        oStep.Description = "Copy Data from Groups to [tempdb].[dbo].[Groups] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from Groups to [tempdb].[dbo].[Groups] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = True
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [tempdb].[dbo].[Products] Step"
        oStep.Description = "Create Table [tempdb].[dbo].[Products] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [tempdb].[dbo].[Products] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from Products to [tempdb].[dbo].[Products] Step"
        oStep.Description = "Copy Data from Products to [tempdb].[dbo].[Products] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from Products to [tempdb].[dbo].[Products] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = True
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [tempdb].[dbo].[SubGroups] Step"
        oStep.Description = "Create Table [tempdb].[dbo].[SubGroups] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [tempdb].[dbo].[SubGroups] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Step"
        oStep.Description = "Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = True
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from Companies to [tempdb].[dbo].[Companies] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [tempdb].[dbo].[Companies] Step")
        oPrecConstraint.StepName = "Create Table [tempdb].[dbo].[Companies] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from Groups to [tempdb].[dbo].[Groups] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [tempdb].[dbo].[Groups] Step")
        oPrecConstraint.StepName = "Create Table [tempdb].[dbo].[Groups] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from Products to [tempdb].[dbo].[Products] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [tempdb].[dbo].[Products] Step")
        oPrecConstraint.StepName = "Create Table [tempdb].[dbo].[Products] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [tempdb].[dbo].[SubGroups] Step")
        oPrecConstraint.StepName = "Create Table [tempdb].[dbo].[SubGroups] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Create Table [tempdb].[dbo].[Companies] Task (Create Table [tempdb].[dbo].[Companies] Task)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from Companies to [tempdb].[dbo].[Companies] Task (Copy Data from Companies to [tempdb].[dbo].[Companies] Task)
Call Task_Sub2(goPackage)

'------------- call Task_Sub3 for task Create Table [tempdb].[dbo].[Groups] Task (Create Table [tempdb].[dbo].[Groups] Task)
Call Task_Sub3(goPackage)

'------------- call Task_Sub4 for task Copy Data from Groups to [tempdb].[dbo].[Groups] Task (Copy Data from Groups to [tempdb].[dbo].[Groups] Task)
Call Task_Sub4(goPackage)

'------------- call Task_Sub5 for task Create Table [tempdb].[dbo].[Products] Task (Create Table [tempdb].[dbo].[Products] Task)
Call Task_Sub5(goPackage)

'------------- call Task_Sub6 for task Copy Data from Products to [tempdb].[dbo].[Products] Task (Copy Data from Products to [tempdb].[dbo].[Products] Task)
Call Task_Sub6(goPackage)

'------------- call Task_Sub7 for task Create Table [tempdb].[dbo].[SubGroups] Task (Create Table [tempdb].[dbo].[SubGroups] Task)
Call Task_Sub7(goPackage)

'------------- call Task_Sub8 for task Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task (Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task)
Call Task_Sub8(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

End Sub


'------------- define Task_Sub1 for task Create Table [tempdb].[dbo].[Companies] Task (Create Table [tempdb].[dbo].[Companies] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Create Table [tempdb].[dbo].[Companies] Task"
        oCustomTask1.Description = "Create Table [tempdb].[dbo].[Companies] Task"
        oCustomTask1.SQLStatement = "CREATE TABLE [tempdb].[dbo].[Companies] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[CompanyID] float NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[CompanyName] nvarchar (255) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from Companies to [tempdb].[dbo].[Companies] Task (Copy Data from Companies to [tempdb].[dbo].[Companies] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from Companies to [tempdb].[dbo].[Companies] Task"
        oCustomTask2.Description = "Copy Data from Companies to [tempdb].[dbo].[Companies] Task"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceSQLStatement = "select `CompanyID`,`CompanyName` from `Companies`"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[tempdb].[dbo].[Companies]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0
        
Call oCustomTask2_Trans_Sub1(oCustomTask2)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("CompanyID", 1)
                        oColumn.Name = "CompanyID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("CompanyName", 2)
                        oColumn.Name = "CompanyName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("CompanyID", 1)
                        oColumn.Name = "CompanyID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("CompanyName", 2)
                        oColumn.Name = "CompanyName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub3 for task Create Table [tempdb].[dbo].[Groups] Task (Create Table [tempdb].[dbo].[Groups] Task)
Public Sub Task_Sub3(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask3 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask3 = oTask.CustomTask

        oCustomTask3.Name = "Create Table [tempdb].[dbo].[Groups] Task"
        oCustomTask3.Description = "Create Table [tempdb].[dbo].[Groups] Task"
        oCustomTask3.SQLStatement = "CREATE TABLE [tempdb].[dbo].[Groups] (" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[GroupID] nvarchar (255) NULL, " & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & "[GroupName] nvarchar (255) NULL" & vbCrLf
        oCustomTask3.SQLStatement = oCustomTask3.SQLStatement & ")"
        oCustomTask3.ConnectionID = 4
        oCustomTask3.CommandTimeout = 0
        oCustomTask3.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask3 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub4 for task Copy Data from Groups to [tempdb].[dbo].[Groups] Task (Copy Data from Groups to [tempdb].[dbo].[Groups] Task)
Public Sub Task_Sub4(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask4 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask4 = oTask.CustomTask

        oCustomTask4.Name = "Copy Data from Groups to [tempdb].[dbo].[Groups] Task"
        oCustomTask4.Description = "Copy Data from Groups to [tempdb].[dbo].[Groups] Task"
        oCustomTask4.SourceConnectionID = 3
        oCustomTask4.SourceSQLStatement = "select `GroupID`,`GroupName` from `Groups`"
        oCustomTask4.DestinationConnectionID = 4
        oCustomTask4.DestinationObjectName = "[tempdb].[dbo].[Groups]"
        oCustomTask4.ProgressRowCount = 1000
        oCustomTask4.MaximumErrorCount = 0
        oCustomTask4.FetchBufferSize = 1
        oCustomTask4.UseFastLoad = True
        oCustomTask4.InsertCommitSize = 0
        oCustomTask4.ExceptionFileColumnDelimiter = "|"
        oCustomTask4.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask4.AllowIdentityInserts = False
        oCustomTask4.FirstRow = 0
        oCustomTask4.LastRow = 0
        oCustomTask4.FastLoadOptions = 2
        oCustomTask4.ExceptionFileOptions = 1
        oCustomTask4.DataPumpOptions = 0
        
Call oCustomTask4_Trans_Sub1(oCustomTask4)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask4 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask4_Trans_Sub1(ByVal oCustomTask4 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask4.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("GroupID", 1)
                        oColumn.Name = "GroupID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("GroupName", 2)
                        oColumn.Name = "GroupName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("GroupID", 1)
                        oColumn.Name = "GroupID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("GroupName", 2)
                        oColumn.Name = "GroupName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask4.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub5 for task Create Table [tempdb].[dbo].[Products] Task (Create Table [tempdb].[dbo].[Products] Task)
Public Sub Task_Sub5(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask5 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask5 = oTask.CustomTask

        oCustomTask5.Name = "Create Table [tempdb].[dbo].[Products] Task"
        oCustomTask5.Description = "Create Table [tempdb].[dbo].[Products] Task"
        oCustomTask5.SQLStatement = "CREATE TABLE [tempdb].[dbo].[Products] (" & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[ProductID] nvarchar (255) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[CompanyID] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[GroupID] nvarchar (255) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[SubGroupID] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[ProductName] nvarchar (255) NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[PurPrice] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[RetailPrice] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[PurchasePackingID] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[DiscPer] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[DiscPC] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[StockLimit] float NULL, " & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & "[UnitID] float NULL" & vbCrLf
        oCustomTask5.SQLStatement = oCustomTask5.SQLStatement & ")"
        oCustomTask5.ConnectionID = 2
        oCustomTask5.CommandTimeout = 0
        oCustomTask5.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask5 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub6 for task Copy Data from Products to [tempdb].[dbo].[Products] Task (Copy Data from Products to [tempdb].[dbo].[Products] Task)
Public Sub Task_Sub6(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask6 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask6 = oTask.CustomTask

        oCustomTask6.Name = "Copy Data from Products to [tempdb].[dbo].[Products] Task"
        oCustomTask6.Description = "Copy Data from Products to [tempdb].[dbo].[Products] Task"
        oCustomTask6.SourceConnectionID = 1
        oCustomTask6.SourceSQLStatement = "select `ProductID`,`CompanyID`,`GroupID`,`SubGroupID`,`ProductName`,`PurPrice`,`RetailPrice`,`PurchasePackingID`,`DiscPer`,`DiscPC`,`StockLimit`,`UnitID` from `Products`"
        oCustomTask6.DestinationConnectionID = 2
        oCustomTask6.DestinationObjectName = "[tempdb].[dbo].[Products]"
        oCustomTask6.ProgressRowCount = 1000
        oCustomTask6.MaximumErrorCount = 0
        oCustomTask6.FetchBufferSize = 1
        oCustomTask6.UseFastLoad = True
        oCustomTask6.InsertCommitSize = 0
        oCustomTask6.ExceptionFileColumnDelimiter = "|"
        oCustomTask6.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask6.AllowIdentityInserts = False
        oCustomTask6.FirstRow = 0
        oCustomTask6.LastRow = 0
        oCustomTask6.FastLoadOptions = 2
        oCustomTask6.ExceptionFileOptions = 1
        oCustomTask6.DataPumpOptions = 0
        
Call oCustomTask6_Trans_Sub1(oCustomTask6)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask6 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask6_Trans_Sub1(ByVal oCustomTask6 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask6.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("ProductID", 1)
                        oColumn.Name = "ProductID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("CompanyID", 2)
                        oColumn.Name = "CompanyID"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("GroupID", 3)
                        oColumn.Name = "GroupID"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("SubGroupID", 4)
                        oColumn.Name = "SubGroupID"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("ProductName", 5)
                        oColumn.Name = "ProductName"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("PurPrice", 6)
                        oColumn.Name = "PurPrice"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("RetailPrice", 7)
                        oColumn.Name = "RetailPrice"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("PurchasePackingID", 8)
                        oColumn.Name = "PurchasePackingID"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("DiscPer", 9)
                        oColumn.Name = "DiscPer"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("DiscPC", 10)
                        oColumn.Name = "DiscPC"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("StockLimit", 11)
                        oColumn.Name = "StockLimit"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("UnitID", 12)
                        oColumn.Name = "UnitID"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ProductID", 1)
                        oColumn.Name = "ProductID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("CompanyID", 2)
                        oColumn.Name = "CompanyID"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("GroupID", 3)
                        oColumn.Name = "GroupID"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("SubGroupID", 4)
                        oColumn.Name = "SubGroupID"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ProductName", 5)
                        oColumn.Name = "ProductName"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("PurPrice", 6)
                        oColumn.Name = "PurPrice"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("RetailPrice", 7)
                        oColumn.Name = "RetailPrice"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("PurchasePackingID", 8)
                        oColumn.Name = "PurchasePackingID"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("DiscPer", 9)
                        oColumn.Name = "DiscPer"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("DiscPC", 10)
                        oColumn.Name = "DiscPC"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("StockLimit", 11)
                        oColumn.Name = "StockLimit"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("UnitID", 12)
                        oColumn.Name = "UnitID"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask6.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

'------------- define Task_Sub7 for task Create Table [tempdb].[dbo].[SubGroups] Task (Create Table [tempdb].[dbo].[SubGroups] Task)
Public Sub Task_Sub7(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask7 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask7 = oTask.CustomTask

        oCustomTask7.Name = "Create Table [tempdb].[dbo].[SubGroups] Task"
        oCustomTask7.Description = "Create Table [tempdb].[dbo].[SubGroups] Task"
        oCustomTask7.SQLStatement = "CREATE TABLE [tempdb].[dbo].[SubGroups] (" & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[SubGroupID] float NULL, " & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & "[SubGroupName] nvarchar (255) NULL" & vbCrLf
        oCustomTask7.SQLStatement = oCustomTask7.SQLStatement & ")"
        oCustomTask7.ConnectionID = 4
        oCustomTask7.CommandTimeout = 0
        oCustomTask7.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask7 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub8 for task Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task (Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task)
Public Sub Task_Sub8(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask8 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask8 = oTask.CustomTask

        oCustomTask8.Name = "Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task"
        oCustomTask8.Description = "Copy Data from SubGroups to [tempdb].[dbo].[SubGroups] Task"
        oCustomTask8.SourceConnectionID = 3
        oCustomTask8.SourceSQLStatement = "select `SubGroupID`,`SubGroupName` from `SubGroups`"
        oCustomTask8.DestinationConnectionID = 4
        oCustomTask8.DestinationObjectName = "[tempdb].[dbo].[SubGroups]"
        oCustomTask8.ProgressRowCount = 1000
        oCustomTask8.MaximumErrorCount = 0
        oCustomTask8.FetchBufferSize = 1
        oCustomTask8.UseFastLoad = True
        oCustomTask8.InsertCommitSize = 0
        oCustomTask8.ExceptionFileColumnDelimiter = "|"
        oCustomTask8.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask8.AllowIdentityInserts = False
        oCustomTask8.FirstRow = 0
        oCustomTask8.LastRow = 0
        oCustomTask8.FastLoadOptions = 2
        oCustomTask8.ExceptionFileOptions = 1
        oCustomTask8.DataPumpOptions = 0
        
Call oCustomTask8_Trans_Sub1(oCustomTask8)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask8 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask8_Trans_Sub1(ByVal oCustomTask8 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask8.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("SubGroupID", 1)
                        oColumn.Name = "SubGroupID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("SubGroupName", 2)
                        oColumn.Name = "SubGroupName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("SubGroupID", 1)
                        oColumn.Name = "SubGroupID"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 118
                        oColumn.Size = 0
                        oColumn.DataType = 5
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("SubGroupName", 2)
                        oColumn.Name = "SubGroupName"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 102
                        oColumn.Size = 255
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask8.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

