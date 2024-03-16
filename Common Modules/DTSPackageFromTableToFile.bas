Attribute VB_Name = "DTSPackageFromTableToFile"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\Documents and Settings\Farhan\My Documents\New Package1.bas
'Package Name: New Package
'Package Description: DTS package description
'Generated Date: 10/1/2007
'Generated Time: 3:38:11 PM
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2
Private FileName As String

Public Sub DTSPackage(File As String, ServerName As String, DatabaseName As String)
        FileName = File
        Set goPackage = goPackageOld
        goPackage.Name = "Activity Package"
        goPackage.Description = "DTS Package Description"
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

Set oConnection = goPackage.Connections.New("SQLOLEDB")

'        oConnection.ConnectionProperties("Integrated Security") = "SSPI"
'        oConnection.ConnectionProperties("Persist Security Info") = True
'        oConnection.ConnectionProperties("Initial Catalog") = "Awan"
'        oConnection.ConnectionProperties("Data Source") = "(LOCAL)"

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = "sa"
        oConnection.ConnectionProperties("Initial Catalog") = DatabaseName
        oConnection.ConnectionProperties("Data Source") = ServerName
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ServerName
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = DatabaseName
        oConnection.UseTrustedConnection = True
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = "C:\WINDOWS\Log\" & FileName & ".txt"
        oConnection.ConnectionProperties("Mode") = 3
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 2
        oConnection.ConnectionProperties("Column Lengths") = "5,25,17,200,5,5,5"
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("First Row Column Name") = True
        oConnection.ConnectionProperties("Column Names") = "UserNo,FormType,EntryDate,Description,isNew,isEdit,isDelete"
        oConnection.ConnectionProperties("Number of Column") = 7
        oConnection.ConnectionProperties("Text Qualifier Col Mask: 0=no, 1=yes, e.g. 0101") = "0101000"
        oConnection.ConnectionProperties("Blob Col Mask: 0=no, 1=yes, e.g. 0101") = "0000000"
                
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "C:\WINDOWS\Log\" & FileName & ".txt"
        oConnection.ConnectionTimeout = 60
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

        oStep.Name = "Copy Data from Activity to C:\WINDOWS\Log\" & FileName & ".txt Step"
        oStep.Description = "Copy Data from Activity to C:\WINDOWS\Log\" & FileName & ".txt Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copied data in table C:\WINDOWS\Log\" & FileName & ".txt"
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

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Copied data in table C:\WINDOWS\abc (Copied data in table C:\WINDOWS\abc)
Call Task_Sub1(goPackage)

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


'------------- define Task_Sub1 for task Copied data in table C:\WINDOWS\abc (Copied data in table C:\WINDOWS\abc)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Copied data in table C:\WINDOWS\Log\" & FileName & ".txt"
        oCustomTask1.Description = "Copied data in table C:\WINDOWS\Log\" & FileName & ".txt"
        oCustomTask1.SourceConnectionID = 1
        oCustomTask1.SourceSQLStatement = "select [UserNo],[FormType],[EntryDate],[Description],[isNew],[isEdit],[isDelete] from [dbo].[ActivityLog]"
        oCustomTask1.DestinationConnectionID = 2
        oCustomTask1.DestinationObjectName = "C:\WINDOWS\Log\" & FileName & ".txt"
        oCustomTask1.ProgressRowCount = 1000
        oCustomTask1.MaximumErrorCount = 0
        oCustomTask1.FetchBufferSize = 1
        oCustomTask1.UseFastLoad = True
        oCustomTask1.InsertCommitSize = 0
        oCustomTask1.ExceptionFileColumnDelimiter = "|"
        oCustomTask1.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask1.AllowIdentityInserts = False
        oCustomTask1.FirstRow = 0
        oCustomTask1.LastRow = 0
        oCustomTask1.FastLoadOptions = 2
        oCustomTask1.ExceptionFileOptions = 1
        oCustomTask1.DataPumpOptions = 0
        
Call oCustomTask1_Trans_Sub1(oCustomTask1)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask1_Trans_Sub1(ByVal oCustomTask1 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask1.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("UserNo", 1)
                        oColumn.Name = "UserNo"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 24
                        oColumn.Size = 0
                        oColumn.DataType = 17
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("FormType", 2)
                        oColumn.Name = "FormType"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 8
                        oColumn.Size = 25
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("EntryDate", 3)
                        oColumn.Name = "EntryDate"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 24
                        oColumn.Size = 0
                        oColumn.DataType = 135
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Description", 4)
                        oColumn.Name = "Description"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 8
                        oColumn.Size = 200
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("isNew", 5)
                        oColumn.Name = "isNew"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 24
                        oColumn.Size = 0
                        oColumn.DataType = 11
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("isEdit", 6)
                        oColumn.Name = "isEdit"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 24
                        oColumn.Size = 0
                        oColumn.DataType = 11
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("isDelete", 7)
                        oColumn.Name = "isDelete"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 24
                        oColumn.Size = 0
                        oColumn.DataType = 11
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("UserNo", 1)
                        oColumn.Name = "UserNo"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 24
                        oColumn.Size = 5
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("FormType", 2)
                        oColumn.Name = "FormType"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 8
                        oColumn.Size = 25
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("EntryDate", 3)
                        oColumn.Name = "EntryDate"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 24
                        oColumn.Size = 17
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Description", 4)
                        oColumn.Name = "Description"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 8
                        oColumn.Size = 200
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("isNew", 5)
                        oColumn.Name = "isNew"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 24
                        oColumn.Size = 5
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("isEdit", 6)
                        oColumn.Name = "isEdit"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 24
                        oColumn.Size = 5
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("isDelete", 7)
                        oColumn.Name = "isDelete"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 24
                        oColumn.Size = 5
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = False
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask1.Transformations.Add oTransformation
        Set oTransformation = Nothing
        
End Sub
