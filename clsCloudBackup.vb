'clsCloudBackup.vb
'	CloudBackup's Class Used to Synchronize Two WD myCloud Mirror Folders (Suitable for Scheduling)...
'   Copyright © 2006-2015, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   10/09/15    Ken Clark		Created;
'=================================================================================================================================
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.ComponentModel
Imports System.Collections.Specialized

Public Class clsCloudBackup
    Public Sub New()
        Dim at As Type = GetType(AssemblyProductAttribute)
        Dim r() As Object = [Assembly].GetEntryAssembly.GetCustomAttributes(at, False)
        Dim pt As AssemblyProductAttribute = CType(r(0), AssemblyProductAttribute)
        mApplicationName = pt.Product
    End Sub
#Region "Properties"
#Region "Win32 API"
    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As UInt32
    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As UInt32, ByVal fForce As UInt32) As UInt32
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure
    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1
#End Region
    Public Enum ByteEnum As Long
        KB = 1024
        MB = KB * 1024
        GB = MB * 1024
        TB = GB * 1024
    End Enum
    Private lockLogMessage As Object = New Object
    Private lockSendComplete As Object = New Object
    Private mApplicationName As String = ""
    Private mEmailError As String = ""
    Private mFilesCopied As Integer = 0
    Private mFilesSkipped As Integer = 0
    Private mLogFilePath As String = ""
    Private mLogMaxSizeMB As Integer = 10
    Private mMaxRetries As Integer = 100
    Private mSendComplete As Boolean = False
    Private mSourceDirectory As DirectoryInfo = Nothing
    Private mTargetDirectory As DirectoryInfo = Nothing
    Private mTestMode As Boolean = False
    Public ReadOnly Property ApplicationName As String
        Get
            Return mApplicationName
        End Get
    End Property
    Public ReadOnly Property LogFilePath As String
        Get
            Return mLogFilePath
        End Get
    End Property
    Public Property LogMaxSizeMB() As Integer
        Get
            Return mLogMaxSizeMB
        End Get
        Set(ByVal Value As Integer)
            If Value > 30 Then Throw New ArgumentException("LogMaxSizeMB cannot be larger than 30MB or file will not be able to be edited using available Windows tools.")
            mLogMaxSizeMB = Value
        End Set
    End Property
    Public ReadOnly Property FilesCopied As Integer
        Get
            Return mFilesCopied
        End Get
    End Property
    Public ReadOnly Property FilesSkipped As Integer
        Get
            Return mFilesSkipped
        End Get
    End Property
    Public ReadOnly Property SourceDirectory As DirectoryInfo
        Get
            Return mSourceDirectory
        End Get
    End Property
    Public ReadOnly Property TargetDirectory As DirectoryInfo
        Get
            Return mTargetDirectory
        End Get
    End Property
    Public ReadOnly Property TestMode As Boolean
        Get
            Return mTestMode
        End Get
    End Property
#End Region
    Private Sub Backup(ByVal SourceFolder As DirectoryInfo)
        LogMessage("<prefix>Scanning {0}...{1}", SourceFolder.FullName, vbCrLf)
        Dim diList As DirectoryInfo() = SourceFolder.GetDirectories()
        Array.Sort(Of DirectoryInfo)(diList, New DirectoryInfoComparer())
        For Each di As DirectoryInfo In diList
            Backup(di)
        Next
        Dim fiList As FileInfo() = SourceFolder.GetFiles()
        Array.Sort(Of FileInfo)(fiList, New FileInfoComparer())
        For Each fiSource As FileInfo In fiList
            BackupFile(SourceFolder, fiSource)
        Next
    End Sub
    Private Sub BackupFile(ByVal SourceFolder As DirectoryInfo, ByVal SourceFile As FileInfo)
        Dim TargetFolder As DirectoryInfo = Nothing,Targetfile As FileInfo = Nothing
        Dim tempSourceDrive As String = "", tempTargetDrive As String = ""
        Dim Retries As Integer = 0
        While True
            Try
                If tempSourceDrive = "" Then    'Otherwise we've been through the loop and have encountered a PathTooLongException
                    TargetFolder = New DirectoryInfo(SourceFolder.FullName.Replace(mSourceDirectory.FullName, mTargetDirectory.FullName))
                    Targetfile = New FileInfo(String.Format("{0}\{1}", TargetFolder.FullName, SourceFile.Name))
                End If
                If Not Targetfile.Exists OrElse SourceFile.LastWriteTimeUtc > Targetfile.LastWriteTimeUtc OrElse SourceFile.Length <> Targetfile.Length Then
                    If Not TargetFolder.Exists Then
                        LogMessage("<prefix>Creating {0}...", TargetFolder.FullName)
                        TargetFolder.Create()
                        LogMessage(" Done.{0}", vbCrLf)
                    End If
                    LogMessage("<prefix>Copying {0} to {1}...", SourceFile.FullName, TargetFolder.FullName)
                    If Not mTestMode Then SourceFile.CopyTo(Targetfile.FullName, True)
                    LogMessage(" Done.{0}", vbCrLf)
                    mFilesCopied += 1
                    If tempSourceDrive <> "" Then UnMapDrive(tempSourceDrive) : tempSourceDrive = ""
                    If tempTargetDrive <> "" Then UnMapDrive(tempTargetDrive) : tempTargetDrive = ""
                End If
                Exit While
            Catch ex As PathTooLongException
                Try
                    LogMessage("{0}<prefix>Handling PathTooLongException;{0}", vbCrLf)
                    tempSourceDrive = GetFreeDrive() : MapDrive(tempSourceDrive, SourceFolder.FullName)
                    LogMessage("<prefix>{0}Mapping {1}:\ as {2}{3}", New Object() {vbTab, tempSourceDrive, SourceFolder.FullName, vbCrLf})
                    SourceFile = New FileInfo(String.Format("{0}:\{1}", tempSourceDrive, SourceFile.Name))
                    tempTargetDrive = GetFreeDrive() : MapDrive(tempTargetDrive, TargetFolder.FullName)
                    LogMessage("<prefix>{0}Mapping {1}:\ as {2}{3}", New Object() {vbTab, tempTargetDrive, TargetFolder.FullName, vbCrLf})
                    Targetfile = New FileInfo(String.Format("{0}:\{1}", tempTargetDrive, SourceFile.Name))
                Catch ex2 As Exception
                    LogMessage("<prefix>{0}: {1}; File {2} not copied{3}", New Object() {CType(ex2, Object).GetType.Name, ex2.Message.Replace(vbCrLf, ""), SourceFile.FullName, vbCrLf})
                    mFilesSkipped += 1 : Exit While
                End Try
                'Catch ex As DirectoryNotFoundException : TargetFolder.Create()
            Catch ex As IOException When ex.Message.StartsWith("The process cannot access the file because another process has locked a portion of the file.")
                LogMessage("{0}<prefix>{1}: {2}; File {3} not copied{0}", New Object() {vbCrLf, CType(ex, Object).GetType.Name, ex.Message.Replace(vbCrLf, ""), SourceFile.FullName})
                mFilesSkipped += 1 : Exit While
            Catch ex As IOException When ex.Message.StartsWith("The specified network name is no longer available.")
                LogMessage("{0}<prefix>{1}: {2};{0}", vbCrLf, CType(ex, Object).GetType.Name, ex.Message.Replace(vbCrLf, ""))
                If Retries < mMaxRetries Then
                    Thread.Sleep(5000)  'Sleep 5 seconds before trying again...
                    Retries += 1
                Else
                    LogMessage("<prefix>{1}: {2}; Attempted copy {3:#,##0} times but problem persists; skipping file and moving on...{0}", New Object() {vbCrLf, CType(ex, Object).GetType.Name, ex.Message.Replace(vbCrLf, ""), mMaxRetries})
                    mFilesSkipped += 1 : Exit While
                End If
            Catch ex As Exception
                LogMessage("{0}<prefix>{1}: {2}; File {3} not copied{0}", New Object() {vbCrLf, CType(ex, Object).GetType.Name, ex.Message.Replace(vbCrLf, ""), SourceFile.FullName})
                mFilesSkipped += 1 : Exit While
            End Try
        End While
        If tempSourceDrive <> "" Then UnMapDrive(tempSourceDrive) : tempSourceDrive = ""
        If tempTargetDrive <> "" Then UnMapDrive(tempTargetDrive) : tempTargetDrive = ""
    End Sub
    Private Function FormatElapsed(ByVal milliseconds As Integer, Optional ByVal Format As Boolean = False) As String
        FormatElapsed = ""
        Dim ts As TimeSpan = New TimeSpan(0, 0, 0, 0, milliseconds)
        If Not Format Then
            If ts.Days > 0 Then FormatElapsed &= String.Format("{0} Days, ", ts.Days)
            If ts.Hours > 0 Then FormatElapsed &= String.Format("{0} Hours, ", ts.Hours)
            If ts.Minutes > 0 Then FormatElapsed &= String.Format("{0} Minutes, ", ts.Minutes)
            FormatElapsed &= String.Format("{0}.{1:000} Seconds", ts.Seconds, ts.Milliseconds)
        Else
            'By default TimeSpan.ToString displays 7 decimals, let's only go to 3...
            If ts.Days > 0 Then
                FormatElapsed = String.Format("{0}.{1:00}:{2:00}:{3:00}.{4:000}", New Object() {ts.Days, ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds})
            Else
                FormatElapsed = String.Format("{0:00}:{1:00}:{2:00}.{3:000}", New Object() {ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds})
            End If
        End If
    End Function
    Private Sub GetCommandLineArgs(ByVal args() As String)
        'Note: args(0) should be the path of the CloudBackup.exe...
        Select Case args.Length
            Case 0 : Throw New NotSupportedException("No command line arguments specified!")
            Case 3, 4, 5
                Dim argLogFile As Short = 3
                Dim argTest As Short = 4
                mSourceDirectory = New DirectoryInfo(args(1))
                mTargetDirectory = New DirectoryInfo(args(2))
                If Not mSourceDirectory.Exists Then Throw New ApplicationException(String.Format("SourceDirectory ({0}) does not exist or is inaccessible!", mSourceDirectory.FullName))
                If mSourceDirectory.Name.ToLower = "public" Then Throw New ApplicationException("Public Folder is too big to backup!")

                Select Case args.Length
                    Case 3  'Only Source and Target Directories provided
                        mLogFilePath = String.Format("{0}\{1}.{2}.log", mSourceDirectory.FullName, mApplicationName, mSourceDirectory.Name)
                    Case 4  'args(3) Could be LogFileName or TEST
                        If args(3).ToLower <> "test" Then
                            mLogFilePath = args(3)
                        Else
                            mLogFilePath = String.Format("{0}\{1}.{2}.log", mSourceDirectory.FullName, mApplicationName, mSourceDirectory.Name)
                            mTestMode = True
                        End If
                    Case 5  'args(4) Could be TEST or invalid
                        If args(4).ToLower <> "test" Then Throw New NotSupportedException(String.Format("{0} argument not supported!", args(4)))
                        mTestMode = True
                End Select
                Dim LogFileInfo As FileInfo = New FileInfo(mLogFilePath)
                If Not LogFileInfo.Directory.Exists Then Throw New ApplicationException(String.Format("Specified log file path is invalid - parent directory ({0}) does not exist or is inaccessible!", LogFileInfo.Directory.Name))
            Case Else : Throw New NotSupportedException(String.Format("Invalid number of command line arguments specified ({0})!", args.Length))
        End Select
    End Sub
    Private Function GetFreeDrive() As String
        Dim alphabet As New StringCollection()
        Dim lowerBound As Integer = Convert.ToInt16("b"c)
        Dim upperBound As Integer = Convert.ToInt16("z"c)
        For i As Integer = lowerBound To upperBound - 1
            Dim driveLetter As Char = ChrW(i)
            alphabet.Add(driveLetter.ToString())
        Next
        'Get all current drives
        Dim drives As DriveInfo() = DriveInfo.GetDrives()
        For Each drive As DriveInfo In drives
            alphabet.Remove(drive.Name.Substring(0, 1).ToLower())
        Next
        If alphabet.Count = 0 Then Throw New ApplicationException("No drives available.")
        Return alphabet(0).ToUpper
    End Function
    Public Function LogMessage(ByVal Format As String, ByVal Arg1 As Object, Optional ByVal Indent As Short = 0) As String
        Return LogMessage(String.Format(Format, Arg1), Indent)
    End Function
    Public Function LogMessage(ByVal Format As String, ByVal Arg1 As Object, ByVal Arg2 As Object, Optional ByVal Indent As Short = 0) As String
        Return LogMessage(String.Format(Format, Arg1, Arg2), Indent)
    End Function
    Public Function LogMessage(ByVal Format As String, ByVal Arg1 As Object, ByVal Arg2 As Object, ByVal Arg3 As Object, Optional ByVal Indent As Short = 0) As String
        Return LogMessage(String.Format(Format, Arg1, Arg2, Arg3), Indent)
    End Function
    Public Function LogMessage(ByVal Format As String, ByVal Args() As Object, Optional ByVal Indent As Short = 0) As String
        Return LogMessage(String.Format(Format, Args), Indent)
    End Function
    Public Function LogMessage(ByVal Message As String, Optional ByVal Indent As Short = 0) As String
        If Message.IndexOf("<prefix>") <> -1 Then
            Dim ThreadID As Integer = Threading.Thread.CurrentThread.ManagedThreadId
            Dim Prefix As String = String.Format("[{0:0000}] {1:MM/dd/yyyy HH:mm:ss.ffff}; {2}", ThreadID, Now, New String(vbTab, Indent))
            Message = Message.Replace("<prefix>", Prefix)
            'If Message.EndsWith(vbCrLf) Then Message = Message.Replace(vbCrLf, vbCrLf & Prefix).Replace(vbLf, vbCrLf).Replace(vbCr & vbCrLf, vbCrLf)
        End If

        SyncLock lockLogMessage
            'Write the message to the Immediate Window...
            Try : Debug.Write(Message) : Catch ex As Exception : End Try
            If mLogFilePath = "" Then Return Message
            'See if our file is too big to handle... If so, we'll rename it accordingly, and open a new one...
            If Message.EndsWith(vbCrLf) Then
                Dim LogFileInfo As FileInfo = Nothing
                Try
                    LogFileInfo = New FileInfo(mLogFilePath)
                    If LogFileInfo.Exists Then
                        If LogFileInfo.Length > (mLogMaxSizeMB * ByteEnum.MB) Then
                            Dim dtModified As Date = LogFileInfo.LastWriteTime
                            Dim NewFileName As String = String.Format("{0}.{1:yyyyMMddHHmmss}{2}", Path.GetFileNameWithoutExtension(LogFileInfo.Name), dtModified, LogFileInfo.Extension)
                            Rename(mLogFilePath, String.Format("{0}\{1}", LogFileInfo.DirectoryName, NewFileName))
                            'TODO: If we successfully renamed our existing file, now police any older files that need to be deleted...
                            'Dim LogDirInfo As New DirectoryInfo(LogFileInfo.DirectoryName)
                            'Dim LogFileList() As FileInfo = LogDirInfo.GetFiles(String.Format("{0}.*{1}", Path.GetFileNameWithoutExtension(LogFileInfo.Name), LogFileInfo.Extension))
                            'For Each iFileInfo As FileInfo In LogFileList
                            '    If DateDiff(DateInterval.DayOfYear, iFileInfo.LastWriteTime, Now) > mLogRetentionDays Then iFileInfo.Delete()
                            'Next
                        End If
                    End If
                Finally : LogFileInfo = Nothing
                End Try
            End If
            'Write the message to the file...
            Dim mLogFileWriter As New StreamWriter(mLogFilePath, True)
            mLogFileWriter.Write(Message) : mLogFileWriter.Flush() : mLogFileWriter.Close() : mLogFileWriter = Nothing
        End SyncLock
        Return Message
    End Function
    Public Function MapDrive(ByVal DriveLetter As String, ByVal UNCPath As String) As Boolean
        Dim nr As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String

        nr = New NETRESOURCE
        nr.lpRemoteName = UNCPath
        nr.lpLocalName = DriveLetter & ":"
        strUsername = Nothing '(add parameters to pass this if necessary)
        strPassword = Nothing '(add parameters to pass this if necessary)
        nr.dwType = RESOURCETYPE_DISK

        Dim result As Integer
        result = WNetAddConnection2(nr, strPassword, strUsername, 0)
        Return CBool(result = 0)
    End Function
    Public Sub Run(ByVal args() As String)
        GetCommandLineArgs(args)
        Dim Body As String = ""
        Dim Message As String = String.Format("[{0}{1}{2:HH:mm:ss}]{3}", New Object() {mApplicationName, vbTab, Now, vbCrLf})
        LogMessage("<prefix>{0}", Message) : Body &= Message
        Message = String.Format("{1} will be {2} {3}.{0}", New Object() {vbCrLf, mSourceDirectory.FullName, IIf(mTargetDirectory.Exists, "merged into", "copied to"), mTargetDirectory.FullName})
        LogMessage("<prefix>{0}", Message) : Body &= Message
        Message = String.Format("**** TEST MODE ****{0}", vbCrLf)
        If mTestMode Then LogMessage("<prefix>{0}", Message) : Body &= Message
        Dim sw As Stopwatch = New Stopwatch : sw.Start()
        Backup(mSourceDirectory)
        sw.Stop()
        Message = String.Format("{0}{1:#,##0} Files {2}Copied ({3:#,##0} skipped){4}", New Object() {vbTab, mFilesCopied, IIf(mTestMode, "[Test] ", ""), mFilesSkipped, vbCrLf})
        LogMessage("<prefix>{0}", Message) : Body &= Message
        Message = String.Format("{0} Complete @ {1:HH:mm:ss} (Elapsed: {2}){3}", New Object() {mApplicationName, Now, FormatElapsed(sw.ElapsedMilliseconds), vbCrLf})
        LogMessage("<prefix>{0}", Message) : Body &= Message
        LogMessage("<prefix>Sending E-mail...{0}", vbCrLf)
        SendMail(String.Format("{0} {1} {2}", mApplicationName, mSourceDirectory.FullName, IIf(mFilesSkipped = 0, "Succeeded", "Failed")), Body)
    End Sub
    Private Sub SendCompletedCallback(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
        mEmailError = ""
        SyncLock lockSendComplete : mSendComplete = False : End SyncLock
        Dim token As String = CStr(e.UserState)
        If e.Cancelled Then LogMessage("<prefix>Send canceled.{0}", vbCrLf)
        If e.Error IsNot Nothing Then mEmailError = e.Error.ToString()
        SyncLock lockSendComplete : mSendComplete = True : End SyncLock
    End Sub
    Public Sub SendMail(ByVal Subject As String, ByVal Body As String)
        Dim smtpClient As SmtpClient = New SmtpClient("smtp.comcast.net", 587)   '465
        smtpClient.Credentials = New NetworkCredential("kfc12", "cvn65BigE")
        smtpClient.EnableSsl = True
        'Email structure 
        Dim Email As MailMessage = New MailMessage("kfc12@comcast.net", "kfc12@comcast.net")
        'Email.Sender = "Ken Clark"
        Email.Subject = Subject
        Email.Body = Body
        'Email.Sender = New MailAddress("Ken Clark")
        Email.Attachments.Add(New Attachment(mLogFilePath))

        'smtpClient.Send(Email)
        AddHandler smtpClient.SendCompleted, AddressOf SendCompletedCallback
        smtpClient.SendAsync(Email, Subject)
        While True
            Thread.Sleep(1000)
            SyncLock lockSendComplete
                If mSendComplete Then Email.Dispose() : LogMessage("<prefix>E-mail sent.{0}", vbCrLf) : Exit While
            End SyncLock
        End While
    End Sub
    Public Function UnMapDrive(ByVal DriveLetter As String) As Boolean
        Dim rc As Integer
        rc = WNetCancelConnection2(DriveLetter & ":", 0, ForceDisconnect)
        Return CBool(rc = 0)
    End Function

    Public Overloads Shared Sub Main()
        System.Environment.ExitCode = Main(System.Environment.GetCommandLineArgs())
    End Sub
    Private Overloads Shared Function Main(ByVal args() As String) As Integer
        Dim cb As clsCloudBackup = Nothing
        Try
            cb = New clsCloudBackup() : cb.Run(args) : Return cb.FilesSkipped
        Catch ex As NotSupportedException
            Dim Message As String = String.Format("Usage:{0}{1} <SourceDirectory> <TargetDirectory>", vbCrLf, cb.ApplicationName)
            Debug.WriteLine(Message) : Console.WriteLine(Message)
            Return -1
        Catch ex As Exception
            If cb IsNot Nothing Then
                cb.LogMessage("<prefix>{1}{0}", vbCrLf, ex.ToString)
            Else
                Debug.WriteLine(ex.ToString) : Console.WriteLine(ex.ToString)
            End If
            Return -1
        End Try
    End Function
End Class
Public Class DirectoryInfoComparer
    Implements IComparer(Of DirectoryInfo)
    Public Function Compare(x As DirectoryInfo, y As DirectoryInfo) As Integer Implements IComparer(Of DirectoryInfo).Compare
        Return x.FullName.CompareTo(y.FullName)
    End Function
End Class
Public Class FileInfoComparer
    Implements IComparer(Of FileInfo)
    Public Function Compare(x As FileInfo, y As FileInfo) As Integer Implements IComparer(Of FileInfo).Compare
        Return x.FullName.CompareTo(y.FullName)
    End Function
End Class