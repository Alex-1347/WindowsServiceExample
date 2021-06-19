Imports System.Threading
Imports System.Timers
Imports Timer = System.Timers.Timer

Public Class PromocodeCacheCreater
    Private Timer As Timer
    Protected Overrides Async Sub OnStart(ByVal args() As String)
        Try
            Timer = New Timer(CInt(My.Settings.UpdateIntervalMinutes) * 60 * 1000) 'milliseconds
            AddHandler Timer.Elapsed, AddressOf TimerElapsed
            Timer.Start()
            Await BuildCache()
        Catch ex As Exception
            EventLog.WriteEntry("PromocodeCacheCreater", "Exception in OnStart " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace)
            My.Computer.FileSystem.WriteAllText(My.Settings.OutPathRoot & "Index.htm", "Exception in OnStart " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace, False)
            My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"Exception in OnStart & {IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace} {vbCrLf}", True)
        End Try
    End Sub
    Private Async Sub TimerElapsed(Sender As Object, E As ElapsedEventArgs)
        Await BuildCache()
    End Sub

    Protected Overrides Sub OnStop()
        Try
            If Timer IsNot Nothing AndAlso Timer.Enabled Then
                Timer.Stop()
            End If
        Catch Ex As Exception
            EventLog.WriteEntry("PromocodeCacheCreater", "Exception in OnStop " & IIf(String.IsNullOrWhiteSpace(Ex.Message), "Empty", Ex.Message) & Environment.NewLine & Ex.StackTrace)
            My.Computer.FileSystem.WriteAllText(My.Settings.OutPathRoot & "Index.htm", "Exception in OnStop " & IIf(String.IsNullOrWhiteSpace(Ex.Message), "Empty", Ex.Message) & Environment.NewLine & Ex.StackTrace, False)
            My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"Exception in OnStop & {IIf(String.IsNullOrWhiteSpace(Ex.Message), "Empty", Ex.Message) & Environment.NewLine & Ex.StackTrace} {vbCrLf}", True)
        End Try
    End Sub

    Public Shared TargetServerPath As String
    Public Shared TargetServerRoot As String
    Public Shared UrlList As List(Of String)
    Public Shared IsUrlCollected As Boolean = False

    Async Function BuildCache() As Task
        My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"{Now} Start {vbCrLf}", True)
        Dim Pos2 As Integer = My.Settings.URL.LastIndexOf("/")
        TargetServerPath = My.Settings.URL.Substring(0, Pos2)
        Dim TargetURI As Uri = New Uri(My.Settings.URL)
        TargetServerRoot = $"{TargetURI.Scheme}://{TargetURI.Host}"
        Try
            Dim HTML As String = Await ReadHtml(My.Settings.URL)
            If String.IsNullOrEmpty(HTML) Then Return
            UrlList = New List(Of String)
            IsUrlCollected = True
            Await Task.Run(Sub() ProcessingHTML(HTML))
            My.Computer.FileSystem.WriteAllText(My.Settings.OutPathRoot & "Index.htm", HTML, False)
            My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"{Now} cache created {My.Settings.OutPathRoot & "Index.htm"} ({HTML.Length} chars){vbCrLf}", True)
        Catch ex As Exception
            EventLog.WriteEntry("PromocodeCacheCreater", "Exception in BuildCache " & My.Settings.URL & " " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace)
            My.Computer.FileSystem.WriteAllText(My.Settings.OutPathRoot & "Index.htm", "Exception in BuildCache " & My.Settings.URL & " " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace, False)
            My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"Exception in BuildCache {My.Settings.URL} {IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace} {vbCrLf}", True)
        End Try
        '
        Dim IISvDir As New List(Of String)
        Dim PhysPath As New List(Of String)
        IsUrlCollected = False
        Dim J As Integer = 0
        For Each OneLink As String In UrlList
            If OneLink.StartsWith("/") And OneLink <> "/" Then
                ' ++ debug
                'If J < 10 Then
                '    J += 1
                'Else
                '    Exit For
                'End If
                ' ++
                Dim CurLink As String = (My.Settings.URL & OneLink).Replace("//", "/").Replace("https:/", "https://").Replace("http:/", "http://")
                Dim CurDir As String = (My.Settings.OutPathRoot & OneLink.Replace("/", "\")).Replace("\\", "\")
                IISvDir.Add(OneLink)
                PhysPath.Add(CurDir)
                Debug.Print(CurLink)
                Try
                    Dim HTML1 As String = Await ReadHtml(CurLink)
                    Await Task.Run(Sub() ProcessingHTML(HTML1))
                    If Not My.Computer.FileSystem.DirectoryExists(CurDir) Then
                        My.Computer.FileSystem.CreateDirectory(CurDir)
                    End If
                    My.Computer.FileSystem.WriteAllText(CurDir & "Index.htm", HTML1, False)
                    My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"{Now} cache created {CurDir & "Index.htm"} ({HTML1.Length} chars){vbCrLf}", True)
                    HTML1 = Nothing
                    Thread.Sleep(CInt(My.Settings.DelaySec) * 1000)
                Catch ex As Exception
                    EventLog.WriteEntry("PromocodeCacheCreater", "Exception in BuildCache " & CurLink & " " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace)
                    My.Computer.FileSystem.WriteAllText(CurDir & "Index.htm", "Exception in BuildCache " & CurLink & " " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace, False)
                    My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"Exception in BuildCache {CurLink} {IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace} {vbCrLf}", True)
                End Try
            End If
        Next
        '
        If Not CBool(My.Settings.AvoidCreateVdir) Then
            For i As Integer = 0 To IISvDir.Count - 1
                Dim Process As Process = New Process()
                Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
                startInfo.WindowStyle = ProcessWindowStyle.Hidden
                startInfo.FileName = "cmd.exe"
                'appcmd add vdir /app.name:"Default Web Site/" /path:"/Promokodi/1" /physicalPath:"F:\Promokodi\1"
                startInfo.Arguments = $"c:\Windows\System32\inetsrv\appcmd add vdir /app.name:""{My.Settings.IISAppName}"" /path:""/{(My.Settings.IISvDir & IISvDir(i)).Replace("//", "/")}"" /physicalPath:""{PhysPath(i)}"" "
                startInfo.Verb = "runas"
                Process.StartInfo = startInfo
                Try
                    Process.Start()
                Catch ex As Exception
                    EventLog.WriteEntry("PromocodeCacheCreater", "Exception in create Vdir '" & startInfo.Arguments & "' " & IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace)
                    My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"Exception in create Vdir '{startInfo.Arguments}' {IIf(String.IsNullOrWhiteSpace(ex.Message), "Empty", ex.Message) & Environment.NewLine & ex.StackTrace} {vbCrLf}", True)
                End Try
            Next

        End If
        My.Computer.FileSystem.WriteAllText(My.Settings.LogFile, $"{Now} Finish {vbCrLf}", True)
    End Function

End Class
