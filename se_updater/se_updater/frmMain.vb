Imports System.IO

Public Class frmMain
    Const LogLine As String = "{0} - {1}"

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If BackgroundWorker1.IsBusy Then
            BackgroundWorker1.CancelAsync()
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblProgress.Text = "Verifying that current version of Solid Edge is installed..."
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Check if Solid Edge exists on this PC.
        Dim InstallData As New SEInstallDataLib.SEInstallData
        If InstallData.GetMajorVersion = CInt(My.Settings.SEVersion) Then
            'Load log file writer.
            Dim LogWriter As StreamWriter = New StreamWriter(Path.Combine(Path.GetTempPath, My.Settings.LogFileName))

            Try
                'Report progress that we are going to update the users option.xml if it is different.
                BackgroundWorker1.ReportProgress(10, "Updating registry for Option.xml for current user...")
                UpdateOptionXML(LogWriter)
                Application.DoEvents()
                If BackgroundWorker1.CancellationPending = True Then
                    Exit Sub
                End If

                'Report progress that we are going to process the list of static files.
                BackgroundWorker1.ReportProgress(33, "Processing list of static files against network files...")
                ProcessStaticList(LogWriter)

                Application.DoEvents()
                If BackgroundWorker1.CancellationPending = True Then
                    Exit Sub
                End If
            Catch ex As Exception
                LogWriter.WriteLine(LogLine, Now.ToString, ex.Message)
            End Try

            Try
                'Report progress that we are going to check on current version of Solid Edge Maintenance Pack installed and install if necessary.
                BackgroundWorker1.ReportProgress(66, "Verifying Maintenance Pack level and installing update if necessary...")
                ServicePackInstall(LogWriter)

                Application.DoEvents()
                If BackgroundWorker1.CancellationPending = True Then
                    Exit Sub
                End If
            Catch ex As Exception
                LogWriter.WriteLine(LogLine, Now.ToString, ex.Message)
            End Try

            LogWriter.Close()
        End If

    End Sub
    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        pbMain.Value = e.ProgressPercentage
        lblProgress.Text = TryCast(e.UserState, String)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Debug.Print(DateTime.Now.ToString)
        Me.Close()
    End Sub

    Private Sub ProcessStaticList(ByRef LogWriter As System.IO.StreamWriter)
        'Verify that the static.txt file exists in the application path.
        If File.Exists(Path.Combine(My.Application.Info.DirectoryPath, "static.txt")) = False Then
            LogWriter.WriteLine(String.Format(LogLine, Now.ToString, "Static.txt file is not found."))
        Else
            'Create Reader for static.txt file
            Dim StaticList As New StreamReader(Path.Combine(My.Application.Info.DirectoryPath, "static.txt"))
            'Loop through all files found in the static.txt file.
            Do While StaticList.Peek <> -1
                Dim StaticFile As String = StaticList.ReadLine
                Dim StaticNet As String = Path.Combine(My.Settings.UpdatesPath, StaticFile)
                Dim InstallData As New SEInstallDataLib.SEInstallData
                Dim StaticLocal As String = System.IO.Path.Combine(InstallData.GetInstalledPath, StaticFile)
                InstallData = Nothing
                If File.Exists(StaticNet) = True Then
                    'File found on network updates location.
                    If File.Exists(StaticLocal) = True Then
                        'File found on local install location.
                        If clsFunctions.FileCompare(StaticNet, StaticLocal) = False Then
                            'Files are not the same.
                            My.Computer.FileSystem.CopyFile(StaticNet, StaticLocal, True)
                            LogWriter.WriteLine(String.Format(LogLine, Now.ToString, StaticFile & " has been updated."))
                        End If
                    Else
                        'File does not exist locally, copy network version to local directory.
                        My.Computer.FileSystem.CopyFile(StaticNet, StaticLocal, True)
                        LogWriter.WriteLine(String.Format(LogLine, Now.ToString, StaticFile & " has been updated."))
                    End If
                Else
                    LogWriter.WriteLine(String.Format(LogLine, Now.ToString, StaticFile & " does not exist."))
                End If
            Loop
            StaticList.Close()
        End If
    End Sub

    Private Sub ServicePackInstall(ByRef LogWriter As StreamWriter)
        Dim InstallData As New SEInstallDataLib.SEInstallData
        If CInt(My.Settings.SPVersion) > InstallData.GetServicePackVersion Then 'Newer version of Service Pack is ready to be installed.
            Dim strSPPath As String = My.Settings.SPPath
            Dim strLocalPath As String = Path.Combine(Path.GetTempPath, "SEMP")
            'Copy Maintenance Pack Directory ot local temporary directory.
            My.Computer.FileSystem.CopyDirectory(strSPPath, strLocalPath, True)

            Dim myProcess As New System.Diagnostics.Process
            myProcess.StartInfo.FileName = Path.Combine(strLocalPath, "setup.exe")
            myProcess.StartInfo.WorkingDirectory = strLocalPath
            myProcess.StartInfo.Arguments = " /s /v""/passive"""
            myProcess.StartInfo.UseShellExecute = False
            Try
                myProcess.Start()
                'Wait for Maintenance Pack to Install
                myProcess.WaitForExit()
            Catch ex As Exception
                LogWriter.WriteLine(String.Format(LogLine, Now.ToString, "Maintenance Pack failed to instal. " & ex.Message))
            End Try

            My.Computer.FileSystem.DeleteDirectory(strLocalPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

            LogWriter.WriteLine(String.Format(LogLine, Now.ToString, "Maintenance Pack Successfully installed."))
        End If
    End Sub

    Private Sub UpdateOptionXML(ByRef LogWriter As StreamWriter)
        Dim CurrentOptionXML As String
        CurrentOptionXML = My.Computer.Registry.GetValue(
            My.Settings.OptionRegKey, "AdminOptionsFile", Nothing).ToString
        If CurrentOptionXML <> My.Settings.OptionXML Then
            Dim autoshell = My.Computer.Registry.CurrentUser.OpenSubKey(
                My.Settings.OptionRegKey.Replace("HKEY_CURRENT_USER\", String.Empty))
            autoshell.SetValue("AdminOptionsFile", My.Settings.OptionXML)
            autoshell.Close()
            LogWriter.WriteLine(String.Format(LogLine, Now.ToString, "Admin Option XML file updated."))

        End If
    End Sub
End Class
