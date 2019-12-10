Imports System.IO

Module Module1

    Private mSetupDirectory As String
    Private mPracticeMgmtDirectory
    Private mForteEmrDirectory As String

    Sub Main()

        ' Sets private variable with locations
        SetLocations()

        ' Close any chiro apps that would necessitate a system restart
        EnsureChiroAndEmrClosed()

        ' Locate, launch and wait for Chiro installer to exit
        InstallPracticeManagement()

        ' Locate, launch and wait for EMR installer to exit if client already has EMR installed
        If IsForteEMRAlreadyInstalled() Then InstallForteEMRUpdate()

    End Sub

    Private Sub EnsureChiroAndEmrClosed()
        Try
            Dim runningApps As List(Of String) = New List(Of String)
            Dim execNames() As String = GetProcessExecNames()

            ' If they're running any of the apps, add it to the runningApps list
            For Each oProc As Process In Process.GetProcesses
                If execNames.Contains(oProc.ProcessName) Then runningApps.Add(oProc.ProcessName)
            Next

            ' If there's any work to be done... 
            If runningApps.Count > 0 Then

                Console.WriteLine()
                Console.WriteLine("Practice Management and EMR software must be closed to apply this update.")
                Console.WriteLine("Please save and close now or press [CTRL] + C to abort the installation.")
                Console.WriteLine()
                Console.WriteLine()

                For iIndex As Integer = 60 To 0 Step -1

                    Console.SetCursorPosition(0, 4)
                    Console.Write(String.Format("Auto-Closing applications in {0} seconds.", iIndex.ToString))

                    Threading.Thread.Sleep(1000)


                    ' Spacebar adds more time/resets whereas ENTER exits the loop immediately and starts the update
                    If Console.KeyAvailable Then
                        If Console.ReadKey.Key = ConsoleKey.Spacebar Then
                            iIndex = 60
                        ElseIf Console.ReadKey.Key = ConsoleKey.Enter Then
                            Exit For
                        End If
                    End If

                Next

                ' They were given time to save/close/terminate the update, proceed
                KillRunningApps(runningApps)

            End If

        Catch
            ' No handler, allow the installer to continue - they may need to restart

        Finally
            ' Finished with this functionality - remove it from the screen altogether
            Console.Clear()
        End Try

    End Sub

    Private Sub KillRunningApps(appsToKill As List(Of String))

        For Each strApp As String In appsToKill
            KillProc(strApp)
        Next

    End Sub

    Private Sub KillProc(procName As String)

        Try

            For Each oProc As Process In Process.GetProcessesByName(procName)
                oProc.Kill()
            Next

        Catch
            ' No catch, they may have to restart since we can't kill whatever process it is
        End Try

    End Sub


    Private Function GetProcessExecNames() As String()
        Return {"ASHN", "Billing", "CAWC", "CE106466", "CopayRestoreUtility", "CustomFormGenerator",
                                 "DataExtractionUtility", "Daysheet", "DBUtility", "DocumentPlus", "EMDEONIntegration", "EZNotes",
                                 "FileServer", "FormsCenter", "ForteEMR", "Graphs", "Inventory", "LabCorp", "MediNotes", "PatientPortionUtility",
                                 "PayorIDUtility", "PM", "PolicyManual", "ProspectCenter", "RecordCenter", "ReportsModule", "SpringCharts",
                                 "StartingBalanceUtility", "TelevoxExport", "UserOptions", "WritePad"}
    End Function

    Private Function IsForteEMRAlreadyInstalled() As Boolean

        Dim keySoftware As Microsoft.Win32.RegistryKey
        Dim keyForteHoldings As Microsoft.Win32.RegistryKey

        Try

            ' Base location
            keySoftware = My.Computer.Registry.LocalMachine.OpenSubKey("Software")

            ' Check primary location
            keyForteHoldings = keySoftware.OpenSubKey("Forte Holdings")

            ' If it wasn't found, check wow6432Node
            If keyForteHoldings Is Nothing Then
                keyForteHoldings = keySoftware.OpenSubKey("Wow6432Node").OpenSubKey("Forte Holdings")
            End If

            ' If this is still nothing - then return false
            If keyForteHoldings Is Nothing Then
                Return False
            End If

            ' If the ForteEMR subkey is not nothing, then return true
            Return Not (keyForteHoldings.OpenSubKey("ForteEMR") Is Nothing)

        Catch
            Console.WriteLine("**********************************************************************************")
            Console.WriteLine("* Unable to determine EMR installation status.  Install EMR module update (Y/N)? *")
            Console.WriteLine("**********************************************************************************")

            If Console.ReadLine().Trim.Equals("y", StringComparison.OrdinalIgnoreCase) Then
                Return True
            Else
                Return False
            End If

        End Try

    End Function

    Private Sub InstallPracticeManagement()

        Dim strExecPath As String = Path.Combine(mPracticeMgmtDirectory, "setup.exe")

        If File.Exists(strExecPath) Then

            Console.WriteLine("Launching Practice Management update...")
            ShellandWait(strExecPath)

        Else

            Console.WriteLine("Practice Management Setup.exe was not found.")
            Console.WriteLine("Press <enter> to continue.")
            Console.ReadLine()

        End If

    End Sub

    Private Sub InstallForteEMRUpdate()

        If String.IsNullOrEmpty(mForteEmrDirectory) Then Return

        Dim strExecPath As String = Path.Combine(mForteEmrDirectory, "setup.exe")

        If File.Exists(strExecPath) Then

            Console.WriteLine("Launching ForteEMR update...")
            ShellandWait(strExecPath)

        Else

            Console.WriteLine("ForteEMR Setup.exe was not found.")
            Console.WriteLine("Press <enter> to continue.")
            Console.ReadLine()

        End If

    End Sub

    Private Function ShellandWait(ByVal ProcessPath As String) As Boolean
        Dim objProcess As Process
        Try
            objProcess = New Process()
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()

            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()

            'Free resources associated with this process
            objProcess.Close()

            Return True
        Catch
            Return False
        End Try

    End Function
    Private Sub SetLocations()
        Dim strTemp As String = String.Empty

        ' Get location of the app
        mSetupDirectory = My.Application.Info.DirectoryPath

        ' 1 is for Chiro
        strTemp = Path.Combine(mSetupDirectory, "1")
        If Directory.Exists(strTemp) Then
            mPracticeMgmtDirectory = strTemp
        End If

        ' 5 is for HealthPro
        strTemp = Path.Combine(mSetupDirectory, "5")
        If Directory.Exists(strTemp) Then
            mPracticeMgmtDirectory = strTemp
        End If

        strTemp = Path.Combine(mSetupDirectory, "EMR")
        If Directory.Exists(strTemp) Then
            mForteEmrDirectory = strTemp
        End If

    End Sub

End Module
