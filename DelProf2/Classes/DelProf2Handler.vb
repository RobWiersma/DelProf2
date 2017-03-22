Imports System.IO
Imports System.Text

Public Class DelProf2Handler

    Dim machineNames As String() = {"MACHINE1", "MACHINE2", "MACHINE3"}

    Public Function deleteProfiledOnAllServers(username As String, machineList As List(Of String)) As List(Of String)

        'delete profiles in machinelist and return a list of successful deletion to check against input
        Try
            Dim successfulDeleteResults As New List(Of String)

            Parallel.For(0, machineList.Count, Sub(i)
                                                   Dim returnResult As Boolean = deleteProfileOnServer(machineList(i), username)
                                                   If returnResult = True Then
                                                       successfulDeleteResults.Add(machineList(i))
                                                   End If
                                               End Sub)
            Debug.WriteLine("Pause")
            Return successfulDeleteResults
        Catch ex As Exception
            Throw New Exception("There was an error deleting profiles for account " + username + vbNewLine + ex.ToString)
        End Try

    End Function

    Public Function checkProfilesXDaysOldOnAllServers(numDays As String) As Dictionary(Of String, List(Of String))

        'Checks profiles on each XEN server and returns a dictionary in the format of username:list of servers it resides on (that hasnt been used in x amount of days)
        Try
            Dim machineResults As New Dictionary(Of String, List(Of String))

            Parallel.For(0, machineNames.Count, Sub(i)
                                                    Dim returnedUsernameList As List(Of String) = findProfilesXDaysOld(machineNames(i), numDays)

                                                    For Each username As String In returnedUsernameList
                                                        If Not machineResults.Keys.Contains(username) Then
                                                            machineResults.Add(username, New List(Of String))
                                                        End If
                                                        machineResults(username).Add(machineNames(i))
                                                    Next
                                                End Sub)
            'Debug.WriteLine("Pause")
            Return machineResults
        Catch ex As Exception
            Throw New Exception("There was an error checking all XEN Servers for accounts " + numDays + " old" + vbNewLine + ex.ToString)
        End Try

    End Function
    Public Function checkProfileOnAllServers(username As String) As Dictionary(Of String, Boolean)

        'Check if profile exists on all XEN Servers
        Try
            Dim machineResults As New Dictionary(Of String, Boolean)

            Parallel.For(0, machineNames.Count, Sub(i)
                                                    Debug.WriteLine("Does profile exist on machine: " + machineNames(i) + "? " + doesProfileExistOnMachine(machineNames(i), username).ToString)
                                                    machineResults.Add(machineNames(i), doesProfileExistOnMachine(machineNames(i), username))
                                                End Sub)
            'Debug.WriteLine("Pause")
            Return machineResults
        Catch ex As Exception
            Throw New Exception("There was an error checking all XEN Servers for Profile" + vbNewLine + ex.ToString)
        End Try

    End Function

    Private Function deleteProfileOnServer(computerName As String, username As String) As Boolean

        'DO NOT RUN THIS IF YOU DONT KNOW WHAT IT DOES, YOU CAN ROYALLY SCREW PROFILES
        'Deletes a single profile on a single machine

        Try

            Dim outputString As String = ""

            Dim currentdirectory As String = System.AppDomain.CurrentDomain.BaseDirectory

            Dim command As String = "/c:" + computerName + " /id:" + username + " /u"

            Dim startInfo As New ProcessStartInfo("DelProf2.exe", command)
            startInfo.UseShellExecute = False
            startInfo.WorkingDirectory = currentdirectory
            startInfo.Verb = "RunAs"
            startInfo.RedirectStandardOutput = True
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
            startInfo.CreateNoWindow = True

            Dim zipper As System.Diagnostics.Process = System.Diagnostics.Process.Start(startInfo)
            outputString = zipper.StandardOutput.ReadToEnd
            zipper.StandardOutput.Close()

            Dim positiveMatchString As String = "The following user profiles match the deletion criteria:"

            If outputString.Contains("No user profiles match the deletion criteria") Then
                Return False
            ElseIf outputString.Contains("The following user profiles match the deletion criteria:") Then
                'shorten outputstring to the good part at the end
                Dim outputStringLength As Integer = outputString.Count
                Dim indexOfDeletion As Integer = outputString.IndexOf(positiveMatchString) + positiveMatchString.Count
                outputString = outputString.Substring(indexOfDeletion, outputStringLength - indexOfDeletion)

                'check for successful processing of the job
                If outputString.Contains("... done.") Then
                    Return True
                Else
                    Return False
                End If
            Else
                Debug.WriteLine("Something went wrong here...")
                Return False
            End If

            Return False
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw New Exception("There was an error trying to delete user " + username + " on server " + computerName + vbNewLine + ex.ToString)
        End Try

    End Function

    Private Function doesProfileExistOnMachine(computerName As String, username As String) As Boolean

        'Use delprof2 to look up machine and read console output to check if we have a valid entry

        Try

            Dim outputString As String = ""

            Dim currentdirectory As String = System.AppDomain.CurrentDomain.BaseDirectory

            Dim command As String = "/l /c:" + computerName + " /id:" + username

            Dim startInfo As New ProcessStartInfo("DelProf2.exe", command)
            startInfo.UseShellExecute = False
            startInfo.WorkingDirectory = currentdirectory
            startInfo.Verb = "RunAs"
            startInfo.RedirectStandardOutput = True
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
            startInfo.CreateNoWindow = True

            Dim zipper As System.Diagnostics.Process = System.Diagnostics.Process.Start(startInfo)
            outputString = zipper.StandardOutput.ReadToEnd
            zipper.StandardOutput.Close()

            If outputString.Contains("No user profiles match the deletion criteria") Then
                Return False
            ElseIf outputString.Contains("The following user profiles match the deletion criteria:") Then
                Return True
            Else
                Debug.WriteLine("Something went wrong here...")
                Return False
            End If

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw New Exception("There was an error checking if the profile: " + username + " exists on " + computerName + vbNewLine + ex.ToString)
        End Try

    End Function

    Private Function findProfilesXDaysOld(computerName As String, numOfDays As String) As List(Of String)

        'Look up inactive profiles X number of days old and return a list of those usernames

        Try

            Dim outputString As String = ""

            Dim currentdirectory As String = System.AppDomain.CurrentDomain.BaseDirectory

            Dim command As String = "/l /c:" + computerName + " /d:" + numOfDays

            Dim startInfo As New ProcessStartInfo("DelProf2.exe", command)
            startInfo.UseShellExecute = False
            startInfo.WorkingDirectory = currentdirectory
            startInfo.Verb = "RunAs"
            startInfo.RedirectStandardOutput = True
            startInfo.WindowStyle = ProcessWindowStyle.Hidden
            startInfo.CreateNoWindow = True

            Dim zipper As System.Diagnostics.Process = System.Diagnostics.Process.Start(startInfo)
            outputString = zipper.StandardOutput.ReadToEnd
            zipper.StandardOutput.Close()

            Dim positiveMatchString As String = "The following user profiles match the deletion criteria:"

            If outputString.Contains("No user profiles match the deletion criteria") Then
                Return Nothing
            ElseIf outputString.Contains(positiveMatchString) Then

                'Read through each line and gather usernames
                Dim outputStringLength As Integer = outputString.Count
                Dim indexOfDeletion As Integer = outputString.IndexOf(positiveMatchString) + positiveMatchString.Count
                outputString = outputString.Substring(indexOfDeletion, outputStringLength - indexOfDeletion)
                Dim outputStringLines As String() = outputString.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

                Dim usernameList As New List(Of String)

                For Each line As String In outputStringLines
                    If line <> "" Then
                        Dim indexOfUsers As Integer = line.IndexOf("\Users\") + 7
                        Dim username As String = line.Substring(indexOfUsers, line.Length - indexOfUsers)
                        usernameList.Add(username)
                    End If
                Next

                'Debug.WriteLine("Pause")

                Return usernameList
            Else
                Debug.WriteLine("Something went wrong here...")
                Return Nothing
            End If

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Throw New Exception("There was an error checking for profiles on " + computerName + " " + numOfDays + " old" + vbNewLine + ex.ToString)
        End Try

    End Function

    Public Function printCSVAuditXNumberOfDaysOld(numDays As String) As Boolean

        Try
            Dim oldProfileAuditDictionary As Dictionary(Of String, List(Of String)) = checkProfilesXDaysOldOnAllServers(numDays)

            Dim sb As New StringBuilder

            Dim headerLine As String = "username"
            For Each machineName As String In machineNames
                headerLine &= "," + machineName
            Next

            sb.AppendLine(headerLine)

            For Each key As String In oldProfileAuditDictionary.Keys
                Dim XENServers(machineNames.Count - 1) As String

                For i As Integer = 0 To machineNames.Count - 1
                    If oldProfileAuditDictionary(key).Contains(machineNames(i)) Then
                        XENServers(i) = "True"
                    End If
                Next

                Dim usernameLine As String = key
                For Each serverStatus As String In XENServers
                    usernameLine &= "," + serverStatus
                Next

                sb.AppendLine(usernameLine)
            Next

            Dim CSVString As String = sb.ToString

            Dim filename As String = "CSVAudit" + numDays + "DaysOld.csv"
            File.WriteAllText(filename, CSVString)

            'Debug.WriteLine("Pause")

            Return True
        Catch ex As Exception
            Throw New Exception("There was an error printing CSV Audit for profiles " + numDays + " old." + vbNewLine + ex.ToString)
        End Try

    End Function

End Class
