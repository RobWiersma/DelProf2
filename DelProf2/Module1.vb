Module Module1

    Sub Main()

        Dim delprof2 As New DelProf2Handler

        Dim machineList As New List(Of String)
        machineList.Add("MACHINE1")
        machineList.Add("MACHINE2")

        delprof2.deleteProfiledOnAllServers("testUsername", machineList)

    End Sub

End Module
