Imports Microsoft.Win32

Module Module1

    Sub Main()

        Console.WriteLine("Find Registered Com Dll - By Joel Meikle")

        Console.Write("Enter dll name to search for:  ")
        Dim szSearch As String = Console.ReadLine

        If String.IsNullOrEmpty(szSearch) Then Return

        'Console.WriteLine("Looking for: " & szSearch)
        Console.Write("Return first result (y/n):  ")
        'Dim szQuickSearch As String = Console.ReadLine
        Dim bQuickSearch As Boolean = Console.ReadKey.Key = ConsoleKey.Y
        Console.WriteLine()

        Dim lsResults = FindKeys(szSearch, bQuickSearch) ' LCase(szQuickSearch) = "y")

        If lsResults Is Nothing OrElse lsResults.Count = 0 Then
            Console.WriteLine("Found 0 results for " & szSearch)
        Else
            Console.WriteLine(String.Format("Found {0} results for {1}", lsResults.Count, szSearch))
            For Each szvalue In lsResults
                Console.WriteLine(szvalue)
            Next


            Console.WriteLine()
            Console.Write("Do you want to open this folder now? (y/n) :  ")

            'If LCase(Console.ReadLine) = "y" Then
            If Console.ReadKey.Key = ConsoleKey.Y Then
                OpenResultFolder(lsResults(0))
            End If

        End If

        Console.WriteLine()
        Console.Write("Press Enter to exit")
        Console.ReadLine()

        'Environment.Is64BitOperatingSystem

    End Sub

    Private Sub OpenResultFolder(resultValue As String)

        Try
            'System.Diagnostics.Process.Start("explorer", IO.Path.GetDirectoryName(resultValue))
            System.Diagnostics.Process.Start("explorer", "/select," & resultValue)
        Catch ex As Exception
            Console.WriteLine("Error opening folder:")
            Console.WriteLine(ex.Message)
        End Try

    End Sub

    Private Function FindKeys(search As String, returnFirst As Boolean) As List(Of String)

        'Console.WriteLine(Microsoft.Win32.Registry.ClassesRoot.Name)
        Dim lsResults As New List(Of String)

        Dim lPath As String = "SOFTWARE\Classes"
        If Environment.Is64BitOperatingSystem Then lPath &= "\Wow6432Node"
        lPath &= "\CLSID"

        Dim searchRoot = Registry.LocalMachine.OpenSubKey(lPath, False)

        Dim arNames = searchRoot.GetSubKeyNames

        For Each szCLSID In arNames
            'does it have a sub key of InProcServer32 ?
            Using lSearchSubKey = searchRoot.OpenSubKey(szCLSID & "\InprocServer32", False)

                If lSearchSubKey Is Nothing OrElse lSearchSubKey.ValueCount = 0 Then
                    'If lSearchSubKey IsNot Nothing Then lSearchSubKey.Close
                    Continue For 'doesn't exist?
                End If


                Dim szInProcDefaultValue As String = CStr(lSearchSubKey.GetValue(String.Empty, String.Empty))
                If String.IsNullOrEmpty(szInProcDefaultValue) Then Continue For

                'does it have our dll name?
                If LCase(szInProcDefaultValue).Contains(LCase(search)) Then
                    'match...
                    'lsResults.Add(szInProcDefaultValue & vbCrLf & " (CLSID = " & szCLSID & ")")

                    If Not lsResults.Contains(szInProcDefaultValue) Then
                        lsResults.Add(szInProcDefaultValue)
                    End If

                    If returnFirst Then Exit For
                End If

                'lSearchSubKey.Close()

            End Using
        Next

        arNames = Nothing

        Return lsResults
    End Function

End Module
