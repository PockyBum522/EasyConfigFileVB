Imports System.Text.RegularExpressions
Imports System.IO

Class EasyConfigFile
    Property ConfigFilePath As String

    Public Shared ReadOnly userFolder As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) & "\"
    Public Shared ReadOnly userTempFolder As String = IO.Path.GetTempPath

    Public Function ReadSetting(settingNameToRead As String, ByRef settingContent As String) As Boolean
        Dim TextLine As String
        Dim firstCharPos As Integer
        Dim endOfLineCharPos As Integer
        Dim foundSetting As Boolean = False
        Dim lengthToGet As Integer

        Me.MakeFileStructure(ConfigFilePath)

        If System.IO.File.Exists(ConfigFilePath) = True Then

            Dim cfgFileReader As New System.IO.StreamReader(ConfigFilePath)

            'Validate what we're trying to read
            If InStr(settingNameToRead, "~") Or InStr(settingNameToRead, "`") Then
                MessageBox.Show("Error: Setting name may not contain tildes (~) or backticks (`)")
                Return False
            End If

            'Look for setting save line
            Do While cfgFileReader.Peek() <> -1
                TextLine = cfgFileReader.ReadLine()

                If (InStr(TextLine, settingNameToRead + "~")) Then
                    firstCharPos = InStr(TextLine, "~")

                    If (firstCharPos > 0) Then
                        'MessageBox.Show("The String was found.", "iLogic")
                        endOfLineCharPos = InStr(firstCharPos, TextLine, "`")
                        lengthToGet = endOfLineCharPos - firstCharPos + 1
                        settingContent = TextLine.Substring(firstCharPos, lengthToGet - 2)

                        ' Make sure we didn't find a setting that just has nothing in it.
                        If lengthToGet > 0 Then
                            foundSetting = True
                        End If
                    End If
                End If
            Loop

            cfgFileReader.Close()
        Else
            MessageBox.Show("File read error: Cannot find " + ConfigFilePath)

            Return foundSetting
        End If

        Return foundSetting
    End Function

    Public Function WriteSetting(settingNameToWrite As String, settingContent As String) As Boolean
        Dim fileContents As String = ""
        Dim newContent As String = ""
        Dim wroteDataSuccess As Boolean = False

        Me.MakeFileStructure(ConfigFilePath)

        'Validate what we're about to write
        If InStr(settingNameToWrite, "~") Or InStr(settingNameToWrite, "`") _
            Or InStr(settingContent, "~") Or InStr(settingContent, "`") Then

            MessageBox.Show("Error: Setting name and contents may not contain tildes (~) or backticks (`)")
            Return False
        Else
            If Me.ReadFullFile(fileContents) Then
                'If the file read in okay, see if the setting exists in it
                Dim expressionToTest As Regex = New Regex(settingNameToWrite + "~")
                Dim regexMatch As Match = expressionToTest.Match(fileContents) 'Check it against expr

                If regexMatch.Success Then
                    'If the setting does exist in the contents
                    newContent = Regex.Replace(fileContents, settingNameToWrite + "~.+?`", settingNameToWrite + "~" + settingContent + "`")
                Else
                    newContent = settingNameToWrite + "~" + settingContent + "`" + vbNewLine + fileContents
                End If
            End If

            Try
                If newContent <> "" Then
                    Try
                        Dim cfgFileWriter As StreamWriter = New StreamWriter(ConfigFilePath, False)

                        cfgFileWriter.Write(newContent)
                        cfgFileWriter.Flush()
                        cfgFileWriter.Close()

                        Try
                            cfgFileWriter.Dispose()
                        Catch
                        End Try

                        wroteDataSuccess = True
                    Catch
                        MessageBox.Show("Error writing file " + ConfigFilePath)
                    End Try
                End If
            Catch
                MessageBox.Show("Error writing file " + ConfigFilePath)

                Return False
            End Try

        End If

        Return wroteDataSuccess
    End Function

    Private Function ReadFullFile(ByRef fileContent As String) As Boolean
        fileContent = ""
        Dim readFile As Boolean

        If System.IO.File.Exists(ConfigFilePath) = True Then

            Dim cfgFileReader As New System.IO.StreamReader(ConfigFilePath)

            'Look for setting save line
            Do While cfgFileReader.Peek() <> -1
                fileContent += cfgFileReader.ReadLine() + vbNewLine
            Loop

            cfgFileReader.Close()
            readFile = True
        Else
            Return False
        End If

        Return readFile
    End Function

    Public Sub New(configFilePath As String)
        Me.ConfigFilePath = configFilePath
    End Sub

    Private Sub MakeFileStructure(fullPath As String)
        If fullPath <> "" Then
            Try
                My.Computer.FileSystem.CreateDirectory(Path.GetDirectoryName(fullPath))

                If System.IO.File.Exists(fullPath) = False Then
                    System.IO.File.Create(fullPath).Dispose()
                End If
            Catch ex As FileNotFoundException
                MessageBox.Show("Error 126: Can't create data folder structure." + vbNewLine + vbNewLine +
                                    "Program data: " + ex.ToString)
            End Try
        Else
            MessageBox.Show("Error 127: folderPath Empty")
        End If
    End Sub
End Class
