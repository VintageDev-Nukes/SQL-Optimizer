#Region "Imports"

Imports System.IO
Imports System.Text.RegularExpressions

#End Region

Public Class frmMain

#Region "TabControl1"

#Region "Variables"

    Dim remError As String = Nothing
    Dim remWarning As String = Nothing
    Dim remMsg As String = Nothing
    Dim totalCorrections As Integer = 0
    Dim ComillasEscapadas As Integer = 0
    Dim ComillasPuestas As Integer = 0
    Dim totalCh As Integer = 0

#End Region

#Region "Funciones"

    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then cnt += 1
        Next
        Return cnt
    End Function

    Function ToUnixTime(time As Date) As Long
        ToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), time)
        Return ToUnixTime
    End Function

#End Region

#Region "Eventos Form General"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextLabel.Add_Text_With_Color(RichTextBox1, "SQL String Cleaner", RichTextBox1.ForeColor, New Font("Microsoft Sans serif", 8, FontStyle.Bold))
        RichTextLabel.Add_Text_With_Color(RichTextBox1, " es una herramienta con la que podrás limpiar o escapar grandes cadenas de texto sin esfuerzo alguno.", RichTextBox1.ForeColor, New Font("Microsoft Sans serif", 8, FontStyle.Regular))
        TextBox2.Text = Application.StartupPath.ToString
        Label4.Text = "© Ikillnukes - " & Date.Now.Year & " | v0.0.1 Alpha"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            With OpenFileDialog1

                .Reset() ' resetea

                .InitialDirectory = Environment.SpecialFolder.Desktop
                .Filter = "Sql Files (*.sql)|*.sql|All files (*.*)|*.*"
                .FilterIndex = 1
                .RestoreDirectory = True

                Dim ret As DialogResult = .ShowDialog ' abre el diálogo

                ' si se presionó el botón aceptar ...
                If ret = Windows.Forms.DialogResult.OK Then

                    TextBox1.Text = .FileName

                End If

                .Dispose()

            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
            ' Configuración del FolderBrowserDialog
            With FolderBrowserDialog1

                .Reset() ' resetea

                ' leyenda
                .Description = " Seleccionar una carpeta "
                ' Path " Mis documentos "
                .SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                ' deshabilita el botón " crear nueva carpeta "
                .ShowNewFolderButton = True
                '.RootFolder = Environment.SpecialFolder.Desktop
                '.RootFolder = Environment.SpecialFolder.StartMenu

                Dim ret As DialogResult = .ShowDialog ' abre el diálogo

                ' si se presionó el botón aceptar ...
                If ret = Windows.Forms.DialogResult.OK Then

                    TextBox2.Text = .SelectedPath

                End If

                .Dispose()

            End With
        Catch oe As Exception
            MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    'Dim WithEvents Timer1 As New Timer With {.Interval = 1}

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    Timer1.Enabled = True
    'End Sub

    'Private Sub InfiniteProgressBar(ByVal ProgressBar As ProgressBar, _
    '                                Optional value As Int32 = 1)

    '    Select Case ProgressBar.Value

    '        Case Is < ProgressBar.Maximum
    '            ProgressBar.Value += value
    '        Case Is >= ProgressBar.Maximum ' Si el valor es igual o mayor que el valor máximo del rango...
    '            ProgressBar.Value = ProgressBar.Minimum ' Seteamos el valor mínimo (0) a la barra de progreso...

    '    End Select

    'End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs)
    '    InfiniteProgressBar(ProgressBar1)
    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            TextBox3.AppendText("[" & Date.Now & "] " & "Comenzando a leer y filtrar resultados del archivo...")
            Dim Texto As String = File.ReadAllText(TextBox1.Text)
            Dim Match As Match = Regex.Match(Texto, "(?<prueba>(?<=[\(]).*(?=[\)]))")
            Dim Match2 As Match = Regex.Match(Texto, "(?<prueba2>.(.*(?=[\(])))")
            'Do While Match.Success
            '    MsgBox(Match.Groups(1).ToString)
            '    Match = Match.NextMatch()
            'Loop

            Dim preFileName As String = TextBox1.Text.Substring(TextBox1.Text.LastIndexOf("\") + 1)
            Dim FileFormat As String = preFileName.Substring(preFileName.LastIndexOf(".") + 1)
            Dim FileName As String = preFileName.Replace("." & FileFormat, "")

            'MsgBox(FileName)

            Dim strStreamWriter As StreamWriter = Nothing
            Dim strStreamW As Stream = Nothing
            Dim PathArchivo As String = ".\Variables.txt"
            Dim PathArchivo2 As String = ".\INSERTS.txt"
            Dim PathArchivo3 As String = ".\preVariables.txt"
            Dim PathArchivo4 As String = TextBox2.Text & "\" & FileName & "-reparado." & FileFormat

            Dim PathCarpeta As String = Replace(TextBox1.Text, PathArchivo.Substring(PathArchivo.LastIndexOf("/") + 1), "")

            ProgressBar1.Value = 5

            'If Not Directory.Exists(PathCarpeta) Then
            '    Directory.CreateDirectory(PathCarpeta)
            'End If

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Separando archivo...")

            Do While Match.Success

                Try

                    If File.Exists(PathArchivo) Then
                        strStreamW = File.Open(PathArchivo, FileMode.Append)
                    Else
                        strStreamW = File.Create(PathArchivo)
                    End If

                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                    strStreamWriter.Write(Match.Groups("prueba").ToString & Environment.NewLine)
                    Match = Match.NextMatch()

                    strStreamWriter.Close()

                    strStreamW = Nothing
                    strStreamWriter = Nothing

                    If File.Exists(PathArchivo2) Then
                        strStreamW = File.Open(PathArchivo2, FileMode.Append)
                    Else
                        strStreamW = File.Create(PathArchivo2)
                    End If

                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                    strStreamWriter.Write(Match2.Groups("prueba2").ToString & Environment.NewLine)
                    Match2 = Match2.NextMatch()

                    strStreamWriter.Close()

                    strStreamW = Nothing
                    strStreamWriter = Nothing

                Catch ex As Exception
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Error al Guardar la información en el archivo. ")
                    strStreamWriter.Close()
                End Try
            Loop

            ProgressBar1.Value = 30

            Dim lines As String() = File.ReadAllLines(TextBox1.Text)

            Dim MaxCommas As Integer = CountCharacter(lines(0), ",") + 1

            Dim TimesSplitted As Integer = 0

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Leyendo todas las líneas del archivo...")

            Dim Line As String = File.ReadAllText(PathArchivo)

            ProgressBar1.Value = 35

            'Dim Result As String() = Line.Split(", ")

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Corrigiendo errores del Strings...")

            Dim RegReplace As String = Line.Replace(Environment.NewLine, "," & Environment.NewLine)

            For Each Result As String In RegReplace.Split(",")

                Select Case True

                    Case Result.StartsWith("'") AndAlso Result.EndsWith("'")
                        Dim posResultado = Result
                        Dim preResultado = String.Format("'{0}'", Result.Substring(1, Result.Length - 2).Replace("'", "\'"))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write(preResultado & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        If Not preResultado = posResultado Then
                            remMsg = "1"
                            totalCh = CountCharacter(preResultado, "\'")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If

                    Case Result.StartsWith("'") AndAlso Not Result.EndsWith("'")
                        Dim posResultado = Result
                        Dim preResultado As String = String.Format("'{0}", Result.Substring(1, Result.Length - 1).Replace("'", "\'"))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write(preResultado & "'" & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        ComillasPuestas += 1

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\'")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If
                        remError = "1"

                    Case Not Result.StartsWith("'") AndAlso Result.EndsWith("'")
                        Dim posResultado = Result
                        Dim preResultado As String = Result.Replace("'", "\'")
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing


                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write("'" & preResultado & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        If Not preResultado = posResultado Then
                            remMsg = "1"
                        End If
                        remError = "2"

                    Case Not Result.StartsWith("'") AndAlso Not Result.EndsWith("'") AndAlso Result.Contains("'")
                        Dim posResultado = Result
                        Dim preResultado As String = String.Format("{0}", Result.Replace("'", "\'"))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write("'" & preResultado & "'" & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        ComillasPuestas += 2

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\'")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If
                        remError = "3"

                        '-----------------------

                    Case Result.StartsWith("""") AndAlso Result.EndsWith("""")
                        Dim posResultado = Result
                        Dim preResultado = String.Format("""{0}""", Result.Substring(1, Result.Length - 2).Replace("""", "\"""))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write(preResultado & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\""")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If

                    Case Result.StartsWith("""") AndAlso Not Result.EndsWith("""")
                        Dim posResultado = Result
                        Dim preResultado As String = String.Format("""{0}", Result.Substring(1, Result.Length - 1).Replace("""", "\"""))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write(preResultado & """" & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        ComillasPuestas += 1

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\""")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If
                        remError = "1"

                    Case Not Result.StartsWith("""") AndAlso Result.EndsWith("""")
                        Dim posResultado = Result
                        Dim preResultado As String = Result.Replace("""", "\""")
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write("""" & preResultado & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        ComillasPuestas += 1

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\""")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If
                        remError = "2"

                    Case Not Result.StartsWith("""") AndAlso Not Result.EndsWith("""") AndAlso Result.Contains("""")
                        Dim posResultado = Result
                        Dim preResultado As String = String.Format("{0}", Result.Replace("""", "\"""))
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing

                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write("""" & preResultado & """" & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        ComillasPuestas += 2

                        If Not preResultado = posResultado Then
                            remMsg = "1"

                            totalCh = CountCharacter(preResultado, "\""")
                            ComillasEscapadas += totalCh
                            totalCh = 0
                        End If
                        remError = "3"

                    Case Else
                        Try

                            If File.Exists(PathArchivo3) Then
                                strStreamW = File.Open(PathArchivo3, FileMode.Append)
                            Else
                                strStreamW = File.Create(PathArchivo3)
                            End If

                            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                            Dim Splittado As String = Nothing


                            TimesSplitted += 1

                            If TimesSplitted = MaxCommas Then
                                Splittado = ""
                                TimesSplitted = 0
                            ElseIf TimesSplitted < MaxCommas Then
                                Splittado = ","
                            End If

                            strStreamWriter.Write(Result & Splittado)

                            strStreamWriter.Close()

                            strStreamW = Nothing
                            strStreamWriter = Nothing

                        Catch ex As Exception
                            remError = "4"
                            strStreamWriter.Close()
                        End Try

                        'remError = "5"

                End Select

                'MsgBox(remError)

                If remMsg = "1" Then
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Msg: #1 (String Cambiada)")
                ElseIf remError = "1" Then
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Msg: #2 (String arreglada; Error #1: [Faltaba la última comilla])")
                ElseIf remError = "2" Then
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Msg: #2 (String arreglada; Error: #2: [Faltaba la primera comilla])")
                ElseIf remError = "3" Then
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Msg: #2 (String arreglada; Error: #3: [Faltaban las comillas de extremos])")
                ElseIf remError = "4" Then
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Error al Guardar la información en el archivo.")
                    'ElseIf remError = "5" Then
                    'TextBox3.AppendText("Bug Tremendoooorl!")
                End If

                remError = Nothing
                remWarning = Nothing
                remMsg = Nothing

            Next Result

            ProgressBar1.Value = 70

            totalCorrections = ComillasEscapadas + ComillasPuestas

            Dim preLines() As String = File.ReadAllLines(PathArchivo3)
            Dim preLines2() As String = File.ReadAllLines(PathArchivo2)
            Dim LinesMax As Integer = preLines.Length
            Dim LineTime As Integer = -1

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Ensamblando todo el código de nuevo...")

            Dim PathArchivo5 As String = ".\SQLOptiminer-StringCleaner-" & ToUnixTime(Date.Now) & ".log"

            Do

                LineTime += 1

                Try

                    If File.Exists(PathArchivo4) Then
                        strStreamW = File.Open(PathArchivo4, FileMode.Append)
                    Else
                        strStreamW = File.Create(PathArchivo4)
                    End If

                    strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                    strStreamWriter.Write(preLines2(LineTime) & " (" & preLines(LineTime) & ");" & Environment.NewLine)

                    strStreamWriter.Close()

                    strStreamW = Nothing
                    strStreamWriter = Nothing

                Catch ex As Exception
                    TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Error al Guardar la información en el archivo.")
                    strStreamWriter.Close()
                End Try

            Loop Until LineTime = LinesMax

            ProgressBar1.Value = 95

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Registrando todo en un Log, espere...")

            Try

                If File.Exists(PathArchivo5) Then
                    strStreamW = File.Open(PathArchivo5, FileMode.Append)
                Else
                    strStreamW = File.Create(PathArchivo5)
                End If

                strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

                strStreamWriter.WriteLine("SQL Optimizer [String Cleaner] v0.0.1 Alpha" & vbCrLf & vbCrLf & "[" & Date.Now & "] Correciones totales hechas: " & totalCorrections & vbCrLf & "   - Comillas escapadas: " & ComillasEscapadas & vbCrLf & "   - Comillas puestas: " & ComillasPuestas & vbCrLf & vbCrLf & "Log generado automáticamente por SQLOptimizer" & vbCrLf & "Aplicación hecha por Ikillnukes")

                strStreamWriter.Close()

                strStreamW = Nothing
                strStreamWriter = Nothing

            Catch ex As Exception
                TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "Error al Guardar la información en el archivo.")
                strStreamWriter.Close()
            End Try

            TextBox3.AppendText(Environment.NewLine & "[" & Date.Now & "] " & "¡Proceso terminado!")

            ProgressBar1.Value = 100

            If MsgBox("¡Proceso terminado! ¿Desea abrir el Log?", MsgBoxStyle.YesNo, "Información") = MsgBoxResult.Yes Then
                System.Diagnostics.Process.Start("Notepad.Exe", PathArchivo5)
            End If

            ComillasEscapadas = 0
            ComillasPuestas = 0
            totalCorrections = 0

            File.Delete(PathArchivo)
            File.Delete(PathArchivo2)
            File.Delete(PathArchivo3)

        Catch ex As Exception When TextBox1.Text = String.Empty
            MsgBox("¡Tienes que seleccionar un archivo!", MsgBoxStyle.Critical, "Error")
        Catch ex As Exception When TextBox2.Text = String.Empty
            MsgBox("¡Te has olvidado de la carpeta de Destino!", MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

#End Region

#Region "TabControl2"



#End Region

End Class
