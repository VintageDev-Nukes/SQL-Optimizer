'La Class de Oro:

Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Net.Mail
Imports Ionic.Zip
Imports System.Diagnostics

'========================================

'Class Low Mouse Hook hecha por Sim0n: http://sim0n.wordpress.com/2009/03/28/vbnet-mouse-hook-class/

Public Class MouseHook
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As MouseProcDelegate, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As MSLLHOOKSTRUCT) As Integer
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Integer) As Integer
    Private Delegate Function MouseProcDelegate(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer

    Private Structure MSLLHOOKSTRUCT
        Public pt As Point
        Public mouseData As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    Public Enum Wheel_Direction
        WheelUp
        WheelDown
    End Enum

    Private Const HC_ACTION As Integer = 0
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_LBUTTONDBLCLK As Integer = &H203
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205
    Private Const WM_RBUTTONDBLCLK As Integer = &H206
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_MBUTTONUP As Integer = &H208
    Private Const WM_MBUTTONDBLCLK As Integer = &H209
    Private Const WM_MOUSEWHEEL As Integer = &H20A

    Private MouseHook As Integer
    Private MouseHookDelegate As MouseProcDelegate

    Public Event Mouse_Move(ByVal ptLocat As Point)
    Public Event Mouse_Left_Down(ByVal ptLocat As Point)
    Public Event Mouse_Left_Up(ByVal ptLocat As Point)
    Public Event Mouse_Left_DoubleClick(ByVal ptLocat As Point)
    Public Event Mouse_Right_Down(ByVal ptLocat As Point)
    Public Event Mouse_Right_Up(ByVal ptLocat As Point)
    Public Event Mouse_Right_DoubleClick(ByVal ptLocat As Point)
    Public Event Mouse_Middle_Down(ByVal ptLocat As Point)
    Public Event Mouse_Middle_Up(ByVal ptLocat As Point)
    Public Event Mouse_Middle_DoubleClick(ByVal ptLocat As Point)
    Public Event Mouse_Wheel(ByVal ptLocat As Point, ByVal Direction As Wheel_Direction)

    Public Sub New()
        MouseHookDelegate = New MouseProcDelegate(AddressOf MouseProc)
        MouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
    End Sub

    Private Function MouseProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer
        If (nCode = HC_ACTION) Then
            Select Case wParam
                Case WM_MOUSEMOVE
                    RaiseEvent Mouse_Move(lParam.pt)
                Case WM_LBUTTONDOWN
                    RaiseEvent Mouse_Left_Down(lParam.pt)
                Case WM_LBUTTONUP
                    RaiseEvent Mouse_Left_Up(lParam.pt)
                Case WM_LBUTTONDBLCLK
                    RaiseEvent Mouse_Left_DoubleClick(lParam.pt)
                Case WM_RBUTTONDOWN
                    RaiseEvent Mouse_Right_Down(lParam.pt)
                Case WM_RBUTTONUP
                    RaiseEvent Mouse_Right_Up(lParam.pt)
                Case WM_RBUTTONDBLCLK
                    RaiseEvent Mouse_Right_DoubleClick(lParam.pt)
                Case WM_MBUTTONDOWN
                    RaiseEvent Mouse_Middle_Down(lParam.pt)
                Case WM_MBUTTONUP
                    RaiseEvent Mouse_Middle_Up(lParam.pt)
                Case WM_MBUTTONDBLCLK
                    RaiseEvent Mouse_Middle_DoubleClick(lParam.pt)
                Case WM_MOUSEWHEEL
                    Dim wDirection As Wheel_Direction
                    If lParam.mouseData < 0 Then
                        wDirection = Wheel_Direction.WheelDown
                    Else
                        wDirection = Wheel_Direction.WheelUp
                    End If
                    RaiseEvent Mouse_Wheel(lParam.pt, wDirection)
            End Select
        End If
        Return CallNextHookEx(MouseHook, nCode, wParam, lParam)
    End Function

    Protected Overrides Sub Finalize()
        UnhookWindowsHookEx(MouseHook)
        MyBase.Finalize()
    End Sub
End Class

'Class para loggear todo, hecha por mí:

Public Class Escribir
    Public Shared Sub Escribir(ByVal texto As String, ByVal PathArchivo As String)
        Dim strStreamWriter As StreamWriter
        Dim sRenglon As String = Nothing
        Dim strStreamW As Stream = Nothing
        strStreamWriter = Nothing
        Dim ContenidoArchivo As String = Nothing

        Try

            Dim PathCarpeta As String = Replace(PathArchivo, PathArchivo.Substring(PathArchivo.LastIndexOf("/") + 1), "")

            If Directory.Exists(PathCarpeta) = False Then
                Directory.CreateDirectory(PathCarpeta)
            End If

            If File.Exists(PathArchivo) Then
                strStreamW = File.Open(PathArchivo, FileMode.Append)
            Else
                strStreamW = File.Create(PathArchivo)
            End If

            strStreamWriter = New StreamWriter(strStreamW, System.Text.Encoding.Default)

            strStreamWriter.Write(texto & Environment.NewLine)

            strStreamWriter.Close()

        Catch ex As Exception
            MsgBox("Error al Guardar la ingormacion en el archivo. " & ex.ToString, MsgBoxStyle.Critical, Application.ProductName)
            strStreamWriter.Close()
        End Try
    End Sub
End Class

'Ejemplos: (Hechos por Sim0n y editados por mí)

'Dentro de un Form:

'Private WithEvents mHook As New MouseHook

'Private Sub mHook_Mouse_Left_DoubleClick(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Left_DoubleClick
'    Escribir.Escribir("Mouse Left Double Click At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Left_Down(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Left_Down
'    Escribir.Escribir("Mouse Left Down At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Left_Up(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Left_Up
'    Escribir.Escribir("Mouse Left Up At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Middle_DoubleClick(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Middle_DoubleClick
'    Escribir.Escribir("Mouse Middle Double Click At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Middle_Down(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Middle_Down
'    Escribir.Escribir("Mouse Middle Down At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Middle_Up(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Middle_Up
'    Escribir.Escribir("Mouse Middle Up At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Move(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Move
'    ''Escribir.Escribir("Mouse Move At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Right_DoubleClick(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Right_DoubleClick
'    Escribir.Escribir("Mouse Right Double Click At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Right_Down(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Right_Down
'    Escribir.Escribir("Mouse Right Down At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Right_Up(ByVal ptLocat As System.Drawing.Point) Handles mHook.Mouse_Right_Up
'    Escribir.Escribir("Mouse Right Up At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'Private Sub mHook_Mouse_Wheel(ByVal ptLocat As System.Drawing.Point, ByVal Direction As MouseHook.Wheel_Direction) Handles mHook.Mouse_Wheel
'    Escribir.Escribir("Mouse Scroll: " & Direction.ToString & " At: (" & ptLocat.X & "," & ptLocat.Y & ")", "C:/Carpeta/Archivo.txt")
'End Sub

'================================================

'Class Low KeyBoard Hook Hecha por Sim0n: http://sim0n.wordpress.com/2009/03/28/vbnet-keyboard-hook-class/

Public Class KeyboardHook

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As KBDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    <Flags()> _
    Private Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    Public Shared Event KeyDown(ByVal Key As Keys)
    Public Shared Event KeyUp(ByVal Key As Keys)

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const HC_ACTION As Integer = 0
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101
    Private Const WM_SYSKEYDOWN = &H104
    Private Const WM_SYSKEYUP = &H105

    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    Private HHookID As IntPtr = IntPtr.Zero

    Private Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        If (nCode = HC_ACTION) Then
            Dim struct As KBDLLHOOKSTRUCT
            Select Case wParam
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    RaiseEvent KeyDown(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
                Case WM_KEYUP, WM_SYSKEYUP
                    RaiseEvent KeyUp(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
            End Select
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    Public Sub New()
        HHookID = SetWindowsHookEx(WH_KEYBOARD_LL, KBDLLHookProcDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
        If HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
        MyBase.Finalize()
    End Sub

End Class

'Ejemplos (Hechos por Elektro H@cker): 

'Dentro de un Form:

'Public WithEvents KeysHook As New KeyboardHook

'Dim Auto_Backspace_Key As Boolean = True
'Dim Auto_Enter_Key As Boolean = True
'Dim Auto_Tab_Key As Boolean = True
'Dim No_F_Keys As Boolean = False

'Private Sub KeysHook_KeyDown(ByVal Key As Keys) Handles KeysHook.KeyDown

'    Select Case Control.ModifierKeys

'        Case 393216 ' Alt-GR + Key

'            Select Case Key
'                Case Keys.D1 : Key_Listener("|")
'                Case Keys.D2 : Key_Listener("@")
'                Case Keys.D3 : Key_Listener("#")
'                Case Keys.D4 : Key_Listener("~")
'                Case Keys.D5 : Key_Listener("€")
'                Case Keys.D6 : Key_Listener("¬")
'                Case Keys.E : Key_Listener("€")
'                Case Keys.Oem1 : Key_Listener("[")
'                Case Keys.Oem5 : Key_Listener("\")
'                Case Keys.Oem7 : Key_Listener("{")
'                Case Keys.Oemplus : Key_Listener("]")
'                Case Keys.OemQuestion : Key_Listener("}")
'                Case Else : Key_Listener("")
'            End Select

'        Case 65536 ' LShift/RShift + Key

'            Select Case Key
'                Case Keys.D0 : Key_Listener("=")
'                Case Keys.D1 : Key_Listener("!")
'                Case Keys.D2 : Key_Listener("""")
'                Case Keys.D3 : Key_Listener("·")
'                Case Keys.D4 : Key_Listener("$")
'                Case Keys.D5 : Key_Listener("%")
'                Case Keys.D6 : Key_Listener("&")
'                Case Keys.D7 : Key_Listener("/")
'                Case Keys.D8 : Key_Listener("(")
'                Case Keys.D9 : Key_Listener(")")
'                Case Keys.Oem1 : Key_Listener("^")
'                Case Keys.Oem5 : Key_Listener("ª")
'                Case Keys.Oem6 : Key_Listener("¿")
'                Case Keys.Oem7 : Key_Listener("¨")
'                Case Keys.OemBackslash : Key_Listener(">")
'                Case Keys.Oemcomma : Key_Listener(";")
'                Case Keys.OemMinus : Key_Listener("_")
'                Case Keys.OemOpenBrackets : Key_Listener("?")
'                Case Keys.OemPeriod : Key_Listener(":")
'                Case Keys.Oemplus : Key_Listener("*")
'                Case Keys.OemQuestion : Key_Listener("Ç")
'                Case Keys.Oemtilde : Key_Listener("Ñ")
'                Case Else : Key_Listener("")
'            End Select

'        Case Else

'            If Key.ToString.Length = 1 Then ' Single alpha key

'                If Control.IsKeyLocked(Keys.CapsLock) Or Control.ModifierKeys = Keys.Shift Then
'                    Key_Listener(Key.ToString.ToUpper)
'                Else
'                    Key_Listener(Key.ToString.ToLower)
'                End If

'            Else

'                Select Case Key ' Single special key 
'                    Case Keys.Add : Key_Listener("+")
'                    Case Keys.Back : Key_Listener("{BackSpace}")
'                    Case Keys.D0 : Key_Listener("0")
'                    Case Keys.D1 : Key_Listener("1")
'                    Case Keys.D2 : Key_Listener("2")
'                    Case Keys.D3 : Key_Listener("3")
'                    Case Keys.D4 : Key_Listener("4")
'                    Case Keys.D5 : Key_Listener("5")
'                    Case Keys.D6 : Key_Listener("6")
'                    Case Keys.D7 : Key_Listener("7")
'                    Case Keys.D8 : Key_Listener("8")
'                    Case Keys.D9 : Key_Listener("9")
'                    Case Keys.Decimal : Key_Listener(".")
'                    Case Keys.Delete : Key_Listener("{Supr}")
'                    Case Keys.Divide : Key_Listener("/")
'                    Case Keys.End : Key_Listener("{End}")
'                    Case Keys.Enter : Key_Listener("{Enter}")
'                    Case Keys.F1 : Key_Listener("{F1}")
'                    Case Keys.F10 : Key_Listener("{F10}")
'                    Case Keys.F11 : Key_Listener("{F11}")
'                    Case Keys.F12 : Key_Listener("{F12}")
'                    Case Keys.F2 : Key_Listener("{F2}")
'                    Case Keys.F3 : Key_Listener("{F3}")
'                    Case Keys.F4 : Key_Listener("{F4}")
'                    Case Keys.F5 : Key_Listener("{F5}")
'                    Case Keys.F6 : Key_Listener("{F6}")
'                    Case Keys.F7 : Key_Listener("{F7}")
'                    Case Keys.F8 : Key_Listener("{F8}")
'                    Case Keys.F9 : Key_Listener("{F9}")
'                    Case Keys.Home : Key_Listener("{Home}")
'                    Case Keys.Insert : Key_Listener("{Insert}")
'                    Case Keys.Multiply : Key_Listener("*")
'                    Case Keys.NumPad0 : Key_Listener("0")
'                    Case Keys.NumPad1 : Key_Listener("1")
'                    Case Keys.NumPad2 : Key_Listener("2")
'                    Case Keys.NumPad3 : Key_Listener("3")
'                    Case Keys.NumPad4 : Key_Listener("4")
'                    Case Keys.NumPad5 : Key_Listener("5")
'                    Case Keys.NumPad6 : Key_Listener("6")
'                    Case Keys.NumPad7 : Key_Listener("7")
'                    Case Keys.NumPad8 : Key_Listener("8")
'                    Case Keys.NumPad9 : Key_Listener("9")
'                    Case Keys.Oem1 : Key_Listener("`")
'                    Case Keys.Oem5 : Key_Listener("º")
'                    Case Keys.Oem6 : Key_Listener("¡")
'                    Case Keys.Oem7 : Key_Listener("´")
'                    Case Keys.OemBackslash : Key_Listener("<")
'                    Case Keys.Oemcomma : Key_Listener(",")
'                    Case Keys.OemMinus : Key_Listener(".")
'                    Case Keys.OemOpenBrackets : Key_Listener("'")
'                    Case Keys.OemPeriod : Key_Listener("-")
'                    Case Keys.Oemplus : Key_Listener("+")
'                    Case Keys.OemQuestion : Key_Listener("ç")
'                    Case Keys.Oemtilde : Key_Listener("ñ")
'                    Case Keys.PageDown : Key_Listener("{AvPag}")
'                    Case Keys.PageUp : Key_Listener("{RePag}")
'                    Case Keys.Space : Key_Listener(" ")
'                    Case Keys.Subtract : Key_Listener("-")
'                    Case Keys.Tab : Key_Listener("{Tabulation}")
'                    Case Else : Key_Listener("")
'                End Select

'            End If

'    End Select

'End Sub

'Public Sub Key_Listener(ByVal key As String)

'    If Auto_Backspace_Key AndAlso key = "{BackSpace}" Then ' Delete character
'        RichTextBox1.Text = RichTextBox1.Text.Substring(0, RichTextBox1.Text.Length - 1)
'    ElseIf Auto_Enter_Key AndAlso key = "{Enter}" Then ' Insert new line
'        RichTextBox1.Text += ControlChars.NewLine
'    ElseIf Auto_Tab_Key AndAlso key = "{Tabulation}" Then ' Insert Tabulation
'        RichTextBox1.Text += ControlChars.Tab
'    ElseIf No_F_Keys AndAlso key.StartsWith("{F") Then ' Ommit F Keys
'    Else ' Print the character
'        RichTextBox1.Text += key
'    End If

'End Sub

'=======================================

'INI_Manager hecho por Elektro: https://foro.elhacker.net/net/libreria_de_snippets_posteen_aqui_sus_snippets-t378770.0.html;msg1860295#msg1860295

Public Class INI_Manager

    ''' <summary>
    ''' The INI File Location.
    ''' </summary>
    Public Shared INI_File As String = IO.Path.Combine(Application.StartupPath, Process.GetCurrentProcess().ProcessName & ".ini")

    ''' <summary>
    ''' Set a value.
    ''' </summary>
    ''' <param name="File">The INI file location</param>
    ''' <param name="ValueName">The value name</param>
    ''' <param name="Value">The value data</param>
    Public Shared Sub Set_Value(ByVal File As String, ByVal ValueName As String, ByVal Value As String)

        Try

            If Not IO.File.Exists(File) Then ' Create a new INI File with "Key=Value""

                My.Computer.FileSystem.WriteAllText(File, ValueName & "=" & Value, False)
                Exit Sub

            Else ' Search line by line in the INI file for the "Key"

                Dim Line_Number As Int64 = 0
                Dim strArray() As String = IO.File.ReadAllLines(File)

                For Each line In strArray
                    If line.ToLower.StartsWith(ValueName.ToLower & "=") Then
                        strArray(Line_Number) = ValueName & "=" & Value
                        IO.File.WriteAllLines(File, strArray) ' Replace "value"
                        Exit Sub
                    End If
                    Line_Number += 1
                Next

                Application.DoEvents()

                My.Computer.FileSystem.WriteAllText(File, vbNewLine & ValueName & "=" & Value, True) ' Key don't exist, then create the new "Key=Value"

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' Load a value.
    ''' </summary>
    ''' <param name="File">The INI file location</param>
    ''' <param name="ValueName">The value name</param>
    ''' <returns>The value itself</returns>
    Public Shared Function Load_Value(ByVal File As String, ByVal ValueName As String) As Object

        If Not IO.File.Exists(File) Then

            Throw New Exception(File & " not found.") ' INI File not found.
            Return Nothing

        Else

            For Each line In IO.File.ReadAllLines(File)
                If line.ToLower.StartsWith(ValueName.ToLower & "=") Then Return line.Split("=").Last
            Next

            Application.DoEvents()

            Throw New Exception("Key: " & """" & ValueName & """" & " not found.") ' Key not found.
            Return Nothing

        End If

    End Function

    ''' <summary>
    ''' Delete a key.
    ''' </summary>
    ''' <param name="File">The INI file location</param>
    ''' <param name="ValueName">The value name</param>
    Public Shared Sub Delete_Value(ByVal File As String, ByVal ValueName As String)

        If Not IO.File.Exists(File) Then

            Throw New Exception(File & " not found.") ' INI File not found.
            Exit Sub

        Else

            Try

                Dim Line_Number As Int64 = 0
                Dim strArray() As String = IO.File.ReadAllLines(File)

                For Each line In strArray
                    If line.ToLower.StartsWith(ValueName.ToLower & "=") Then
                        strArray(Line_Number) = Nothing
                        Exit For
                    End If
                    Line_Number += 1
                Next

                Array.Copy(strArray, Line_Number + 1, strArray, Line_Number, UBound(strArray) - Line_Number)
                ReDim Preserve strArray(UBound(strArray) - 1)

                My.Computer.FileSystem.WriteAllText(File, String.Join(vbNewLine, strArray), False)

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End If

    End Sub

    ''' <summary>
    ''' Sorts the entire INI File.
    ''' </summary>
    ''' <param name="File">The INI file location</param>
    Public Shared Sub Sort_Values(ByVal File As String)

        If Not IO.File.Exists(File) Then

            Throw New Exception(File & " not found.") ' INI File not found.
            Exit Sub

        Else

            Try

                Dim Line_Number As Int64 = 0
                Dim strArray() As String = IO.File.ReadAllLines(File)
                Dim TempList As New List(Of String)

                For Each line As String In strArray
                    If line <> "" Then TempList.Add(strArray(Line_Number))
                    Line_Number += 1
                Next

                TempList.Sort()
                IO.File.WriteAllLines(File, TempList)

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End If

    End Sub

End Class

'Ejemplos:

' INI_Manager.Set_Value(".\Test.ini", "TextValue", TextBox1.Text) ' Save
' TextBox1.Text = INI_Manager.Load_Value(".\Test.ini", "TextValue") ' Load
' INI_Manager.Delete_Value(".\Test.ini", "TextValue") ' Delete
' INI_Manager.Sort_Values(".\Test.ini") ' Sort INI File

'==========================

'Class para dar estilo a un RichTextBox, hecha por mí y mejorada por Elektro:

'https://foro.elhacker.net/net/libreria_de_snippets_posteen_aqui_sus_snippets-t378770.0.html;msg1867833#msg1867833 Respuesta #246
'https://foro.elhacker.net/net/libreria_de_snippets_posteen_aqui_sus_snippets-t378770.0.html;msg1867839#msg1867839 Respuesta #247 (Mejorado por Elektro)

Public Class RichTextLabel

    Public Shared Sub Add_Text_With_Color(ByVal richTextBox As RichTextBox, _
                                              ByVal text As String, _
                                              ByVal color As Color, _
                                              Optional ByVal font As Font = Nothing)

        'richTextBox.Enabled = False
        richTextBox.ReadOnly = True
        richTextBox.BorderStyle = BorderStyle.None
        richTextBox.ScrollBars = RichTextBoxScrollBars.None

        Dim index As Int32 = richTextBox.TextLength
        richTextBox.AppendText(text)
        richTextBox.SelectionStart = index
        richTextBox.SelectionLength = richTextBox.TextLength - index
        richTextBox.SelectionColor = color
        If font IsNot Nothing Then richTextBox.SelectionFont = font
        'richTextBox.BackColor = Drawing.Color.White

    End Sub

End Class

'Ejemplos:

'RichTextLabel.Add_Text_With_Color(RichTextBox1, "algo de texto con Arial al 12", RichTextBox1.ForeColor, New Font("Arial", 12, FontStyle.Bold))
'RichTextLabel.Add_Text_With_Color(RichTextBox1, " ROOOJOOORL xD", Color.Red)
'RichTextLabel.Add_Text_With_Color(RichTextBox1, Environment.NewLine & "nueva linea y algo de texto", Color.Black)

'====================

'Enviar un Email:

Public Class Emaileitor
    Shared Function SendEmail(ByVal Recipients As List(Of String), _
                  ByVal FromAddress As String, _
                  ByVal Subject As String, _
                  ByVal Body As String, _
                  ByVal UserName As String, _
                  ByVal Password As String, _
                  Optional ByVal Server As String = "smtp.live.com", _
                  Optional ByVal Port As Integer = 587, _
                  Optional ByVal Attachments As List(Of String) = Nothing) As String
        Dim Email As New MailMessage()
        Try
            Dim SMTPServer As New SmtpClient
            For Each Attachment As String In Attachments
                Email.Attachments.Add(New Attachment(Attachment))
            Next
            Email.From = New MailAddress(FromAddress)
            For Each Recipient As String In Recipients
                Email.To.Add(Recipient)
            Next
            Email.Subject = Subject
            Email.Body = Body
            SMTPServer.Host = Server
            SMTPServer.Port = Port
            SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
            SMTPServer.EnableSsl = True
            SMTPServer.Send(Email)
            Email.Dispose()
            Return "Email to " & Recipients(0) & " from " & FromAddress & " was sent."
        Catch ex As SmtpException
            Email.Dispose()
            Return "Sending Email Failed. Smtp Error. " & ex.Message
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            Return "Sending Email Failed. Check Port Number."
        Catch Ex As InvalidOperationException
            Email.Dispose()
            Return "Sending Email Failed. Check Port Number."
        End Try
    End Function

End Class

'Ejemplos:

'Dim Recipients As New List(Of String)
'Recipients.Add("<Email1>", "<Email2>")
'Dim FromEmailAddress As String = Recipients(0)
'Dim Subject As String = "<Asunto del Mensaje>"
'Dim Body As String = "<Cuerpo del Mensaje>"
'Dim UserName As String = "<Tu Email>"
'Dim Password As String = "<Contraseña>"
'Dim Port As Integer = 587
'Dim Server As String = "smtp.live.com" 'Puede ser también smtp.gmail.com
'Dim Attachments As New List(Of String)
'Emaileitor.SendEmail(Recipients, FromEmailAddress, Subject, Body, UserName, Password, Server, Port, Attachments)

'=========================

'Thread.Sleep pero sin bloquear la app (Me lo envió BM4):

Public Class Esperar
    Private Sub Esperar(ByVal interval As Integer)
        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
            Application.DoEvents()
        Loop
        sw.Stop()
    End Sub
End Class

'Ejemplos:

'Hace falta un ejemplo? Esperar(Tiempo en milisegundos)

'===================

'Hacer Click:

Public Class Clickar
    Private Declare Function SetCursorPos Lib "user32.dll" ( _
ByVal X As Int32, _
ByVal Y As Int32 _
) As Boolean
    Private Declare Sub mouse_event Lib "user32.dll" ( _
    ByVal dwFlags As Int32, _
    ByVal dx As Int32, _
    ByVal dy As Int32, _
    ByVal cButtons As Int32, _
    ByVal dwExtraInfo As Int32 _
    )
    Private Const MOUSEEVENTF_ABSOLUTE = &H8000 ' absolute move
    Private Const MOUSEEVENTF_LEFTDOWN = &H2 ' left button down
    Private Const MOUSEEVENTF_LEFTUP = &H4 ' left button up
    Private Const MOUSEEVENTF_MOVE = &H1 ' mouse move
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Private Const MOUSEEVENTF_MIDDLEUP = &H40
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8
    Private Const MOUSEEVENTF_RIGHTUP = &H10

    Public Shared Sub MouseLeftClick(ByVal PosX As Long, ByVal PosY As Long)
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, PosX, PosY, 0, 0)
        Call mouse_event(MOUSEEVENTF_LEFTUP, PosX, PosY, 0, 0)
    End Sub

    Public Shared Sub MouseRightClick(ByVal PosX As Long, ByVal PosY As Long)
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, PosX, PosY, 0, 0)
        Call mouse_event(MOUSEEVENTF_RIGHTUP, PosX, PosY, 0, 0)
    End Sub
End Class

'Ejemplos:

'Clickar.SetCursorPos(500, 600) 'Moverá el raton a 500, 600
'Clickar.MouseRightClick(500, 600) 'Hará click derecho en 500, 600
'Clickar.MouseLeftClick(500, 600) 'Hará click izquierdo en 500, 600

'=======================

Public Class Extraer
    Public Shared Sub Extraer(ByVal ZipAExtraer As String, ByVal DirectorioExtraccion As String)
        Try

            Using zip1 As ZipFile = ZipFile.Read(ZipAExtraer)
                Dim e As ZipEntry
                For Each e In zip1
                    e.Extract(DirectorioExtraccion, ExtractExistingFileAction.OverwriteSilently)
                Next
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class

'Ejemplo:

'Extractor sacado de por ahí y adaptado por mí
'PD: Solo funciona con .Zips, creo xD
' Ejemplo: Extraer.Extraer("File.zip", ".\Directorio A Extraer\SubDirectorio")

Public Class Comprimir

    Public Shared Sub Comprimir(ByVal NombreDirectorio, ByVal NombreGuardar)
        Using zip As ZipFile = New ZipFile()
            zip.AddDirectory(NombreDirectorio)
            zip.Save(NombreGuardar)
        End Using
    End Sub

End Class

'Compresor sacado de por ahí y adaptado por mí
' Ejemplo: Comprimir.Comprimir(".\Directorio A Comprimir\SubDirectorio", "File Compreso.zip")

'========================

'Leer línea customizada de un TXT

Public Class Leer
    Public Shared Sub Leer(archivo As String)
        Dim lines As String() = IO.File.ReadAllLines(archivo)
    End Sub
End Class

'Ejemplo: lines(1) 'esto lee la línea 2 del archivo.txt

'=============

Public Class Finder
    Shared Function Find_String_Occurrences(ByVal Input_String As String, ByVal Search_String As String) As Integer

        Dim Input_String_Pos As Int32
        Dim Input_String_Count As Int32

        Do
            Input_String_Pos = Input_String.IndexOf(Search_String, Input_String_Pos)
            If Input_String_Pos <> -1 Then
                Input_String_Count += 1
                Input_String_Pos += Search_String.Length
            End If
        Loop Until Input_String_Pos = -1

        Return Input_String_Count

    End Function
End Class

'Ejemplo:

' MsgBox(Find_String_Occurrences("Hello World", "o"))            ' Result: 2
' MsgBox(Find_String_Occurrences("Hello me Hello you", "Hello")) ' Result: 2

'=========

'Updater creado por Ikillnukes
' Ejemplos: Updater.Comprobar("https://dl.dropboxusercontent.com/s/2iin21gf8g629j9/upt.txt?dl=1", ".\Temp\", "1")
'La url puede ser de cualquier tipo yo recomiendo que uséis Dropbox, puesto que es directo y la url no sufre cambios.
'El directorio puede ser cualquier sitio
'El texto es la cadena que se va a comprobar, en caso de que no sea la misma que la del texto descargado previamente en Updatear, se va a llevar a acabo la funcion Updatear

Public Class Updater

    Public Shared Sub Comprobar(ByVal url As String, ByVal directorio As String, ByVal texto As String)
        Dim patha As String = directorio & "upt.txt"
        Dim patha2 As String = directorio & "Update.zip"
        Dim patha3 As String = directorio & "upt.exe"

        If File.Exists(patha) Then
            File.Delete(patha)
        End If

        If File.Exists(patha2) Then
            File.Delete(patha2)
        End If

        If File.Exists(patha3) Then
            File.Delete(patha3)
        End If

        If Not File.Exists(patha) Then
            My.Computer.Network.DownloadFile(
        url,
        patha)
        End If

        If File.Exists(patha) Then

            Dim lines As String() = File.ReadAllLines(patha)

            If Not lines(0) = texto Then
                If MsgBox("¡Atención! Su aplicación está desactualizada." & vbCrLf & "Pulse ""Sí"" para continuar con la instalación. O ""No"" para descartar cambios.", MsgBoxStyle.YesNo, "¡Atención! Su app está desactualizada...") = MsgBoxResult.Yes Then
                    My.Computer.Network.DownloadFile(
            lines(1),
            patha2)
                    Extraer.Extraer(patha2, directorio)
                    Dim psi As New ProcessStartInfo()
                    psi.UseShellExecute = True
                    psi.FileName = patha3
                    Process.Start(psi)
                    Application.Exit()
                End If
            End If

        End If
    End Sub

End Class



'Dim url As String = "https://dl.dropboxusercontent.com/s/2iin21gf8g629j9/upt.txt?dl=1" 'Esta es la Url de donde va a comprobarse todo
'    Dim texto As String = INI_Manager.Load_Value(".\Test.ini", "AppVer") 'Aquí está la cadena de texto que se chekea

'Sub Updatear()
'Updater.Comprobar(url, ".\Temp\", texto)
'End Sub

'Dim WithEvents temer As New System.Windows.Forms.Timer With {.Interval = 15000, .Enabled = True}

'Private Sub Temer_Start(sender As Object, e As EventArgs) Handles temer.Tick
'Updatear() 'Aquí se chekea cada 15 secs esa función
'End Sub

Public Class Randomize
    Public Shared Function Random_String(ByVal Length As Int32, _
                                   Optional ByVal Characters As String = _
                                   "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" _
                                  ) As String

        Select Case Length

            Case Is < 1 ' Is 0 or negative
                Throw New Exception("Length must be greater than 0")

            Case Else ' Is greater than 0

                Dim str As String = String.Empty
                Dim rand As New Random, rand_length As Int32 = Characters.Length

                Do Until str.Length = Length
                    str &= Characters.Substring(rand.Next(0, rand_length), 1)
                Loop

                Return str
        End Select

    End Function
End Class

'Ejemplos:

'MsgBox(Randomize.Random_String(8))
