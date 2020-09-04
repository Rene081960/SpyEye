Imports System.Threading
Imports System.IO.Ports

Public Class Form1

    Dim lastx, lasty As Int16
    Dim closeme As Boolean = False

    Dim checkMouse As Boolean = True


    Dim ArduinoPort As SerialPort = Nothing


    Private Sub ShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowToolStripMenuItem.Click
        If ShowToolStripMenuItem.Text = "Show" Then
            Me.Show()
            Me.WindowState = FormWindowState.Normal
            ShowToolStripMenuItem.Text = "Hide"
        Else
            ShowToolStripMenuItem.Text = "Show"

            Me.WindowState = FormWindowState.Minimized
            Me.Hide()
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        closeme = True
        Me.Close()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ShowToolStripMenuItem.Text = "Show"
        Me.WindowState = FormWindowState.Minimized
        Me.Hide()
        If Not closeme Then e.Cancel = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim x, y, w, h As Int16
        Dim strsend = ""
        w = Screen.PrimaryScreen.Bounds.Width
        h = Screen.PrimaryScreen.Bounds.Height
        x = Cursor.Position.X
        y = Cursor.Position.Y
        Label2.Text = "mouse at:" & x.ToString & "," & y.ToString

        x = Math.Round(10 * (x / w))
        y = Math.Round(10 * (y / h))
        If (x > 9) Then x = 9
        If (y > 9) Then y = 9

        If Not ArduinoPort.IsOpen Then
            Label1.Text = "Connection lost!"
            TextBox1.AppendText(vbNewLine & "ERROR: Connection lost!" & vbNewLine)
            Timer1.Enabled = False
            Exit Sub
        End If

        If (lastx <> x Or lasty <> y) Then
            strsend = "xy" + x.ToString + y.ToString
            lastx = x
            lasty = y
            ArduinoPort.WriteLine(strsend)
        End If

        If MouseButtons = MouseButtons.None Then checkMouse = True

        If checkMouse And MouseButtons = MouseButtons.Left Then
            strsend = "by" + x.ToString + y.ToString
            checkMouse = False
            ArduinoPort.WriteLine(strsend)
        End If


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, MyBase.Load
        Label1.Text = "Not connected"
        If ArduinoPort IsNot Nothing Then
            If ArduinoPort.IsOpen Then ArduinoPort.Close()
        End If
        TextBox1.AppendText(vbNewLine)
        If Not FindArduino().Equals("") Then
            Timer1.Enabled = True
        End If
    End Sub





    Private Function FindArduino() As String
        Timer1.Enabled = False
        TextBox1.AppendText("Searching for Arduino." & vbNewLine)
        ' see https://docs.microsoft.com/en-us/dotnet/visual-basic/developing-apps/programming/computer-resources/how-to-receive-strings-from-serial-ports
        Dim AvailablePorts() As String = SerialPort.GetPortNames()
        Dim port As String = ""

        If My.Computer.Ports.SerialPortNames.Count < 1 Then
            TextBox1.AppendText("No serial ports available." & vbNewLine)
            Return ""
        End If
        For Each port In AvailablePorts
            ' try to connect to serial port sp

            Try
                ArduinoPort = New SerialPort()
                TextBox1.AppendText("Checking port:" & port & vbNewLine)
                ArduinoPort.DtrEnable = False 'this should prevent the arduino to reset on connecting!
                ArduinoPort.PortName = port
                ArduinoPort.ReadTimeout = 500
                ArduinoPort.BaudRate = 57600
                ArduinoPort.Parity = Parity.None
                ArduinoPort.DataBits = 8
                ArduinoPort.StopBits = StopBits.One

                Try
                    ArduinoPort.Open()
                Catch
                    Continue For
                End Try


                ' the arduino will reset when we connect, so give it some time to boot
                Thread.Sleep(2500)

                ' write V to serial port
                ArduinoPort.WriteLine("V")
                ' read the response
                Dim response As String = ArduinoPort.ReadLine()
                If response.Contains("Arduino-eyes") Then
                    ' we found the arduino project!
                    Label1.Text = "Connected to " & port
                    TextBox1.AppendText("Found Arduino on port:" & port & vbNewLine)
                    'ArduinoPort.Close()
                    Return port
                Else
                    ArduinoPort.Close()
                End If

            Catch ex As TimeoutException
                'TextBox1.AppendText("Error: Serial Port read timed out." & vbNewLine)
            End Try

        Next
        Return port
    End Function
End Class
