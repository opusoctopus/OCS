'Serial Port Interfacing with VB.net 2010 Express Edition
'Copyright (C) 2010 Richard Myrick T. Arellaga
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see .


Imports System
Imports System.ComponentModel
Imports System.Threading
Imports System.IO.Ports
Public Class FormTimbanganCSKD
    Dim myPort As Array 'COM Ports detected on the system will be stored here
    Delegate Sub SetTextCallback(ByVal [text] As String) 'Added to prevent threading errors during receiveing of data

    Private Sub FormTimbanganCSKD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'When our form loads, auto detect all serial ports in the system and populate the cmbPort Combo box.
        myPort = IO.Ports.SerialPort.GetPortNames() 'Get all com ports available
        cmbBaud.Items.Add(9600) 'Populate the cmbBaud Combo box to common baud rates used
        cmbBaud.Items.Add(19200)
        cmbBaud.Items.Add(38400)
        cmbBaud.Items.Add(57600)
        cmbBaud.Items.Add(115200)

        For i = 0 To UBound(myPort)
            cmbPort.Items.Add(myPort(i))
        Next
        cmbPort.Text = cmbPort.Items.Item(0) 'Set cmbPort text to the first COM port detected
        cmbBaud.Text = cmbBaud.Items.Item(0) 'Set cmbBaud text to the first Baud rate on the list

        Button7.Enabled = False 'Initially Disconnect Button is Disabled
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        comPort.PortName = cmbPort.Text 'Set SerialPort1 to the selected COM port at startup
        comPort.BaudRate = cmbBaud.Text 'Set Baud rate to the selected value on

        'Other Serial Port Property
        comPort.Parity = IO.Ports.Parity.None
        comPort.StopBits = IO.Ports.StopBits.One
        comPort.DataBits = 8 'Open our serial port
        comPort.Open()

        Button6.Enabled = False 'Disable Connect button
        Button7.Enabled = True 'and Enable Disconnect button
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        comPort.Close() 'Close our Serial Port

        Button6.Enabled = True
        Button7.Enabled = False
    End Sub

    Private Sub comPort_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles comPort.DataReceived
        ReceivedText(comPort.ReadExisting()) 'Automatically called every time a data is received at the serialPort
    End Sub

    Private Sub ReceivedText(ByVal [text] As String)
        'compares the ID of the creating Thread to the ID of the calling Thread
        If Me.RichTextBox1.InvokeRequired Then
            Dim x As New SetTextCallback(AddressOf ReceivedText)
            Me.Invoke(x, New Object() {(text)})
        Else
            Me.RichTextBox1.Text &= [text]
            text = RichTextBox1.Text.Substring(12)
        End If
    End Sub

    Private Sub cmbPort_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPort.SelectedIndexChanged
        If comPort.IsOpen = False Then
            comPort.PortName = cmbPort.Text 'pop a message box to user if he is changing ports
        Else 'without disconnecting first.
            MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub cmbBaud_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBaud.SelectedIndexChanged
        If comPort.IsOpen = False Then
            comPort.BaudRate = cmbBaud.Text 'pop a message box to user if he is changing baud rate
        Else 'without disconnecting first.
            MsgBox("Valid only if port is Closed", vbCritical)
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        comPort.Write(TextBox3.Text & vbCr)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        comPort.Write(TextBox3.Text & vbCr)
    End Sub
End Class