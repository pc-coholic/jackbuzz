'          _\|/_              
'          (o o)                         
' +-----oOO-{_}-OOo--------------------------------------------------------+
' | JackBuzz 0.1                                                           |
' +------------------------------------------------------------------------+
' | Translates the input of Buzz-Buzzers into YDKJ-Keyboard-Input          |
' +------------------------------------------------------------------------+
' | Author: Martin Gross <martin@pc-coholic.de>                            |
' +------------------------------------------------------------------------+
' | Created @ 04.01.2010                                                   |
' +------------------------------------------------------------------------+
' | This program is free software; you can redistribute it and/or modify   |
' | it under the terms of the GNU General Public License as published by   |
' | the Free Software Foundation; version 3 of the License.                |
' |                                                                        |
' | This program is distributed in the hope that it will be useful,        |
' | but WITHOUT ANY WARRANTY; without even the implied warranty of         |
' | MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          |
' | GNU General Public License for more details.                           |
' +------------------------------------------------------------------------+
' | Changes                                                                |
' | Date       | Who?   | What has been done?                              |
' | 04.01.2010 | Martin | Project has been created.                        |
' +------------------------------------------------------------------------+

Public Class Form1
    Declare Function BuzzerOpen Lib "Buzzer.dll" () As Integer
    Declare Function BuzzerClose Lib "Buzzer.dll" () As Integer 'was: Void
    Declare Function BuzzerSetLEDs Lib "Buzzer.dll" (ByVal buffer As Byte(), ByVal buzzers As Integer) As Integer
    Declare Function BuzzerGetButtons Lib "Buzzer.dll" (ByVal buttons As Byte(), ByVal buzzers As Integer) As Integer
    Dim varBuzzers

    Private Sub Form1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MsgBox(BuzzerClose)
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Do
            varBuzzers = BuzzerOpen()
            If varBuzzers = 0 Then
                If MsgBox("No Buzz-buzzers have been found. Please check cabling and retry again.", MsgBoxStyle.RetryCancel, "No buzzers found") = MsgBoxResult.Cancel Then
                    Application.Exit()
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
    End Sub

    Private Sub btnPlayer1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayer1.Click
        SetLEDon(0)
    End Sub

    Private Sub btnPlayer2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayer2.Click
        SetLEDon(1)
    End Sub

    Private Sub btnPlayer3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayer3.Click
        SetLEDon(2)
    End Sub

    Private Sub btnPlayerAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlayerAdmin.Click
        SetLEDon(3)
    End Sub

    Private Sub tmrBlink_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrBlink.Tick
        SetLEDoff()
    End Sub

    Private Function SetLEDon(ByVal WhichControler)
        Dim varBuffer(3) As Byte
        varBuffer(0) = 0
        varBuffer(1) = 0
        varBuffer(2) = 0
        varBuffer(3) = 0

        varBuffer(WhichControler) = 1

        BuzzerSetLEDs(varBuffer, varBuzzers)

        btnPlayer1.Enabled = False
        btnPlayer2.Enabled = False
        btnPlayer3.Enabled = False
        btnPlayerAdmin.Enabled = False

        tmrBlink.Enabled = True

        Return 1
    End Function

    Private Function SetLEDoff()
        Dim varBuffer(3) As Byte
        varBuffer(0) = 0
        varBuffer(1) = 0
        varBuffer(2) = 0
        varBuffer(3) = 0

        BuzzerSetLEDs(varBuffer, varBuzzers)

        btnPlayer1.Enabled = True
        btnPlayer2.Enabled = True
        btnPlayer3.Enabled = True
        btnPlayerAdmin.Enabled = True

        tmrBlink.Enabled = False

        Return 1
    End Function

    Private Sub tmrButtons_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrButtons.Tick
        pbBuzzer1ButtonB.Visible = False
        pbBuzzer1Button1.Visible = False
        pbBuzzer1Button2.Visible = False
        pbBuzzer1Button3.Visible = False
        pbBuzzer1Button4.Visible = False
        pbBuzzer2ButtonB.Visible = False
        pbBuzzer2Button1.Visible = False
        pbBuzzer2Button2.Visible = False
        pbBuzzer2Button3.Visible = False
        pbBuzzer2Button4.Visible = False
        pbBuzzer3ButtonB.Visible = False
        pbBuzzer3Button1.Visible = False
        pbBuzzer3Button2.Visible = False
        pbBuzzer3Button3.Visible = False
        pbBuzzer3Button4.Visible = False
        pbBuzzer4ButtonB.Visible = False
        pbBuzzer4Button1.Visible = False
        pbBuzzer4Button2.Visible = False
        pbBuzzer4Button3.Visible = False
        pbBuzzer4Button4.Visible = False

        Dim varButtons(3) As Byte
        BuzzerGetButtons(varButtons, varBuzzers)

        'Buzzer 1
        If varButtons(0) = 1 Then
            pbBuzzer1ButtonB.Visible = True
            System.Windows.Forms.SendKeys.Send("q")
        End If

        If varButtons(0) = 2 Then
            pbBuzzer1Button4.Visible = True
            System.Windows.Forms.SendKeys.Send("4")
        End If

        If varButtons(0) = 4 Then
            pbBuzzer1Button3.Visible = True
            System.Windows.Forms.SendKeys.Send("3")
        End If

        If varButtons(0) = 8 Then
            pbBuzzer1Button2.Visible = True
            System.Windows.Forms.SendKeys.Send("2")
        End If

        If varButtons(0) = 16 Then
            pbBuzzer1Button1.Visible = True
            System.Windows.Forms.SendKeys.Send("1")
        End If

        'Buzzer 2
        If varButtons(1) = 1 Then
            pbBuzzer2ButtonB.Visible = True
            System.Windows.Forms.SendKeys.Send("b")
        End If

        If varButtons(1) = 2 Then
            pbBuzzer2Button4.Visible = True
            System.Windows.Forms.SendKeys.Send("4")
        End If

        If varButtons(1) = 4 Then
            pbBuzzer2Button3.Visible = True
            System.Windows.Forms.SendKeys.Send("3")
        End If

        If varButtons(1) = 8 Then
            pbBuzzer2Button2.Visible = True
            System.Windows.Forms.SendKeys.Send("2")
        End If

        If varButtons(1) = 16 Then
            pbBuzzer2Button1.Visible = True
            System.Windows.Forms.SendKeys.Send("1")
        End If

        'Buzzer 3
        If varButtons(2) = 1 Then
            pbBuzzer3ButtonB.Visible = True
            System.Windows.Forms.SendKeys.Send("p")
        End If

        If varButtons(2) = 2 Then
            pbBuzzer3Button4.Visible = True
            System.Windows.Forms.SendKeys.Send("4")
        End If

        If varButtons(2) = 4 Then
            pbBuzzer3Button3.Visible = True
            System.Windows.Forms.SendKeys.Send("3")
        End If

        If varButtons(2) = 8 Then
            pbBuzzer3Button2.Visible = True
            System.Windows.Forms.SendKeys.Send("2")
        End If

        If varButtons(2) = 16 Then
            pbBuzzer3Button1.Visible = True
            System.Windows.Forms.SendKeys.Send("1")
        End If

        'Buzzer 4
        If varButtons(3) = 1 Then
            pbBuzzer4ButtonB.Visible = True
            System.Windows.Forms.SendKeys.Send("{escape}")
        End If

        If varButtons(3) = 2 Then
            pbBuzzer4Button4.Visible = True
            'System.Windows.Forms.SendKeys.Send("4")
        End If

        If varButtons(3) = 4 Then
            pbBuzzer4Button3.Visible = True
            'System.Windows.Forms.SendKeys.Send("3")
        End If

        If varButtons(3) = 8 Then
            pbBuzzer4Button2.Visible = True
            'System.Windows.Forms.SendKeys.Send("2")
        End If

        If varButtons(3) = 16 Then
            pbBuzzer4Button1.Visible = True
            System.Windows.Forms.SendKeys.Send("n")
        End If
    End Sub
End Class
