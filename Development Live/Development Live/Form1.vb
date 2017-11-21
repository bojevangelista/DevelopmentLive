Public Class Form1
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True


            If ((InStr(Label34.Text, TextBox1.Text)) Or (Label35.Text = TextBox1.Text)) Then
                If (Label35.Text = TextBox1.Text) Then
                    Label35.Text = "-"
                    Label36.Text = "Task Completed"
                    Label48.Text = Label48.Text + 1
                ElseIf (TextBox1.Text = Label34.Text.Substring(1, 6)) Then
                    If (Label35.Text = "-") Then
                        Label35.Text = TextBox1.Text
                        Label36.Text = "Scan Successful"
                        Label34.Text = Label34.Text.Replace(Label34.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label36.Text = "Task on going"
                    End If


                Else
                    Label36.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label28.Text, TextBox1.Text)) Or (Label29.Text = TextBox1.Text)) Then
                If (Label29.Text = TextBox1.Text) Then
                    Label34.Text = Label34.Text + Label29.Text + ", "
                    Label29.Text = "-"
                    Label30.Text = "Task Completed"
                    Label47.Text = Label47.Text + 1
                ElseIf (TextBox1.Text = Label28.Text.Substring(1, 6)) Then
                    If (Label29.Text = "-") Then
                        Label29.Text = TextBox1.Text
                        Label30.Text = "Scan Successful"
                        Label28.Text = Label28.Text.Replace(Label28.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label30.Text = "Task on going"
                    End If


                Else
                    Label30.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label19.Text, TextBox1.Text)) Or (Label20.Text = TextBox1.Text)) Then
                If (Label20.Text = TextBox1.Text) Then
                    Label28.Text = Label28.Text + Label20.Text + ", "
                    Label20.Text = "-"
                    Label21.Text = "Task Completed"
                    Label46.Text = Label46.Text + 1
                ElseIf (TextBox1.Text = Label19.Text.Substring(1, 6)) Then
                    If (Label20.Text = "-") Then
                        Label20.Text = TextBox1.Text
                        Label21.Text = "Scan Successful"
                        Label19.Text = Label19.Text.Replace(Label19.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label21.Text = "Task on going"
                    End If


                Else
                    Label21.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label16.Text, TextBox1.Text)) Or (Label17.Text = TextBox1.Text)) Then
                If (Label17.Text = TextBox1.Text) Then
                    Label19.Text = Label19.Text + Label17.Text + ", "
                    Label17.Text = "-"
                    Label18.Text = "Task Completed"
                    Label43.Text = Label43.Text + 1
                ElseIf (TextBox1.Text = Label16.Text.Substring(1, 6)) Then
                    If (Label17.Text = "-") Then
                        Label17.Text = TextBox1.Text
                        Label18.Text = "Scan Successful"
                        Label16.Text = Label16.Text.Replace(Label16.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label18.Text = "Task on going"
                    End If


                Else
                    Label18.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label7.Text, TextBox1.Text)) Or (Label8.Text = TextBox1.Text)) Then
                If (Label8.Text = TextBox1.Text) Then
                    Label16.Text = Label16.Text + Label8.Text + ", "
                    Label8.Text = "-"
                    Label9.Text = "Task Completed"
                    Label42.Text = Label42.Text + 1
                ElseIf (TextBox1.Text = Label7.Text.Substring(1, 6)) Then
                    If (Label8.Text = "-") Then
                        Label8.Text = TextBox1.Text
                        Label9.Text = "Scan Successful"
                        Label7.Text = Label7.Text.Replace(Label7.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label9.Text = "Task on going"
                    End If


                Else
                    Label9.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label4.Text, TextBox1.Text)) Or (Label5.Text = TextBox1.Text)) Then
                If (Label5.Text = TextBox1.Text) Then
                    Label7.Text = Label7.Text + Label5.Text + ", "
                    Label5.Text = "-"
                    Label6.Text = "Task Completed"
                    Label41.Text = Label41.Text + 1
                ElseIf (TextBox1.Text = Label4.Text.Substring(1, 6)) Then
                    If (Label5.Text = "-") Then
                        Label5.Text = TextBox1.Text
                        Label6.Text = "Scan Successful"
                        Label4.Text = Label4.Text.Replace(Label4.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label6.Text = "Task on going"
                    End If


                Else
                    Label6.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label3.Text, TextBox1.Text)) Or (Label2.Text = TextBox1.Text)) Then
                If (Label2.Text = TextBox1.Text) Then
                    Label4.Text = Label4.Text + Label2.Text + ", "
                    Label2.Text = "-"
                    Label1.Text = "Task Completed"
                    Label37.Text = Label37.Text + 1
                ElseIf (TextBox1.Text = Label3.Text.Substring(1, 6)) Then
                    If (Label2.Text = "-") Then
                        Label2.Text = TextBox1.Text
                        Label1.Text = "Scan Successful"
                        Label3.Text = Label3.Text.Replace(Label3.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label1.Text = "Task on going"
                    End If


                Else
                    Label1.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label10.Text, TextBox1.Text)) Or (Label11.Text = TextBox1.Text)) Then
                If (Label11.Text = TextBox1.Text) Then
                    Label3.Text = Label3.Text + Label11.Text + ", "
                    Label11.Text = "-"
                    Label12.Text = "Task Completed"
                    Label38.Text = Label38.Text + 1
                ElseIf (TextBox1.Text = Label10.Text.Substring(1, 6)) Then
                    If (Label11.Text = "-") Then
                        Label11.Text = TextBox1.Text
                        Label12.Text = "Scan Successful"
                        Label10.Text = Label10.Text.Replace(Label10.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label12.Text = "Task on going"
                    End If


                Else
                    Label12.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label13.Text, TextBox1.Text)) Or (Label14.Text = TextBox1.Text)) Then
                If (Label14.Text = TextBox1.Text) Then
                    Label10.Text = Label10.Text + Label14.Text + ", "
                    Label14.Text = "-"
                    Label15.Text = "Task Completed"
                    Label39.Text = Label39.Text + 1
                ElseIf (TextBox1.Text = Label13.Text.Substring(1, 6)) Then
                    If (Label14.Text = "-") Then
                        Label14.Text = TextBox1.Text
                        Label15.Text = "Scan Successful"
                        Label13.Text = Label13.Text.Replace(Label13.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label15.Text = "Task on going"
                    End If


                Else
                    Label15.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label22.Text, TextBox1.Text)) Or (Label23.Text = TextBox1.Text)) Then
                If (Label23.Text = TextBox1.Text) Then
                    Label13.Text = Label13.Text + Label23.Text + ", "
                    Label23.Text = "-"
                    Label24.Text = "Task Completed"
                    Label40.Text = Label40.Text + 1
                ElseIf (TextBox1.Text = Label22.Text.Substring(1, 6)) Then
                    If (Label23.Text = "-") Then
                        Label23.Text = TextBox1.Text
                        Label24.Text = "Scan Successful"
                        Label22.Text = Label22.Text.Replace(Label22.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label24.Text = "Task on going"
                    End If


                Else
                    Label24.Text = "Still on queue"
                End If
            ElseIf ((InStr(Label25.Text, TextBox1.Text)) Or (Label26.Text = TextBox1.Text)) Then
                If (Label26.Text = TextBox1.Text) Then
                    Label22.Text = Label22.Text + Label26.Text + ", "
                    Label26.Text = "-"
                    Label27.Text = "Task Completed"
                    Label44.Text = Label44.Text + 1

                ElseIf (TextBox1.Text = Label25.Text.Substring(1, 6)) Then
                    If (Label26.Text = "-") Then
                        Label26.Text = TextBox1.Text
                        Label27.Text = "Scan Successful"
                        Label25.Text = Label25.Text.Replace(Label25.Text.Substring(1, 6) + ", ", "")
                    Else
                        Label27.Text = "Task on going"
                    End If


                Else
                    Label27.Text = "Still on queue"
                End If
            Else

                If (Label32.Text = TextBox1.Text) Then
                    Label25.Text = Label25.Text + Label32.Text + ", "
                    Label32.Text = "-"
                    Label33.Text = "Task Completed"
                    Label45.Text = Label45.Text + 1
                ElseIf (Label32.Text = "-") Then
                    Label32.Text = TextBox1.Text
                    Label33.Text = "Scan Successful"
                Else
                    Label33.Text = "Task on going"
                End If
            End If










            TextBox1.Text = ""
                TextBox1.Focus()
            End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
