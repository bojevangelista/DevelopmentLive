Imports System.Data.SqlClient
Public Class Form1
    Dim user = Environment.UserName
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\" + user + "\Documents\OPF-TableNo.txt")
    'SERVER'
    Dim con As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim con2 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim con3 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True

            If (TextBox1.TextLength = 6) Then
                Timer2.Enabled = False
                '''''''CHECK SCANNED SOMTRACK'''''''''''
                Dim ScannedCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted, PD.Status FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID Is Not NULL AND PD.Status IN (1,2) AND PH.SomtrackID = @Som AND TS.TableID = @TS AND PH.StationID <> 1 GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted, Pd.Status ORDER BY PH.DateStarted ASC"
                Dim ScannedCaseQuery As SqlCommand = New SqlCommand(ScannedCase, con)
                ScannedCaseQuery.Parameters.AddWithValue("@TS", TableSet)
                ScannedCaseQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                Dim accept = 0
                Dim SID = 0
                Dim Status = 0
                con.Open()

                Using reader As SqlDataReader = ScannedCaseQuery.ExecuteReader()

                    If reader.HasRows Then

                        While reader.Read()
                            SID = reader.Item("StationID")
                            Status = reader.Item("Status")
                            accept = 1

                        End While

                    End If
                End Using
                con.Close()

                If accept = 1 Then
                    StationActiveCase(SID, Status)
                    accept = 0
                End If


            End If


            TextBox1.Text = ""
            TextBox1.Focus()
        End If
    End Sub
    Private Sub StationActiveCase(SID, Status)

        '''''''CHECK STATION ACTIVE CASE'''''''''''

        If SID = 11 Or SID = 12 Then
            SID = "11,12"
        End If
        Dim ActiveCase As String = "Select PH.SomtrackID, PH.StationID, PD.Status FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID IN (@SID1, @SID2) AND PD.Status = 1 AND TS.TableID = @TS AND PH.SomtrackID <> @Som GROUP BY PH.SomtrackID, PH.StationID, Pd.Status"
        Dim ActiveCaseQuery As SqlCommand = New SqlCommand(ActiveCase, con)
        ActiveCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        If SID = "11,12" Then
            ActiveCaseQuery.Parameters.AddWithValue("@SID1", "11")
            ActiveCaseQuery.Parameters.AddWithValue("@SID2", "12")
        Else

            ActiveCaseQuery.Parameters.AddWithValue("@SID1", SID)
            ActiveCaseQuery.Parameters.AddWithValue("@SID2", 0)
        End If

        ActiveCaseQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
        Dim update = 0
        Dim CheckNextCase = 0

        con.Open()

        Using reader As SqlDataReader = ActiveCaseQuery.ExecuteReader()
            If reader.HasRows Then
                If SID = "11,12" Then
                    If reader.HasRows = 2 Then
                        'Err : Active Case'
                        StationError(SID, 1)
                    Else
                        If Status = 1 Then
                            'Update Case'
                            update = 1

                        Else
                            CheckNextCase = 1
                        End If
                    End If
                Else
                    'Err : Active Case'
                    StationError(SID, 1)
                End If

            Else
                If Status = 1 Then
                    'Update Case'
                    update = 1

                Else
                    CheckNextCase = 1
                End If
            End If
        End Using
        con.Close()

        If update = 1 Then

            ''''UPDATE ACTIVE DETAILS''''
            con.Open()
            Dim UpdateDetails As String = "update PD SET PD.Status = 5, PD.DateEnded = GETDATE() From [SMProduction].[dbo].[ProductionHead] as PH Left Join ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID WHERE PH.SomtrackID = @Som And PD.Status = 1"
            Dim UpdateDetailsQuery As SqlCommand = New SqlCommand(UpdateDetails, con)
            UpdateDetailsQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
            UpdateDetailsQuery.ExecuteNonQuery()
            con.Close()

            If (SID = "11,12") Then
                ''''UPDATE HEAD''''
                con.Open()
                Dim UpdateHead As String = "UPDATE ProductionHead SET DateEnded = GETDATE(), StationID = 0 WHERE SomtrackID = @Som"
                Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con)
                UpdateHeadQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                UpdateHeadQuery.ExecuteNonQuery()
                con.Close()
            Else
                ''''UPDATE HEAD''''
                con.Open()
                Dim UpdateHead As String = "UPDATE ProductionHead SET StationID = @SID WHERE SomtrackID = @Som"
                Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con)
                UpdateHeadQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                UpdateHeadQuery.Parameters.AddWithValue("@SID", SID + 1)
                UpdateHeadQuery.ExecuteNonQuery()
                con.Close()
            End If


            ''''UPDATE NEXT DETAILS''''
            con.Open()
            Dim UpdateNextDetails As String = "Update PD SET PD.Status = 2 FROM [SMProduction].[dbo].[ProductionHead] as PH LEFT JOIN StationProcess as SP ON SP.StationID = PH.StationID LEFT JOIN ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID And PD.BOMDID = SP.BOMDID WHERE PH.SomtrackID = @Som"
            Dim UpdateNextDetailsQuery As SqlCommand = New SqlCommand(UpdateNextDetails, con)
            UpdateNextDetailsQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
            UpdateNextDetailsQuery.ExecuteNonQuery()
            con.Close()
            GetActive()
            GetPending()
            GetPassedCase()
            StationSuccess(SID, 2)
        ElseIf CheckNextCase = 1 Then
            StationNextCase(SID)
            CheckNextCase = 0
        End If

    End Sub
    Private Sub StationNextCase(SID)


        '''''''CHECK STATION NEXT CASE'''''''''''
        Dim ScannedCase As String = "Select TOP 1 PH.SomtrackID, PH.StationID, PD.Status, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID IN (@SID1, @SID2) AND PD.Status = 2 AND TS.TableID = 1 GROUP BY PH.SomtrackID, PH.StationID, Pd.Status, PH.DateStarted ORDER BY DateStarted ASC"
        Dim ScannedCaseQuery As SqlCommand = New SqlCommand(ScannedCase, con)
        ScannedCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        If SID = "11,12" Then
            ScannedCaseQuery.Parameters.AddWithValue("@SID1", "11")
            ScannedCaseQuery.Parameters.AddWithValue("@SID2", "12")
        Else

            ScannedCaseQuery.Parameters.AddWithValue("@SID1", SID)
            ScannedCaseQuery.Parameters.AddWithValue("@SID2", 0)
        End If


        Dim update = 0

        con.Open()
        Using reader As SqlDataReader = ScannedCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If TextBox1.Text = reader.Item("SomtrackID") Then
                        'update case'
                        update = 1

                    Else
                        'Err : On Queue'
                        StationError(SID, 2)
                    End If
                End While
            Else
            End If
        End Using
        con.Close()


        If update = 1 Then
            Dim UpdateNextDetails As String = ""
            ''''UPDATE NEXT CASE DETAILS''''


            If SID = "11,12" Then
                UpdateNextDetails = "UPDATE PD Set PD.EmployeeID = TM.EmployeeID , PD.DateStarted = GETDATE(), PD.Status = 1 FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PD.Status = 2 And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.SomtrackID = @Som AND TM.EmployeeID NOT IN ( SELECT [EmployeeID] FROM [ProductionDetails] WHERE Status = 1)"
            Else
                UpdateNextDetails = "UPDATE PD Set PD.EmployeeID = TM.EmployeeID , PD.DateStarted = GETDATE(), PD.Status = 1 FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PD.Status = 2 And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.SomtrackID = @Som"

            End If
            con.Open()
            Dim UpdateNextDetailsQuery As SqlCommand = New SqlCommand(UpdateNextDetails, con)
            UpdateNextDetailsQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
            UpdateNextDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
            UpdateNextDetailsQuery.ExecuteNonQuery()
            con.Close()

            If SID = "11,12" Then
                con.Open()
                Dim UpdateHeadStation As String = "Update PH SET PH.StationID = TM.StationID FROM [ProductionDetails] as PD LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFt JOIN TableMembers as TM ON TM.EmployeeID = PD.EmployeeID LEFt JOIN StationProcess as SP ON SP.BOMDID = PD.BOMDID WHERE PH.SomtrackID = @Som AND SP.StationID = 12 AND TM.StationID = 12"
                Dim UpdateHeadStationQuery As SqlCommand = New SqlCommand(UpdateHeadStation, con)
                UpdateHeadStationQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                UpdateHeadStationQuery.ExecuteNonQuery()
                con.Close()
            End If


            GetActive()
            GetPending()
            GetPassedCase()
            StationSuccess(SID, 1)




        End If
    End Sub

    Private Sub StationError(SID, errID)
        Dim ErrorMessage As String

        If errID = 1 Then
            ErrorMessage = "Active Case"

            If SID = 2 Then
                Label27.ForeColor = Color.Red
                Label27.Text = ErrorMessage
            ElseIf SID = 3 Then
                Label24.ForeColor = Color.Red
                Label24.Text = ErrorMessage
            ElseIf SID = 4 Then
                Label15.ForeColor = Color.Red
                Label15.Text = ErrorMessage
            ElseIf SID = 5 Then
                Label12.ForeColor = Color.Red
                Label12.Text = ErrorMessage
            ElseIf SID = 6 Then
                Label1.ForeColor = Color.Red
                Label1.Text = ErrorMessage
            ElseIf SID = 7 Then
                Label6.ForeColor = Color.Red
                Label6.Text = ErrorMessage
            ElseIf SID = 8 Then
                Label9.ForeColor = Color.Red
                Label9.Text = ErrorMessage
            ElseIf SID = 9 Then
                Label18.ForeColor = Color.Red
                Label18.Text = ErrorMessage
            ElseIf SID = 10 Then
                Label21.ForeColor = Color.Red
                Label21.Text = ErrorMessage
            ElseIf SID = "11,12" Then
                Label30.ForeColor = Color.Red
                Label30.Text = ErrorMessage
                Label36.ForeColor = Color.Red
                Label36.Text = ErrorMessage

            End If

        ElseIf errID = 2 Then
            ErrorMessage = "Case on queue"

            If SID = 2 Then
                Label27.ForeColor = Color.Red
                Label27.Text = ErrorMessage
            ElseIf SID = 3 Then
                Label24.ForeColor = Color.Red
                Label24.Text = ErrorMessage
            ElseIf SID = 4 Then
                Label15.ForeColor = Color.Red
                Label15.Text = ErrorMessage
            ElseIf SID = 5 Then
                Label12.ForeColor = Color.Red
                Label12.Text = ErrorMessage
            ElseIf SID = 6 Then
                Label1.ForeColor = Color.Red
                Label1.Text = ErrorMessage
            ElseIf SID = 7 Then
                Label6.ForeColor = Color.Red
                Label6.Text = ErrorMessage
            ElseIf SID = 8 Then
                Label9.ForeColor = Color.Red
                Label9.Text = ErrorMessage
            ElseIf SID = 9 Then
                Label18.ForeColor = Color.Red
                Label18.Text = ErrorMessage
            ElseIf SID = 10 Then
                Label21.ForeColor = Color.Red
                Label21.Text = ErrorMessage
            ElseIf SID = "11,12" Then
                Label30.ForeColor = Color.Red
                Label30.Text = ErrorMessage
                Label36.ForeColor = Color.Red
                Label36.Text = ErrorMessage

            End If
        End If

    End Sub

    Private Sub StationSuccess(SID, succID)
        Dim ErrorMessage As String

        If succID = 1 Then
            ErrorMessage = "Accepted"

            If SID = 2 Then
                Label27.ForeColor = Color.Lime
                Label27.Text = ErrorMessage
            ElseIf SID = 3 Then
                Label24.ForeColor = Color.Lime
                Label24.Text = ErrorMessage
            ElseIf SID = 4 Then
                Label15.ForeColor = Color.Lime
                Label15.Text = ErrorMessage
            ElseIf SID = 5 Then
                Label12.ForeColor = Color.Lime
                Label12.Text = ErrorMessage
            ElseIf SID = 6 Then
                Label1.ForeColor = Color.Lime
                Label1.Text = ErrorMessage
            ElseIf SID = 7 Then
                Label6.ForeColor = Color.Lime
                Label6.Text = ErrorMessage
            ElseIf SID = 8 Then
                Label9.ForeColor = Color.Lime
                Label9.Text = ErrorMessage
            ElseIf SID = 9 Then
                Label18.ForeColor = Color.Lime
                Label18.Text = ErrorMessage
            ElseIf SID = 10 Then
                Label21.ForeColor = Color.Lime
                Label21.Text = ErrorMessage
            ElseIf SID = 11 Then
                Label30.ForeColor = Color.Lime
                Label30.Text = ErrorMessage
            ElseIf SID = 12 Then
                Label36.ForeColor = Color.Lime
                Label36.Text = ErrorMessage

            End If

        ElseIf succID = 2 Then
            ErrorMessage = "Case Passed"

            If SID = 2 Then
                Label27.ForeColor = Color.Lime
                Label27.Text = ErrorMessage
            ElseIf SID = 3 Then
                Label24.ForeColor = Color.Lime
                Label24.Text = ErrorMessage
            ElseIf SID = 4 Then
                Label15.ForeColor = Color.Lime
                Label15.Text = ErrorMessage
            ElseIf SID = 5 Then
                Label12.ForeColor = Color.Lime
                Label12.Text = ErrorMessage
            ElseIf SID = 6 Then
                Label1.ForeColor = Color.Lime
                Label1.Text = ErrorMessage
            ElseIf SID = 7 Then
                Label6.ForeColor = Color.Lime
                Label6.Text = ErrorMessage
            ElseIf SID = 8 Then
                Label9.ForeColor = Color.Lime
                Label9.Text = ErrorMessage
            ElseIf SID = 9 Then
                Label18.ForeColor = Color.Lime
                Label18.Text = ErrorMessage
            ElseIf SID = 10 Then
                Label21.ForeColor = Color.Lime
                Label21.Text = ErrorMessage
            ElseIf SID = 11 Then
                Label30.ForeColor = Color.Lime
                Label30.Text = ErrorMessage
            ElseIf SID = 12 Then
                Label36.ForeColor = Color.Lime
                Label36.Text = ErrorMessage

            End If
        End If

    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        GetActive()
        GetPending()
        GetPassedCase()
        Timer2.Enabled = True

    End Sub
    Private Sub GetActive()
        Label32.Text = ""
        Label26.Text = ""
        Label23.Text = ""
        Label14.Text = ""
        Label11.Text = ""
        Label2.Text = ""
        Label5.Text = ""
        Label8.Text = ""
        Label17.Text = ""
        Label20.Text = ""
        Label29.Text = ""
        Label35.Text = ""
        '''''''CHECK TABLE CASES'''''''''''
        Dim TableActiveCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PD.Status = 1 And PH.StationID Is Not NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableActiveCaseQuery As SqlCommand = New SqlCommand(TableActiveCase, con2)
        TableActiveCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TableActiveCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        Label32.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 2) Then
                        Label26.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 3) Then
                        Label23.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 4) Then
                        Label14.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 5) Then
                        Label11.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 6) Then
                        Label2.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 7) Then
                        Label5.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 8) Then
                        Label8.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 9) Then
                        Label17.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 10) Then
                        Label20.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 11) Then
                        Label29.Text = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 12) Then
                        Label35.Text = reader.Item("SomtrackID").ToString
                    End If
                End While

            End If
        End Using
        con2.Close()

    End Sub
    Private Sub GetPassedCase()
        Label45.Text = "0"
        Label44.Text = "0"
        Label40.Text = "0"
        Label39.Text = "0"
        Label38.Text = "0"
        Label37.Text = "0"
        Label41.Text = "0"
        Label42.Text = "0"
        Label43.Text = "0"
        Label46.Text = "0"
        Label47.Text = "0"
        Label48.Text = "0"
        '''''''CHECK PASSED CASES'''''''''''
        Dim TablePassedCase As String = "SELECT SP.StationID, COUNT(PD.ProductionDetailID) as Done FROM [SMProduction].[dbo].[ProductionDetails] as PD LEFT JOIN TableMembers as TM ON TM.EmployeeID = PD.EmployeeID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID LEFT JOIN StationProcess as SP ON SP.BOMDID = PD.BOMDID WHERE TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND TS.TableID = @TS AND PD.Status = 5 AND SP.StationID = TM.StationID GROUP BY SP.StationID"
        Dim TablePassedCasQuery As SqlCommand = New SqlCommand(TablePassedCase, con2)
        TablePassedCasQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TablePassedCasQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        Label45.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 2) Then
                        Label44.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 3) Then
                        Label40.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 4) Then
                        Label39.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 5) Then
                        Label38.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 6) Then
                        Label37.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 7) Then
                        Label41.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 8) Then
                        Label42.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 9) Then
                        Label43.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 10) Then
                        Label46.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 11) Then
                        Label47.Text = reader.Item("Done").ToString
                    ElseIf (reader.Item("StationID") = 12) Then
                        Label48.Text = reader.Item("Done").ToString
                    End If
                End While

            End If
        End Using
        con2.Close()
    End Sub
    Private Sub GetPending()
        Label25.Text = ""
        Label22.Text = ""
        Label13.Text = ""
        Label10.Text = ""
        Label3.Text = ""
        Label4.Text = ""
        Label7.Text = ""
        Label16.Text = ""
        Label19.Text = ""
        Label28.Text = ""
        Label34.Text = ""



        '''''''CHECK TABLE CASES'''''''''''
        Dim TableCase As String = "SELECT PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PD.Status = 2 AND PH.StationID IS NOT NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableCaseQuery As SqlCommand = New SqlCommand(TableCase, con2)
        TableCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TableCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 2) Then
                        If (InStr(Label25.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label25.Text = Label25.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 3) Then
                        If (InStr(Label22.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label22.Text = Label22.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 4) Then
                        If (InStr(Label13.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label13.Text = Label13.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 5) Then
                        If (InStr(Label10.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label10.Text = Label10.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 6) Then
                        If (InStr(Label3.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label3.Text = Label3.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 7) Then
                        If (InStr(Label4.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label4.Text = Label4.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 8) Then
                        If (InStr(Label7.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label7.Text = Label7.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 9) Then
                        If (InStr(Label16.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label16.Text = Label16.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 10) Then
                        If (InStr(Label19.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label19.Text = Label19.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    ElseIf (reader.Item("StationID") = 11) Then
                        If (InStr(Label28.Text, reader.Item("SomtrackID"))) Then
                        Else
                            Label28.Text = Label28.Text + reader.Item("SomtrackID").ToString + ", "
                            Label34.Text = Label34.Text + reader.Item("SomtrackID").ToString + ", "
                        End If
                    End If
                End While

            End If
        End Using
        con2.Close()
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.Select()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetActive()
        GetPending()
        GetPassedCase()


    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Label33.Text = ""
        Label27.Text = ""
        Label24.Text = ""
        Label15.Text = ""
        Label12.Text = ""
        Label1.Text = ""
        Label6.Text = ""
        Label9.Text = ""
        Label18.Text = ""
        Label21.Text = ""
        Label30.Text = ""
        Label36.Text = ""

        '''''''CHECK CASE DURATION ''''''''''
        Dim CaseDuration As String = "SELECT SP.StationID, DATEDIFF(second,PD.DateStarted,GETDATE()) as Duration FROM [SMProduction].[dbo].[ProductionDetails] as PD LEFT JOIN TableMembers as TM ON TM.EmployeeID = PD.EmployeeID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID LEFT JOIN StationProcess as SP ON SP.BOMDID = PD.BOMDID WHERE TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND TS.TableID = @TS AND PD.Status = 1 AND SP.StationID = TM.StationID GROUP BY SP.StationID, PD.DateStarted,PD.DateEnded"
        Dim CaseDurationQuery As SqlCommand = New SqlCommand(CaseDuration, con3)
        CaseDurationQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = CaseDurationQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        Label33.ForeColor = Color.Lime
                        Label33.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 2) Then
                        Label27.ForeColor = Color.Lime
                        Label27.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 3) Then
                        Label24.ForeColor = Color.Lime
                        Label24.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 4) Then
                        Label15.ForeColor = Color.Lime
                        Label15.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 5) Then
                        Label12.ForeColor = Color.Lime
                        Label12.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 6) Then
                        Label1.ForeColor = Color.Lime
                        Label1.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 7) Then
                        Label6.ForeColor = Color.Lime
                        Label6.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 8) Then
                        Label9.ForeColor = Color.Lime
                        Label9.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 9) Then
                        Label18.ForeColor = Color.Lime
                        Label18.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 10) Then
                        Label21.ForeColor = Color.Lime
                        Label21.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 11) Then
                        Label30.ForeColor = Color.Lime
                        Label30.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    ElseIf (reader.Item("StationID") = 12) Then
                        Label36.ForeColor = Color.Lime
                        Label36.Text = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                    End If
                End While

            End If
        End Using
        con3.Close()


        '''''''CHECK TABLE CASES'''''''''''
        Dim TableCase As String = "SELECT PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND PD.Status = 2 AND PH.StationID IS NOT NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableCaseQuery As SqlCommand = New SqlCommand(TableCase, con3)
        TableCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = TableCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 2) Then
                        Label27.ForeColor = Color.DarkOrange
                        Label27.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 3) Then
                        Label24.ForeColor = Color.DarkOrange
                        Label24.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 4) Then
                        Label15.ForeColor = Color.DarkOrange
                        Label15.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 5) Then
                        Label12.ForeColor = Color.DarkOrange
                        Label12.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 6) Then
                        Label1.ForeColor = Color.DarkOrange
                        Label1.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 7) Then
                        Label6.ForeColor = Color.DarkOrange
                        Label6.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 8) Then
                        Label9.ForeColor = Color.DarkOrange
                        Label9.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 9) Then
                        Label18.ForeColor = Color.DarkOrange
                        Label18.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 10) Then
                        Label21.ForeColor = Color.DarkOrange
                        Label21.Text = "Idle"

                    ElseIf (reader.Item("StationID") = 11) Then

                        If Label29.Text = "" Then
                            Label30.ForeColor = Color.DarkOrange
                            Label30.Text = "Idle"
                        End If
                        If Label35.Text = "" Then
                            Label36.ForeColor = Color.DarkOrange
                            Label36.Text = "Idle"
                        End If

                    End If
                End While

            End If
        End Using
        con3.Close()



    End Sub
End Class
