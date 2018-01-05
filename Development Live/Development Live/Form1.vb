Imports System.Data.SqlClient
Public Class Form1
    Dim user = Environment.UserName
    Dim TableSet = My.Computer.FileSystem.ReadAllText("C:\Users\" + user + "\Documents\OPF-TableNo.txt")
    'SERVER'
    Dim con As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim con2 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")
    Dim con3 As SqlConnection = New SqlConnection("Data Source=SOMNOMED-IBM;Initial Catalog=SMProduction;User ID=SOMNOMED-IBM-Guest;Password=Somnomed01")

    Dim s1case = 0
    Dim s2case = 0
    Dim s3case = 0
    Dim s4case = 0
    Dim s5case = 0
    Dim s6case = 0
    Dim s7case = 0
    Dim s8case = 0
    Dim s9case = 0
    Dim s10case = 0
    Dim s11case = 0

    Dim s1timer = 0
    Dim s2timer = 0
    Dim s3timer = 0
    Dim s4timer = 0
    Dim s5timer = 0
    Dim s6timer = 0
    Dim s7timer = 0
    Dim s8timer = 0
    Dim s9timer = 0
    Dim s10timer = 0
    Dim s11timer = 0
    Dim s12timer = 0

    Dim showwarning = 0


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True

            If (TextBox1.TextLength = 6) Then
                Timer2.Enabled = False
                '''''''CHECK SCANNED SOMTRACK'''''''''''
                Dim ScannedCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted, PD.Status,TS.TableSetID FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID Is Not NULL AND PD.Status IN (1,2) AND PH.SomtrackID = @Som AND TS.TableID = @TS AND PH.TableNo = @TS GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted, Pd.Status, TS.TableSetID  ORDER BY PH.DateStarted ASC"
                Dim ScannedCaseQuery As SqlCommand = New SqlCommand(ScannedCase, con)
                ScannedCaseQuery.Parameters.AddWithValue("@TS", TableSet)
                ScannedCaseQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                Dim accept = 0
                Dim SID = 0
                Dim Status = 0
                Dim CurrentSet = 0
                con.Open()

                Using reader As SqlDataReader = ScannedCaseQuery.ExecuteReader()

                    If reader.HasRows Then

                        While reader.Read()
                            SID = reader.Item("StationID")
                            Status = reader.Item("Status")
                            CurrentSet = reader.Item("TableSetID")

                            accept = 1

                        End While

                    End If
                End Using
                con.Close()

                If accept = 1 Then
                    StationActiveCase(SID, Status, CurrentSet)
                    accept = 0
                End If


            End If


            TextBox1.Text = ""
            TextBox1.Focus()
        End If
    End Sub
    Private Sub StationActiveCase(SID, Status, CurrentSet)

        '''''''CHECK STATION ACTIVE CASE'''''''''''

        If SID = 11 Or SID = 12 Then
            SID = "11,12"
        End If
        Dim ActiveCase As String = "Select PH.SomtrackID, PH.StationID, PD.Status FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID IN (@SID1, @SID2) AND PD.Status = 1 AND TS.TableID = @TS AND PH.TableNo = @TS AND PH.SomtrackID <> @Som GROUP BY PH.SomtrackID, PH.StationID, Pd.Status"
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
                Dim UpdateHead As String = "UPDATE ProductionHead SET DateEnded = GETDATE(), StationID = 0, TableSetID = @TSID WHERE SomtrackID = @Som"
                Dim UpdateHeadQuery As SqlCommand = New SqlCommand(UpdateHead, con)
                UpdateHeadQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                UpdateHeadQuery.Parameters.AddWithValue("@TSID", CurrentSet)

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
            Dim UpdateNextDetails As String = "Update PD SET PD.Status = 2 FROM [SMProduction].[dbo].[ProductionHead] as PH LEFT JOIN StationProcess as SP ON SP.StationID = PH.StationID LEFT JOIN ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID And PD.BOMDID = SP.BOMDID WHERE PH.SomtrackID = @Som AND PD.Status = 3"
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
        Dim ScannedCase As String = "Select TOP 1 PH.SomtrackID, PH.StationID, PD.Status, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.StationID IN (@SID1, @SID2) AND PD.Status = 2 AND TS.TableID = @TS AND PH.TableNo = @TS GROUP BY PH.SomtrackID, PH.StationID, Pd.Status, PH.DateStarted ORDER BY DateStarted ASC"
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
                UpdateNextDetails = "UPDATE PD Set PD.EmployeeID = TM.EmployeeID, PD.DateStarted = GETDATE(), PD.Status = 1 FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND PH.TableNo = @TS And PD.Status = 2 And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.SomtrackID = @Som AND TM.EmployeeID NOT IN ( SELECT [EmployeeID] FROM [ProductionDetails] WHERE Status = 1 ) AND SP.StationID = (SELECT MIN(SP.StationID) FROM ProductionDetails as PD LEFT JOIN StationProcess as SP ON PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead as PH ON PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers as TM ON TM.StationID = SP.StationID LEFT JOIN TableSet as TS ON TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS AND PH.TableNo = @TS And PD.Status = 2 And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.SomtrackID = @Som AND TM.EmployeeID NOT IN ( SELECT [EmployeeID] FROM [ProductionDetails] WHERE Status = 1 ))"
            Else
                UpdateNextDetails = "UPDATE PD Set PD.EmployeeID = TM.EmployeeID , PD.DateStarted = GETDATE(), PD.Status = 1 FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PH.TableNo = @TS And PD.Status = 2 And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PH.SomtrackID = @Som"

            End If
            con.Open()
            Dim UpdateNextDetailsQuery As SqlCommand = New SqlCommand(UpdateNextDetails, con)
            UpdateNextDetailsQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
            UpdateNextDetailsQuery.Parameters.AddWithValue("@TS", TableSet)
            UpdateNextDetailsQuery.ExecuteNonQuery()
            con.Close()

            If SID = "11,12" Then
                con.Open()
                Dim UpdateHeadStation As String = "Update PH Set PH.StationID = TM.StationID FROM ProductionHead As PH left join ProductionDetails as PD ON PD.ProductionHeadID = PH.ProductionHeadID LEFt JOIN TableMembers As TM On TM.EmployeeID = PD.EmployeeID WHERE PH.SomtrackID = @Som AND PD.Status = 1 AND TM.TableMemberStatus = 1 AND TM.StationID IN (11, 12)"
                Dim UpdateHeadStationQuery As SqlCommand = New SqlCommand(UpdateHeadStation, con)
                UpdateHeadStationQuery.Parameters.AddWithValue("@Som", TextBox1.Text)
                UpdateHeadStationQuery.ExecuteNonQuery()
                con.Close()
            ElseIf SID = "1" Then
                con.Open()
                Dim UpdateHeadStation As String = "Update PH Set PH.DateStarted = GETDATE() FROM ProductionHead As PH  WHERE PH.SomtrackID = @Som "
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

            If SID = 1 Then
                Label33.ForeColor = Color.Red
                Label33.Text = ErrorMessage
            ElseIf SID = 2 Then
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
            ErrorMessage = "On queue"

            If SID = 1 Then
                Label33.ForeColor = Color.Red
                Label33.Text = ErrorMessage
            ElseIf SID = 2 Then
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

            If SID = 1 Then
                Label33.ForeColor = Color.Lime
                Label33.Text = ErrorMessage
            ElseIf SID = 2 Then
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

            If SID = 1 Then
                Label33.ForeColor = Color.Lime
                Label33.Text = ErrorMessage
            ElseIf SID = 2 Then
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
        Dim s1active = ""
        Dim s2active = ""
        Dim s3active = ""
        Dim s4active = ""
        Dim s5active = ""
        Dim s6active = ""
        Dim s7active = ""
        Dim s8active = ""
        Dim s9active = ""
        Dim s10active = ""
        Dim s11active = ""
        Dim s12active = ""


        '''''''CHECK TABLE CASES'''''''''''
        Dim TableActiveCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PH.TableNo = @TS And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PD.Status = 1 And PH.StationID Is Not NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableActiveCaseQuery As SqlCommand = New SqlCommand(TableActiveCase, con2)
        TableActiveCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TableActiveCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        s1active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 2) Then
                        s2active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 3) Then
                        s3active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 4) Then
                        s4active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 5) Then
                        s5active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 6) Then
                        s6active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 7) Then
                        s7active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 8) Then
                        s8active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 9) Then
                        s9active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 10) Then
                        s10active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 11) Then
                        s11active = reader.Item("SomtrackID").ToString
                    ElseIf (reader.Item("StationID") = 12) Then
                        s12active = reader.Item("SomtrackID").ToString
                    End If
                End While

            End If
        End Using



        Label32.Text = s1active
        Label32.Text = s1active
        Label26.Text = s2active
        Label26.Text = s2active
        Label23.Text = s3active
        Label14.Text = s4active
        Label11.Text = s5active
        Label2.Text = s6active
        Label5.Text = s7active
        Label8.Text = s8active
        Label17.Text = s9active
        Label20.Text = s10active
        Label29.Text = s11active
        Label35.Text = s12active








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

        Label56.Text = "0"
        Label55.Text = "0"
        Label54.Text = "0"
        Label53.Text = "0"
        Label50.Text = "0"
        Label49.Text = "0"
        Label51.Text = "0"
        Label52.Text = "0"
        Label60.Text = "0"
        Label59.Text = "0"
        Label58.Text = "0"
        Label57.Text = "0"
        '''''''CHECK PASSED CASES'''''''''''
        Dim TablePassedCase As String = "Select SP.StationID, SUM(Case When PD.Status = 5 Then 1 Else 0 End) As Done , SUM(Case When PD.Status = 4 Then 1 Else 0 End) As Redo FROM [SMProduction].[dbo].[ProductionDetails] As PD LEFT JOIN TableMembers As TM On TM.EmployeeID = PD.EmployeeID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID LEFT JOIN StationProcess As SP On SP.BOMDID = PD.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID WHERE TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And TS.TableID = @TS And PH.TableNo = @TS And PD.Status In (4,5) And SP.StationID = TM.StationID And PD.DateEnded BETWEEN TM.TableMemberTimeIn And GETDATE() GROUP BY SP.StationID"
        Dim TablePassedCasQuery As SqlCommand = New SqlCommand(TablePassedCase, con2)
        TablePassedCasQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TablePassedCasQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        Label45.Text = reader.Item("Done").ToString
                        Label56.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 2) Then
                        Label44.Text = reader.Item("Done").ToString
                        Label55.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 3) Then
                        Label40.Text = reader.Item("Done").ToString
                        Label54.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 4) Then
                        Label39.Text = reader.Item("Done").ToString
                        Label53.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 5) Then
                        Label38.Text = reader.Item("Done").ToString
                        Label50.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 6) Then
                        Label37.Text = reader.Item("Done").ToString
                        Label49.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 7) Then
                        Label41.Text = reader.Item("Done").ToString
                        Label51.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 8) Then
                        Label42.Text = reader.Item("Done").ToString
                        Label52.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 9) Then
                        Label43.Text = reader.Item("Done").ToString
                        Label60.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 10) Then
                        Label46.Text = reader.Item("Done").ToString
                        Label59.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 11) Then
                        Label47.Text = reader.Item("Done").ToString
                        Label58.Text = reader.Item("Redo").ToString
                    ElseIf (reader.Item("StationID") = 12) Then
                        Label48.Text = reader.Item("Done").ToString
                        Label57.Text = reader.Item("Redo").ToString
                    End If
                End While

            End If
        End Using
        con2.Close()
    End Sub
    Private Sub GetPending()


        Dim s1pending = ""
        Dim s2pending = ""
        Dim s3pending = ""
        Dim s4pending = ""
        Dim s5pending = ""
        Dim s6pending = ""
        Dim s7pending = ""
        Dim s8pending = ""
        Dim s9pending = ""
        Dim s10pending = ""
        Dim s11pending = ""

        s1case = 0
        s2case = 0
        s3case = 0
        s4case = 0
        s5case = 0
        s6case = 0
        s7case = 0
        s8case = 0
        s9case = 0
        s10case = 0
        s11case = 0


        '''''''CHECK TABLE CASES'''''''''''
        Dim TableCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PH.TableNo = @TS And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PD.Status = 2 And PH.StationID Is Not NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableCaseQuery As SqlCommand = New SqlCommand(TableCase, con2)
        TableCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con2.Open()
        Using reader As SqlDataReader = TableCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        s1case = s1case + 1
                        s1pending = s1pending + reader.Item("SomtrackID").ToString + vbCrLf
                    ElseIf (reader.Item("StationID") = 2) Then
                        s2case = s2case + 1
                        s2pending = s2pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 3) Then
                        s3case = s3case + 1
                        s3pending = s3pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 4) Then
                        s4case = s4case + 1
                        s4pending = s4pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 5) Then
                        s5case = s5case + 1
                        s5pending = s5pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 6) Then
                        s6case = s6case + 1
                        s6pending = s6pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 7) Then
                        s7case = s7case + 1
                        s7pending = s7pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 8) Then
                        s8case = s8case + 1
                        s8pending = s8pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 9) Then
                        s9case = s9case + 1
                        s9pending = s9pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 10) Then
                        s10case = s10case + 1
                        s10pending = s10pending + reader.Item("SomtrackID").ToString + vbCrLf

                    ElseIf (reader.Item("StationID") = 11) Then
                        s11case = s11case + 1
                        s11pending = s11pending + reader.Item("SomtrackID").ToString + vbCrLf

                    End If
                End While

            End If
        End Using

        If s1pending <> "" Then
            Label31.Text = "➤" + s1pending
        Else
            Label31.Text = s1pending
        End If
        If s2pending <> "" Then
            Label25.Text = "➤" + s2pending
        Else
            Label25.Text = s2pending
        End If
        If s3pending <> "" Then
            Label22.Text = "➤" + s3pending
        Else
            Label22.Text = s3pending
        End If
        If s4pending <> "" Then
            Label13.Text = "➤" + s4pending
        Else
            Label13.Text = s4pending
        End If
        If s5pending <> "" Then
            Label10.Text = "➤" + s5pending
        Else
            Label10.Text = s5pending
        End If
        If s6pending <> "" Then
            Label3.Text = "➤" + s6pending
        Else
            Label3.Text = s6pending
        End If
        If s7pending <> "" Then
            Label4.Text = "➤" + s7pending
        Else
            Label4.Text = s7pending
        End If
        If s8pending <> "" Then
            Label7.Text = "➤" + s8pending
        Else
            Label7.Text = s8pending
        End If
        If s9pending <> "" Then
            Label16.Text = "➤" + s9pending
        Else
            Label16.Text = s9pending
        End If
        If s10pending <> "" Then
            Label19.Text = "➤" + s10pending
        Else
            Label19.Text = s10pending
        End If
        If s11pending <> "" Then
            Label28.Text = "➤" + s11pending
            Label34.Text = "➤" + s11pending
        Else
            Label28.Text = s11pending
            Label34.Text = s11pending
        End If







        con2.Close()
    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        TextBox1.Focus()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetActive()
        GetPending()
        GetPassedCase()



    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Dim s1caseStat = 0
        Dim s2caseStat = 0
        Dim s3caseStat = 0
        Dim s4caseStat = 0
        Dim s5caseStat = 0
        Dim s6caseStat = 0
        Dim s7caseStat = 0
        Dim s8caseStat = 0
        Dim s9caseStat = 0
        Dim s10caseStat = 0
        Dim s11caseStat = 0
        Dim s12caseStat = 0

        Dim s1Print = ""
        Dim s2Print = ""
        Dim s3Print = ""
        Dim s4Print = ""
        Dim s5Print = ""
        Dim s6Print = ""
        Dim s7Print = ""
        Dim s8Print = ""
        Dim s9Print = ""
        Dim s10Print = ""
        Dim s11Print = ""
        Dim s12Print = ""

        s1timer = 0
        s2timer = 0
        s3timer = 0
        s4timer = 0
        s5timer = 0
        s6timer = 0
        s7timer = 0
        s8timer = 0
        s9timer = 0
        s10timer = 0
        s11timer = 0
        s12timer = 0

        '''''''CHECK TABLE CASES'''''''''''
        Dim TableCase As String = "Select PH.SomtrackID, PH.StationID, PH.DateStarted FROM ProductionDetails As PD LEFT JOIN StationProcess As SP On PD.BOMDID = SP.BOMDID LEFT JOIN ProductionHead As PH On PH.ProductionHeadID = PD.ProductionHeadID LEFT JOIN TableMembers As TM On TM.StationID = SP.StationID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID WHERE TS.TableID = @TS And PH.TableNo = @TS And TS.TableSetStatus = 1 And TM.TableMemberStatus = 1 And PD.Status = 2 And PH.StationID Is Not NULL GROUP BY PH.SomtrackID, PH.StationID, PH.DateStarted ORDER BY PH.DateStarted ASC"
        Dim TableCaseQuery As SqlCommand = New SqlCommand(TableCase, con3)
        TableCaseQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = TableCaseQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        s1caseStat = 1
                        s1timer = 0
                        Label33.ForeColor = Color.DarkOrange
                        s1Print = "Idle"

                    ElseIf (reader.Item("StationID") = 2) Then
                        s2caseStat = 1
                        s2timer = 0
                        Label27.ForeColor = Color.DarkOrange
                        s2Print = "Idle"
                    ElseIf (reader.Item("StationID") = 3) Then
                        s3caseStat = 1
                        s3timer = 0
                        Label24.ForeColor = Color.DarkOrange
                        s3Print = "Idle"

                    ElseIf (reader.Item("StationID") = 4) Then
                        s4caseStat = 1
                        s4timer = 0
                        Label15.ForeColor = Color.DarkOrange
                        s4Print = "Idle"

                    ElseIf (reader.Item("StationID") = 5) Then
                        s5caseStat = 1
                        s5timer = 0
                        Label12.ForeColor = Color.DarkOrange
                        s5Print = "Idle"

                    ElseIf (reader.Item("StationID") = 6) Then
                        s6caseStat = 1
                        s6timer = 0
                        Label1.ForeColor = Color.DarkOrange
                        s6Print = "Idle"

                    ElseIf (reader.Item("StationID") = 7) Then
                        s7caseStat = 1
                        s7timer = 0
                        Label6.ForeColor = Color.DarkOrange
                        s7Print = "Idle"

                    ElseIf (reader.Item("StationID") = 8) Then
                        s8caseStat = 1
                        s8timer = 0
                        Label9.ForeColor = Color.DarkOrange
                        s8Print = "Idle"

                    ElseIf (reader.Item("StationID") = 9) Then
                        s9caseStat = 1
                        s9timer = 0
                        Label18.ForeColor = Color.DarkOrange
                        s9Print = "Idle"

                    ElseIf (reader.Item("StationID") = 10) Then
                        s10caseStat = 1
                        s10timer = 0
                        Label21.ForeColor = Color.DarkOrange
                        s10Print = "Idle"

                    ElseIf (reader.Item("StationID") = 11) Then


                        If Label29.Text = "" Then
                            s11caseStat = 1
                            s11timer = 0
                            Label30.ForeColor = Color.DarkOrange
                            s11Print = "Idle"
                        End If
                        If Label35.Text = "" Then
                            s12caseStat = 1
                            s12timer = 0
                            Label36.ForeColor = Color.DarkOrange
                            s12Print = "Idle"
                        End If

                    End If
                End While

            End If
        End Using
        con3.Close()


        '''''''CHECK CASE DURATION ''''''''''
        Dim CaseDuration As String = "Select SP.StationID, DATEDIFF( second, PD.DateStarted, GETDATE() ) As Duration, OD.Duration As TargetDuration FROM [SMProduction].[dbo].[ProductionDetails] As PD LEFT JOIN TableMembers As TM On TM.EmployeeID = PD.EmployeeID LEFT JOIN TableSet As TS On TS.TableSetID = TM.TableSetID LEFT JOIN StationProcess As SP On SP.BOMDID = PD.BOMDID LEFT JOIN BillOfMaterialsDetails As BOMD On BOMD.BOMDID =PD.BOMDID LEFT JOIN ( Select [OperationID], SUM(DATEDIFF(second,'00:00:00',[Duration])) as Duration FROM [SMProduction].[dbo].[OperationsDetail] GROUP BY OperationID ) as OD ON OD.OperationID = BOMD.OperationID WHERE TS.TableSetStatus = 1 AND TM.TableMemberStatus = 1 AND TS.TableID = @TS AND PD.Status = 1 AND SP.StationID = TM.StationID GROUP BY SP.StationID, PD.DateStarted, PD.DateEnded, Duration"
        Dim CaseDurationQuery As SqlCommand = New SqlCommand(CaseDuration, con3)
        CaseDurationQuery.Parameters.AddWithValue("@TS", TableSet)

        con3.Open()
        Using reader As SqlDataReader = CaseDurationQuery.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    If (reader.Item("StationID") = 1) Then
                        s1caseStat = 1
                        Label33.ForeColor = Color.Lime
                        s1Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s1timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 2) Then
                            s2caseStat = 1
                            Label27.ForeColor = Color.Lime
                        s2Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s2timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 3) Then
                            s3caseStat = 1
                            Label24.ForeColor = Color.Lime
                        s3Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s3timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 4) Then
                            s4caseStat = 1
                            Label15.ForeColor = Color.Lime
                        s4Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s4timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 5) Then
                            s5caseStat = 1
                            Label12.ForeColor = Color.Lime
                        s5Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s5timer = 1
                        End If

                    ElseIf (reader.Item("StationID") = 6) Then
                            s6caseStat = 1
                            Label1.ForeColor = Color.Lime
                        s6Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s6timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 7) Then
                            s7caseStat = 1
                            Label6.ForeColor = Color.Lime
                        s7Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s7timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 8) Then
                            s8caseStat = 1
                            Label9.ForeColor = Color.Lime
                        s8Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s8timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 9) Then
                            s9caseStat = 1
                            Label18.ForeColor = Color.Lime
                        s9Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s9timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 10) Then
                            s10caseStat = 1
                            Label21.ForeColor = Color.Lime
                        s10Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s10timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 11) Then
                            s11caseStat = 1
                            Label30.ForeColor = Color.Lime
                        s11Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s11timer = 1
                        End If
                    ElseIf (reader.Item("StationID") = 12) Then
                            s12caseStat = 1
                        Label36.ForeColor = Color.Lime
                        s12Print = TimeSpan.FromSeconds(reader.Item("Duration")).ToString
                        If reader.Item("Duration") > reader.Item("TargetDuration") Then
                            s12timer = 1
                        End If
                    End If
                End While

            End If
        End Using
        con3.Close()






        Dim TableActive = 0
        Dim tabletext = ""
        '''''''QUERY FOR SELECTING ACTIVE TABLE'''''''''''
        Dim TableSetIDquery As String = "SELECT TableSetName FROM [SMProduction].[dbo].[TableSet] WHERE TableID = @TS AND TableSetStatus = 1"
        Dim TableSetIDquerycmd As SqlCommand = New SqlCommand(TableSetIDquery, con3)
        TableSetIDquerycmd.Parameters.AddWithValue("@TS", TableSet)
        con3.Open()
        Using reader As SqlDataReader = TableSetIDquerycmd.ExecuteReader()
            If reader.HasRows Then
                While reader.Read()
                    tabletext = "One Piece Flow - Live Entry   " + reader.Item("TableSetName").ToString
                End While
                TableActive = 1
                TextBox1.Enabled = True
                TextBox1.Focus()
            Else
                TableActive = 0

            End If
        End Using
        con3.Close()

        If TableActive = 1 Then
            Me.Text = tabletext


            If s1caseStat = 1 Then
                Label33.Text = s1Print

            Else
                    Label33.Text = ""
            End If

            If s2caseStat = 1 Then
                Label27.Text = s2Print
            Else
                Label27.Text = ""
            End If

            If s3caseStat = 1 Then
                Label24.Text = s3Print
            Else
                Label24.Text = ""
            End If

            If s4caseStat = 1 Then
                Label15.Text = s4Print

            Else
                Label15.Text = ""
            End If

            If s5caseStat = 1 Then
                Label12.Text = s5Print
            Else
                Label12.Text = ""
            End If

            If s6caseStat = 1 Then
                Label1.Text = s6Print
            Else
                Label1.Text = ""
            End If

            If s7caseStat = 1 Then
                Label6.Text = s7Print
            Else
                Label6.Text = ""
            End If

            If s8caseStat = 1 Then
                Label9.Text = s8Print
            Else
                Label9.Text = ""
            End If

            If s9caseStat = 1 Then
                Label18.Text = s9Print
            Else
                Label18.Text = ""
            End If

            If s10caseStat = 1 Then
                Label21.Text = s10Print
            Else
                Label21.Text = ""
            End If

            If s11caseStat = 1 Then
                Label30.Text = s11Print

            Else
                Label30.Text = ""
            End If

            If s12caseStat = 1 Then
                Label36.Text = s12Print
            Else
                Label36.Text = ""
            End If






        Else
            TextBox1.Enabled = False
            Label33.Text = "Disabled"
            Label27.Text = "Disabled"
            Label24.Text = "Disabled"
            Label15.Text = "Disabled"
            Label12.Text = "Disabled"
            Label1.Text = "Disabled"
            Label6.Text = "Disabled"
            Label9.Text = "Disabled"
            Label18.Text = "Disabled"
            Label21.Text = "Disabled"
            Label30.Text = "Disabled"
            Label36.Text = "Disabled"

            Label33.ForeColor = Color.Red
            Label27.ForeColor = Color.Red
            Label24.ForeColor = Color.Red
            Label15.ForeColor = Color.Red
            Label12.ForeColor = Color.Red
            Label1.ForeColor = Color.Red
            Label6.ForeColor = Color.Red
            Label9.ForeColor = Color.Red
            Label18.ForeColor = Color.Red
            Label21.ForeColor = Color.Red
            Label30.ForeColor = Color.Red
            Label36.ForeColor = Color.Red
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If showwarning = 0 Then
            If s1case > 3 Then

                PictureBox14.Visible = True

            End If

            If s2case > 3 Then

                PictureBox15.Visible = True

            End If
            If s3case > 3 Then

                PictureBox16.Visible = True

            End If
            If s4case > 3 Then

                PictureBox17.Visible = True

            End If
            If s5case > 3 Then

                PictureBox18.Visible = True

            End If
            If s6case > 3 Then

                PictureBox19.Visible = True

            End If
            If s7case > 3 Then

                PictureBox20.Visible = True

            End If
            If s8case > 3 Then

                PictureBox21.Visible = True

            End If
            If s9case > 3 Then

                PictureBox22.Visible = True

            End If
            If s10case > 3 Then

                PictureBox23.Visible = True

            End If
            If s11case > 3 Then

                PictureBox24.Visible = True
                PictureBox25.Visible = True

            End If




            If s1timer = 1 Then

            End If
            If s2timer = 1 Then
                PictureBox27.Visible = True

            End If

            If s3timer = 1 Then
                PictureBox28.Visible = True

            End If

            If s4timer = 1 Then
                PictureBox29.Visible = True

            End If

            If s5timer = 1 Then
                PictureBox30.Visible = True

            End If

            If s6timer = 1 Then
                PictureBox31.Visible = True

            End If

            If s7timer = 1 Then
                PictureBox32.Visible = True

            End If

            If s8timer = 1 Then
                PictureBox33.Visible = True

            End If

            If s9timer = 1 Then
                PictureBox34.Visible = True

            End If

            If s10timer = 1 Then
                PictureBox35.Visible = True

            End If

            If s11timer = 1 Then
                PictureBox36.Visible = True

            End If

            If s12timer = 1 Then
                PictureBox37.Visible = True

            End If
            showwarning = 1
        Else
            PictureBox14.Visible = False
            PictureBox15.Visible = False
            PictureBox16.Visible = False
            PictureBox17.Visible = False
            PictureBox18.Visible = False
            PictureBox19.Visible = False
            PictureBox20.Visible = False
            PictureBox21.Visible = False
            PictureBox22.Visible = False
            PictureBox23.Visible = False
            PictureBox24.Visible = False
            PictureBox25.Visible = False
            PictureBox26.Visible = False
            PictureBox27.Visible = False
            PictureBox28.Visible = False
            PictureBox29.Visible = False
            PictureBox30.Visible = False
            PictureBox31.Visible = False
            PictureBox32.Visible = False
            PictureBox33.Visible = False
            PictureBox34.Visible = False
            PictureBox35.Visible = False
            PictureBox36.Visible = False
            PictureBox37.Visible = False

            showwarning = 0
        End If

    End Sub


End Class
