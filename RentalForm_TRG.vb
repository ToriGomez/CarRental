'Tori Gomez
'RCET0265
'Spring 2021
'Car Rental Form
'https://github.com/ToriGomez/CarRental_TRG.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm_TRG
    'Disables all Controls when program loads.
    Private Sub RentalForm_TRG_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisabledControls(False)
    End Sub
    'Checks the userinputs making sure that they are vaild before the user continues on.
    Private Sub DaysTextBox_Validated(sender As Object, e As EventArgs) Handles DaysTextBox.Validated
        If Validation() Then
            DisabledControls(False)
        Else
            DisabledControls(True)
            CalculateButton.Select()
        End If
    End Sub
    'Calculates the output of total distance and total charge fees.
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click
        Dim totalMilesDriven As Integer
        Dim miCharge As Double
        Dim daysCharge As Double = (CDbl(DaysTextBox.Text) * 15)
        Dim originalTot As Double
        Dim totalMi As Integer
        Dim totalpaid As Double
        Dim discount() As Double
        discount = {0.05, 0.03, 0.08, 0.95, 0.97, 0.92}

        'Checks the user inputs for every new customer. Allows view of summary once it's clicked.
        If Validation() Then
            DisabledControls(False)
            SummaryButton.Enabled = True
            SummaryToolStripMenuItem1.Enabled = True
        Else
            DisabledControls(True)
            SummaryButton.Enabled = True
            SummaryToolStripMenuItem1.Enabled = True
        End If
        'Conversion to miles on when the rented car's odometer is in kilometers and calculates total.
        If MilesradioButton.Checked = True Then
            totalMilesDriven = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) + "mi"
        ElseIf KilometersradioButton.Checked = True Then
            totalMilesDriven = CInt((CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) * 0.62)
            TotalMilesTextBox.Text = CStr(totalMilesDriven) + "mi"
        End If
        'Saves the calculated total to a new varible to set for the summary's total.
        totalMi = totalMilesDriven
        'Determines the charge fee for distance driven. First 200 is fee, 200 - 500 has $0.12 charge fee/mile,
        '500 plus has $0.10 charge fee/mile.
        Select Case totalMilesDriven
            Case < 200
                miCharge = 0
            Case 200 To 500
                totalMilesDriven -= 200
                miCharge = CDbl(totalMilesDriven) * 0.12
            Case > 500
                miCharge = 36
                totalMilesDriven -= 500
                miCharge += CDbl(totalMilesDriven) * 0.1
        End Select
        'Writes the total Charge of user input in currency to the Mileage Charge output.
        MileageChargeTextBox.Text = FormatCurrency(miCharge, TriState.True)
        'Wries the total Charge of day fee of the amount of rented days. $15/day
        DayChargeTextBox.Text = FormatCurrency(daysCharge, TriState.True)
        'Accumulates the total Charges of mile and day fees.
        originalTot = (miCharge + daysCharge)
        'Acculmulates total discount/s seleceted by the user. Writes to the total Discount amount and 
        'Writes the total Charge fee. 
        If AAAcheckbox.Checked And Not Seniorcheckbox.Checked Then
            TotalDiscountTextBox.Text = FormatCurrency((discount(0) * originalTot), TriState.True)
            TotalChargeTextBox.Text = FormatCurrency((discount(3) * originalTot), TriState.True)
            totalpaid = CDbl(FormatCurrency((discount(3) * originalTot), TriState.True))
        ElseIf Seniorcheckbox.Checked And Not AAAcheckbox.Checked Then
            TotalDiscountTextBox.Text = FormatCurrency((discount(1) * originalTot), TriState.True)
            TotalChargeTextBox.Text = FormatCurrency((discount(3) * originalTot), TriState.True)
            totalpaid = CDbl(FormatCurrency((discount(3) * originalTot), TriState.True))
        ElseIf AAAcheckbox.Checked And Seniorcheckbox.Checked Then
            TotalDiscountTextBox.Text = FormatCurrency((discount(2) * originalTot), TriState.True)
            TotalChargeTextBox.Text = FormatCurrency((discount(5) * originalTot), TriState.True)
            totalpaid = CDbl(FormatCurrency((discount(5) * originalTot), TriState.True))
        Else
            TotalDiscountTextBox.Text = FormatCurrency((0 * originalTot), TriState.True)
            TotalChargeTextBox.Text = FormatCurrency((originalTot), TriState.True)
            totalpaid = CDbl(FormatCurrency((originalTot), TriState.True))
        End If
        'Writes total miles driven and total owed to the summary control for every click of the 
        'calculate control.
        Summary(totalMi, totalpaid, False)
    End Sub
    'Clears all text boxes and restarts the totals for a new customer.
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click
        NameTextBox.Select()
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Checked = True
        CalculateButton.Enabled = False
        CalculateToolStripMenuItem.Enabled = False
    End Sub
    'Displays the summary of the summary from each calculation click made. With how many customers, 
    'how many miles driven in total, and how much has been charged.
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click
        Dim click As Integer = 0
        Summary(click, CDbl(click), True)
    End Sub
    'Exits the program with a double click exit message box. User can change their mind if clicked.
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click
        Dim state As Double
        state = MsgBox("Would you like to exit?", CType(vbYesNo + vbCritical + vbDefaultButton2, MsgBoxStyle), "Exit")
        If state = vbYes Then
            Me.Close()
        End If
    End Sub
    'Dissables or Enables the view of output text boxes and controls.
    Sub DisabledControls(state As Boolean)
        TotalMilesTextBox.Visible = state
        MileageChargeTextBox.Visible = state
        DayChargeTextBox.Visible = state
        TotalDiscountTextBox.Visible = state
        TotalChargeTextBox.Visible = state
        CalculateButton.Enabled = state
        CalculateToolStripMenuItem.Enabled = state
        MilesDrivenLabel.Visible = state
        MileChargeLabel.Visible = state
        DayChargeLabel.Visible = state
        DiscountLabel.Visible = state
        YouOweLabel.Visible = state
        SummaryButton.Enabled = False
        SummaryToolStripMenuItem1.Enabled = False
    End Sub
    'Accumulates the amounts for each time the calculate button is clicked.
    'When Summary button is clicked, displays the totals accumulated and performs
    'the clear button to start a new customers input while continuing to save on the 
    'summary. To clear the summary, User needs to exit the program.
    Sub Summary(distanceTotal As Integer, totalOwed As Double, state As Boolean)
        Static totalCustomers As Integer
        Static totalDistance As Integer
        Static totalCharge As Double
        If state = False Then
            totalCustomers += 1
            totalDistance += distanceTotal
            totalCharge += totalOwed
        Else
            MsgBox($"Total Customers: {totalCustomers}" & vbNewLine _
                    & $"Total Miles Driven: {totalDistance}mi" & vbNewLine _
                    & $"Total Charges: ${totalCharge}")
            ClearButton.PerformClick()
        End If
    End Sub
    'Function for every input text box to check if the inputs are all valid.
    'If a input is invalid, it displays a message box of what needs to be corrected,
    'and flags the user by highlighting the text box yellow and will tab to the 
    'first error seen. If all inputs are valid, the program will continue on.
    Function Validation() As Boolean
        Dim status As Boolean
        Dim selectionTab(1) As Boolean
        Dim groupTab(4) As Boolean
        Dim problemMessage As String = ""
        Dim nameError As Integer
        Dim numberCorrect As Integer
        Dim cityStateError As Integer
        Try
            nameError = CInt(NameTextBox.Text)
            If nameError = CInt(NameTextBox.Text) Or NameTextBox.Text = "" Then
                status = True
                NameTextBox.Text = ""
                problemMessage &= "Name cannot be a number" & vbNewLine
                NameTextBox.BackColor = Color.LightYellow
                NameTextBox.Select()
                selectionTab(0) = True
            End If
        Catch ex As Exception
            If NameTextBox.Text = "" Then
                status = True
                problemMessage &= "Name is required" & vbNewLine
                NameTextBox.BackColor = Color.LightYellow
                NameTextBox.Select()
                selectionTab(0) = True
            Else
                status = False
                NameTextBox.BackColor = Color.White
            End If
        End Try
        If AddressTextBox.Text = "" Then
            status = True
            problemMessage &= "Address is required" & vbNewLine
            AddressTextBox.BackColor = Color.LightYellow
            If Not selectionTab(0) Then
                AddressTextBox.Select()
                selectionTab(1) = True
            End If
        Else
            status = False
            AddressTextBox.BackColor = Color.White
        End If
        Try
            cityStateError = CInt(CityTextBox.Text)
            If cityStateError = CInt(CityTextBox.Text) Or CityTextBox.Text = "" Then
                status = True
                CityTextBox.Text = ""
                problemMessage &= "City cannot be a number" & vbNewLine
                CityTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            If CityTextBox.Text = "" Then
                status = True
                problemMessage &= "City is required" & vbNewLine
                CityTextBox.BackColor = Color.LightYellow
            Else
                status = False
                CityTextBox.BackColor = Color.White
            End If
        End Try
        If Not status And Not selectionTab(0) And Not selectionTab(1) Then
            groupTab(0) = True
        ElseIf status And Not selectionTab(0) And Not selectionTab(1) Then
            CityTextBox.Select()
        End If
        Try
            cityStateError = CInt(StateTextBox.Text)
            If cityStateError = CInt(StateTextBox.Text) Or StateTextBox.Text = "" Then
                status = True
                StateTextBox.Text = ""
                problemMessage &= "State cannot be a number" & vbNewLine
                StateTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            If StateTextBox.Text = "" Then
                status = True
                problemMessage &= "State is required" & vbNewLine
                StateTextBox.BackColor = Color.LightYellow
            Else
                status = False
                StateTextBox.BackColor = Color.White
            End If
        End Try
        If status And groupTab(0) Then
            StateTextBox.Select()
        ElseIf Not status And groupTab(0) Then
            groupTab(1) = True
        End If
        Try
            numberCorrect = CInt(ZipCodeTextBox.Text)
            ZipCodeTextBox.BackColor = Color.White
            If CInt(ZipCodeTextBox.Text) > 99999 Or CInt(ZipCodeTextBox.Text) < 10000 Then
                status = True
                ZipCodeTextBox.Text = ""
                problemMessage &= "Zip Code must be a 5 digit number" & vbNewLine
                ZipCodeTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            status = True
            ZipCodeTextBox.Text = ""
            problemMessage &= "Zip Code must be a numberical value" & vbNewLine
            ZipCodeTextBox.BackColor = Color.LightYellow
        End Try
        If status And groupTab(1) Then
            ZipCodeTextBox.Select()
        ElseIf Not status And groupTab(1) Then
            groupTab(2) = True
        End If
        Try
            numberCorrect = CInt(BeginOdometerTextBox.Text)
            BeginOdometerTextBox.BackColor = Color.White
            If CInt(BeginOdometerTextBox.Text) > CInt(EndOdometerTextBox.Text) Then
                status = True
                BeginOdometerTextBox.Text = ""
                problemMessage &= "Beginning Odometer must be less than Ending Odometer" & vbNewLine
                BeginOdometerTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            status = True
            BeginOdometerTextBox.Text = ""
            problemMessage &= "Beginning Odometer must be a numberical value" & vbNewLine
            BeginOdometerTextBox.BackColor = Color.LightYellow
        End Try
        Try
            numberCorrect = CInt(EndOdometerTextBox.Text)
            EndOdometerTextBox.BackColor = Color.White
            If CInt(EndOdometerTextBox.Text) < 1 Then
                status = True
                EndOdometerTextBox.Text = ""
                problemMessage &= "Ending Odometer must be greater than zero" & vbNewLine
                EndOdometerTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            status = True
            EndOdometerTextBox.Text = ""
            problemMessage &= "Ending Odometer must be a numberical value" & vbNewLine
            EndOdometerTextBox.BackColor = Color.LightYellow
        End Try
        If status And groupTab(2) Then
            BeginOdometerTextBox.Select()
        ElseIf Not status And groupTab(2) Then
            groupTab(3) = True
        End If
        Try
            numberCorrect = CInt(DaysTextBox.Text)
            DaysTextBox.BackColor = Color.White
            If CInt(DaysTextBox.Text) > 45 Or CInt(DaysTextBox.Text) < 1 Then
                status = True
                DaysTextBox.Text = ""
                problemMessage &= "Number of days must be between 1 to 45" & vbNewLine
                DaysTextBox.BackColor = Color.LightYellow
            End If
        Catch ex As Exception
            status = True
            DaysTextBox.Text = ""
            problemMessage &= "Number of days must be a numberical value" & vbNewLine
            DaysTextBox.BackColor = Color.LightYellow
        End Try
        If status And groupTab(3) Then
            DaysTextBox.Select()
        End If
        If Not problemMessage = "" Then
            MsgBox(problemMessage)
        End If
        Return status
    End Function
End Class
