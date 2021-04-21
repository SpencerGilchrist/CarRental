Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
    Me.close
End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        NameTextBox.Text = ""
        NameTextBox.BackColor = Color.White

        AddressTextBox.Text = ""
        AddressTextBox.BackColor = Color.White

        CityTextBox.Text = ""
        CityTextBox.BackColor = Color.White

        StateTextBox.Text = ""
        StateTextBox.BackColor = Color.White

        ZipCodeTextBox.Text = ""
        ZipCodeTextBox.BackColor = Color.White

        BeginOdometerTextBox.Text = ""
        BeginOdometerTextBox.BackColor = Color.White

        EndOdometerTextBox.Text = ""
        EndOdometerTextBox.BackColor = Color.White

        DaysTextBox.Text = ""
        DaysTextBox.BackColor = Color.White

        TotalMilesTextBox.Text = ""

        MileageChargeTextBox.Text = ""

        DayChargeTextBox.Text = ""

        TotalDiscountTextBox.Text = ""

        TotalChargeTextBox.Text = ""


    End Sub

    'Global Variables

    Dim Days As Integer
    Dim MilesandKilometers As Integer
    Dim MilesDriven As Integer
    Dim Mileage As Integer
    Dim NumberofDays As Integer
    Dim Discount As String
    Dim Total As Integer
    Dim TotalNumberofCustomers As Integer
    Dim TotalNumberOfMiles As Integer
    Dim TotalCharges As Integer

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Validation()

        If Validation() = False Then

        ElseIf Validation() = True Then 'This makes sure that if Validation is False that it wont calculate to prevent crashes

            Calculations()

        End If

        TotalChargeTextBox.Text = ($"{Total}")

        TotalMilesTextBox.Text = ($"{MilesDriven} Miles")

        MileageChargeTextBox.Text = ($"{Mileage}")

        DayChargeTextBox.Text = ($"{Days}")

        TotalDiscountTextBox.Text = Discount

    End Sub

    Public Function Validation() As Boolean

        Dim ErrorMessage As String
        Dim endOdometerReading As Integer
        Dim beginOdometerReading As Integer
        Dim zipCode As Integer

        Try
            NumberofDays = CInt(DaysTextBox.Text)
            DaysTextBox.BackColor = Color.White
        Catch ex As Exception
            DaysTextBox.BackColor = Color.Red
            ErrorMessage &= "Days must be number" & vbNewLine
        End Try

        Try
            endOdometerReading = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            EndOdometerTextBox.BackColor = Color.Red
            ErrorMessage &= "End Odometer must be number" & vbNewLine
        End Try

        Try
            beginOdometerReading = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            BeginOdometerTextBox.BackColor = Color.Red
            ErrorMessage &= "Begin Odometer must be number" & vbNewLine
        End Try

        Try
            zipCode = CInt(ZipCodeTextBox.Text)
        Catch ex As Exception
            ZipCodeTextBox.BackColor = Color.Red
            ErrorMessage &= "Zip Code must be number" & vbNewLine
        End Try

        If StateTextBox.Text = "" Then

            StateTextBox.BackColor = Color.Red
            ErrorMessage &= "State text box is Empty" & vbNewLine

        End If

        If AddressTextBox.Text = "" Then

            AddressTextBox.BackColor = Color.Red
            ErrorMessage &= "Address text box is Empty" & vbNewLine

        End If

        If NameTextBox.Text = "" Then

            NameTextBox.BackColor = Color.Red
            ErrorMessage &= "Name text box is Empty" & vbNewLine

        End If

        If CityTextBox.Text = "" Then

            CityTextBox.BackColor = Color.Red
            ErrorMessage &= "City text box is Empty" & vbNewLine

        End If

        If DaysTextBox.Text = "" Then

            DaysTextBox.BackColor = Color.Red
            ErrorMessage &= "Number of days text box is Empty" & vbNewLine

        End If

        If ErrorMessage = "" Then

            Validation = True

        Else MsgBox(ErrorMessage)

            Validation = False

        End If



    End Function

    Public Function Calculations() As Integer
        '----------------------------------------------------------------------
        If MilesradioButton.Checked Then

            MilesandKilometers = 1

        ElseIf KilometersradioButton.Checked Then

            MilesandKilometers = CInt((1.6) * MilesDriven)

        End If

        MilesDriven = CInt(CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) * CInt(MilesandKilometers)

        NumberofDays = CInt(DaysTextBox.Text)

        Days = NumberofDays * 15
        '---------------------------------------------------------------------

        If MilesDriven < 200 Then

            Mileage = ((MilesDriven) * 0)

        ElseIf MilesDriven > 200 Then

            Mileage = CInt((MilesDriven) * 1.12)

        ElseIf MilesDriven > 500 Then

            Mileage = CInt((MilesDriven) * 1.1)

        Else

        End If
        '-----------------------------------------------------------------------------
        'This Box is the Discount Box and covers all of the discounts seen
        'By the program.
        If AAAcheckbox.Checked Then

            Total = CInt((Days + Mileage) * 0.95)
            Discount = CStr("5%")

        ElseIf Seniorcheckbox.Checked Then

            Total = CInt((Days + Mileage) * 0.97)
            Discount = CStr("3%")

        Else

            Total = CInt(Days + Mileage)
            Discount = CStr("0%")

        End If
        '---------------------------------------------------------------------
        'This is the summary box



    End Function

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click

        TotalNumberofCustomers = 100
        TotalNumberofCustomers = (TotalNumberofCustomers) + 1

        TotalNumberOfMiles = (10000 + Mileage)

        TotalCharges = (1000 + Total)

        MsgBox($"This is our Summary. Total Customers is {TotalNumberofCustomers}, Total number of Miles is {TotalNumberOfMiles}, Total Amount of Charges is {TotalCharges}.")

    End Sub
End Class
