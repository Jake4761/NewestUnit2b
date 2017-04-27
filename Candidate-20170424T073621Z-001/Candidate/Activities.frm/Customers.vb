Imports System.IO
Public Class Customers

    'Creating a structure with FirstName, lastname ,email , Address , PhoneNumber, DateOfBirth and Session Id
    ' To be used when storing data in a text file

    Private Structure CustomerData
        Dim FirstName As String
        Dim LastName As String
        Dim Email As String
        Dim Address As String
        Dim Phone As String
        Dim DOB As String
        Dim SessionID As Integer
    End Structure




    Private Sub btnInput_Click(sender As System.Object, e As System.EventArgs) Handles btnInput.Click

        'Here i am declaring the variables about the customer and giving them the value from their respective textbox
        'I also define Validation as boolean to be used when i validate the details
        'And SessionID which will be given a value later.
        Dim FirstName As String = txtFirstName.Text
        Dim LastName As String = txtLastName.Text
        Dim Email As String = txtEmail.Text
        Dim Address As String = txtAddress.Text
        Dim Phone As String = txtPhone.Text
        Dim DOB As String = dtpDoB.Text
        Dim SessionID As Integer
        Dim Validation As Boolean = True


        'This is a presence check which is being run on the users personal information bar date of birth
        'If any of the variables have nothing inside of them ("") then the validation identifier will be 
        'changed to false, and a message box informing the user what to do will be displayed.
        If FirstName = "" Or LastName = "" Or Email = "" Or Address = "" Or Phone = "" Then
            Validation = False
            MsgBox("Please ensure all fields are filled in")
        End If



        'This is a length check, i am using the .length method to check whether or not the phone number
        'the user inputted is 11 characters long, this is due to phone numbers being 11 character.
        'If it is not then  the validation identifier will be 
        'changed to false, and a message box informing the user what to do will be displayed.
        If Phone.Length <> 11 Then
            Validation = False
            MsgBox("Please make sure there are 11 numbers in the phone number")
        End If


        'This uses the IsNumeric Function to check if the variable phone is a number, if it is not 
        'Then then the validation identifier will be changed to false, and a message box 
        'informing the user what to do will be displayed.

        If IsNumeric(Phone) = False Then
            MsgBox("Please make sure you have entered number for your phone number ")
            Validation = False
        End If


        'Here VB will try to assign the variable "Testing" as a new MailAddress with the value Email.
        'If it is successful then the email is of correct format and no corrective action will be taken.
        'If it is not however then the validation identifier will be 
        'changed to false, and a message box informing the user what to do will be displayed.
        Try
            Dim Testing As New System.Net.Mail.MailAddress(Email)
        Catch
            Validation = False
            MsgBox("Please make sure you have entered the email in the correct format")
        End Try

        'This If Statement will only run if the validation identifier hasnt been turned to false by the above validation
        If Validation = True Then

            'If the combobox currently has no items in it, meaning there are no records then the first
            'SessionID will take the value of 1000, this is an initial value it will take. It will also 
            'Then be added to the combo box.
            If cmbID.Items.Count = 0 Then
                SessionID = 1000
                cmbID.Items.Add(SessionID)
            Else
                'If There are items in the combo box then the sessionID will take the value of the
                'most recent item in the combo box + 1 meaning it will increment by 1 each time.
                'It is then  added to the combo box.
                'For example if the last SessionID was 1004 then the ne

                SessionID = Val(cmbID.GetItemText(cmbID.Items(cmbID.Items.Count - 1))) + 1
                cmbID.Items.Add(SessionID)
            End If



            'Here i am creating an instance of the structure made at the top
            Dim CustomerDetails As New CustomerData
            'I will be using the streamWriter to write to the text file.
            Dim sw As New StreamWriter(Application.StartupPath & "\Customer.txt", True)
            'This will set the location of each of the items in the text file.
            CustomerDetails.FirstName = LSet(txtFirstName.Text, 50)
            CustomerDetails.LastName = LSet(txtLastName.Text, 50)
            CustomerDetails.Email = LSet(txtEmail.Text, 50)                'Filling the structure with data.5
            CustomerDetails.Address = LSet(txtAddress.Text, 50)
            CustomerDetails.Phone = LSet(txtPhone.Text, 50)
            CustomerDetails.DOB = LSet(dtpDoB.Text, 50)
            CustomerDetails.SessionID = LSet(Val(SessionID), 50)
            'Here it is writing it to the text file
            sw.WriteLine(CustomerDetails.SessionID & CustomerDetails.LastName & CustomerDetails.Email & CustomerDetails.Address & CustomerDetails.Phone & CustomerDetails.DOB & CustomerDetails.FirstName)
            sw.Close()

            MsgBox("Saved")



        End If


    End Sub

    Private Sub btnSearch_Click(sender As System.Object, e As System.EventArgs) Handles btnSearch.Click

        'Here i am declaring and assigning the value of the search they want to search
        Dim SearchValue As String = cmbID.Text

            Dim Found As Boolean = False
            'Found will be an identifyer 
            'This is assigning the value of the text document to CustomerData
            Dim Customerdata() As String = File.ReadAllLines(Dir$("Customer.txt"))
            For I = 0 To UBound(Customerdata)
                'This For loop will repeat for all the lines in the text file
                If Trim(Mid(Customerdata(I), 1, 4)) = cmbID.Text And Not cmbID.Text = "" Then
                    Found = True
                End If

            'The below if statement will find if this line in the text document contains the search value

                If Found = True Then


                txtSession.Text = Trim(Mid(Customerdata(I), 1, 50))
                txtLastName.Text = Trim(Mid(Customerdata(I), 51, 50))
                txtEmail.Text = Trim(Mid(Customerdata(I), 101, 50))
                txtAddress.Text = Trim(Mid(Customerdata(I), 151, 50))
                txtPhone.Text = Trim(Mid(Customerdata(I), 201, 50))
                dtpDoB.Text = Trim(Mid(Customerdata(I), 251, 50)) & " 12:00AM"
                txtFirstName.Text = Trim(Mid(Customerdata(I), 301, 50))
                    'These are all statements that assign the value in the text document line to the textboxes
                    MsgBox("SearchFound")
                    I = UBound(Customerdata)
                    'Above stops the for loops
                End If

            Next I

            If Found = False Then
                MsgBox("Search not found please ensure the value is correct")
            End If

    End Sub

    Private Sub Customers_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

       

        'This will go through the text file at the start of the program and load any
        'exhisting sessionIDs
        Dim IDRead() As String = File.ReadAllLines(Dir$("Customer.txt"))
        For I = 0 To UBound(IDRead)
            cmbID.Items.Add(Trim(Mid(IDRead(I), 1, 4)))
        Next I



    End Sub

    Private Sub btnSwapFormActivities_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSwapFormActivities.Click

        'This button is used to navigate between forms, it will create and show the other form,
        'While hiding this one
        Dim form As New Details
        form.Show()
        Me.Hide()

    End Sub
End Class