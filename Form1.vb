Public Class Form1

    Private Sub TextBox1_BorderStyleChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click        'On Clicking the check button in the form through a mouse, the following things happen 
        Dim username As String = TextBox1.Text                ' defining username as a string as the content of TextBox1
        Dim password As String = TextBox2.Text              ' defining password as a string as the content of TextBox2 
        Dim txtout As String                                'Defining txtout as a string which will be used in future as the output in TextBox3 after clicking Check
        Dim var As Integer                        ' Defining Var as Integer
        Dim score As Integer = 0                     ' Similarly defining other variables like score,length
        Dim length As Integer
        Dim FinalScore As Decimal = 0
        Dim lowercount As Integer = 0                        ' initialising the count of different possible
        Dim uppercount As Integer = 0                        ' allowed characters in the password to 0
        Dim intcount As Integer = 0
        Dim charcount As Integer = 0
        txtout = ""
        length = password.Length                         ' length is defined as length of string stored in variable password
        If username.Length = 0 Then             ' if username is not declared an Error Message is shown
            MessageBox.Show("The Username can't be empty", "Error")
            FinalScore = -120                        ' the score is set to -120 
        End If
        If length < 8 Then                                     ' If password length is less than 8, attention message is shown as smaller passwords are usually weaker 
            MessageBox.Show("The Password is smaller than 8 letters", "Attention")
        ElseIf length >= 8 And String.Compare("PASSWORD", password, True) = 0 Then
            FinalScore = -40                                                 ' If "password"(Case Ignored) is set as password, it is a weaker password
            MessageBox.Show("You have set your password as :- " & password, "Error")
        ElseIf String.Compare(username, password, True) = 0 Then              ' If the entered username is set as password, it is again a weak password
            FinalScore = -40
            MessageBox.Show("You are setting your Username as your Password", "Attention") ' Message Box is displayed
        End If
        For var = 0 To length - 1           ' Iterating through the string stored inside password variable
            If CInt(AscW(password(var))) >= 97 And CInt(AscW(password(var))) <= 122 Then      ' ASCII of lowercase alphabets
                lowercount += 1       ' incrementing counts
            ElseIf CInt(AscW(password(var))) >= 65 And CInt(AscW(password(var))) <= 90 Then ' ASCII of Uppercase alphabets
                uppercount += 1
            ElseIf CInt(AscW(password(var))) >= 48 And CInt(AscW(password(var))) <= 57 Then        ' ASCII of Integers
                intcount += 1
            ElseIf CInt(AscW(password(var))) >= 33 And CInt(AscW(password(var))) <= 47 Then         ' Rest all are ASCII of Characters
                charcount += 1
            ElseIf CInt(AscW(password(var))) >= 58 And CInt(AscW(password(var))) <= 64 Then
                charcount += 1
            ElseIf CInt(AscW(password(var))) >= 91 And CInt(AscW(password(var))) <= 96 Then
                charcount += 1
            ElseIf CInt(AscW(password(var))) >= 123 And CInt(AscW(password(var))) <= 127 Then
                charcount += 1
            ElseIf CInt(AscW(password(var))) = 32 Then               ' If space is a character of the password, the password is not accepted
                MessageBox.Show("You have entered a space character in Password", "Error")
                FinalScore = -120          ' A score of -120 is alloted so that the password becomes Invalid
                Exit For
            Else
                MessageBox.Show("Unknown Error occured", "Error")       ' Outside lowercase, uppercase,integers or special characters in the above range, any other character is not allowed
                FinalScore = -120             ' So a score of -120 is alloted so that the password becomes Invalid
                Exit For
            End If
        Next
        If length >= 8 Then
            FinalScore += 20                ' If length is greater than 8 , score of 20 is alloted
        End If
        If lowercount >= 1 Then                ' lowercase characters are weaker , so 15 score is given
            FinalScore += 15
        End If
        If charcount >= 1 Then
            FinalScore += 23                   ' special characters are tougher to guess, so a final score of 23 is added
        End If
        If intcount >= 1 Then
            FinalScore += 22               ' similarly for integers and uppercase characters
        End If
        If uppercount >= 1 Then
            FinalScore += 20
        End If
        If CInt(FinalScore) > 80 Then                        ' Finally Judging the Password Strength in b/w the specified range
            txtout &= "Very Strong"
        ElseIf CInt(FinalScore) > 65 And CInt(FinalScore) <= 80 Then
            txtout &= "Strong"
        ElseIf CInt(FinalScore) > 45 And CInt(FinalScore) <= 65 Then
            txtout &= "Good"
        ElseIf CInt(FinalScore) > 25 And CInt(FinalScore) <= 45 Then
            txtout &= "Weak"                  ' "&" is used for concatenating strings
        ElseIf CInt(FinalScore) >= 0 And CInt(FinalScore) <= 25 Then
            txtout &= "Very Weak"
        Else                                        ' For cases like password having a space(" ") in the string or any character outside the specified range of ASCII is entered , the password becomes invalid 
            txtout &= "Invalid"
        End If
        TextBox3.Text = txtout     ' finally printing the content of txtout in TextBox3
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Restart()                     ' If Restart Button is Clicked , the application restarts 
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()           ' If Exit Button is clicked, the application closes 
    End Sub
End Class
