Imports System.Data.OleDb
Imports System.Windows.Forms


Public Class Student_Login_Window
    Dim Student_ID As String = ""
    Dim Student_Name As String = ""
    Dim Tardy_Num As Integer = 0
    Dim Time As String = ""
    Dim timer As Integer
    Dim mError As Boolean = False
    Dim Print As Boolean = False

    Private Sub Student_Login_Window_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''pops the window to the center of the screen
        CenterToScreen()

        ''Set the starting page to be the first panel to show
        Starting_Page.Visible = True
        ''Disables the other two panels on start
        Tardy_Page.Visible = False
        Early_Dismissal_Page.Visible = False
        End_Label2.Visible = False
        Label3.Visible = False
        AddNew_Panel.Visible = False
    End Sub

    Private Sub Tardy_button_Click_1(sender As Object, e As EventArgs) Handles Tardy_button.Click
        ''Sets the Student_ID variable to the ID that the user entered
        Student_ID = User_Input_box.Text
        Tardy()

    End Sub

    Private Sub Tardy()
        If (check_database()) Then
            Tardy_Page.Visible = True
            Starting_Page.Visible = False
            Tardy_ID_Label2.Text = Student_ID
            Tardy_Name_Label2.Text = Student_Name
            Tardy_Time_Label2.Text = Time
            Error_Label.Text = ""
        End If
    End Sub

    Private Sub Print_Button_Click(sender As Object, e As EventArgs) Handles Print_Button.Click
        Submit_Entry(True)
        If Print = True Then
            Print_pass()
        End If
    End Sub
    Private Sub Submit_Entry(Tardy As Boolean)
        Dim connection As New OleDbConnection
        Dim File_Path As String = "Student Database/report_for_" + DateString + ".accdb"
        ''checks to see if the report data base file exists already
        ''if it does then it creates a new entry
        ''else if creates a new file and then creates and entry
        If System.IO.File.Exists(File_Path) Then
            Console.WriteLine("File Exists.")
            connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student Database/report_for_" + DateString + ".accdb")
        Else
            Console.WriteLine("File Doesnt Exist.")
            Dim new_report As New ADOX.Catalog()
            ''creates file
            new_report.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student DataBase/report_for_" + DateString + ".accdb")
            new_report = Nothing
            Console.WriteLine("File created")
            ''connects to file
            connection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student Database/report_for_" + DateString + ".accdb")
            ''opens file
            connection.Open()
            Dim cmd As New OleDbCommand("", connection)
            ''creates table for report
            cmd.CommandText = "CREATE TABLE Report ([Student_ID] CHAR, [Student_Name] CHAR, [Current_Date] CHAR, [Current_TimeIn] CHAR, [Current_TimeOut] CHAR, [Reason(Early Dismissal)] CHAR, [Tardy/3] CHAR);"
            cmd.ExecuteNonQuery()
            ''adds extra fields
            cmd.CommandText = "ALTER TABLE Report ADD COLUMN Tardy_Type TEXT(25)"
            cmd.ExecuteNonQuery()
            ''closes the connection
            connection.Close()
        End If
        Create_Entry(connection, Tardy)
        ''sets everything to defualt
        If (Not mError) Then
            If Tardy Then
                Tardy_ID_Label2.Text = ""
                Tardy_Name_Label2.Text = ""
                Tardy_Time_Label2.Text = ""
                Tardy_Page.Visible = False
                Label3.Visible = True
                For Each radio In Tardy_Options_Group.Controls
                    If TypeOf radio Is RadioButton Then
                        If radio.checked Then
                            radio.checked = False
                        End If
                    End If
                Next
            Else
                Early_Dismissal_Page.Visible = False
                Early_StudentID_Label2.Text = ""
                Early_Student_Name_Label2.Text = ""
                Time_Label2.Text = ""
                End_Label2.Visible = True
                For Each radio In Early_Options.Controls
                    If TypeOf radio Is RadioButton Then
                        If radio.checked Then
                            radio.checked = False
                        End If
                    End If
                Next
            End If
            timer = 6.0
            Timer1.Start()
        End If
    End Sub

    Private Sub Print_pass()
        Dim Tardy_pass As DYMO.Label.Framework.Label
        ''opens the tardy pass label template
        Tardy_pass = DYMO.Label.Framework.Label.Open("Student DataBase\Tardy_Pass.label")
        ''sets the text objects on the label to the correct name, date, and time
        Tardy_pass.SetObjectText("Nametxt", Student_Name)
        Tardy_pass.SetObjectText("Timetxt", DateString + " " + Time)
        ''sends the label to the printer (comment out if dont have printer)
        Tardy_pass.Print("DYMO LabelWriter 450 Twin Turbo")
        Student_Name = ""
        Student_ID = ""
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ''if the timer hits zero it takes you back to the start
        If timer = 0 Then
            Timer1.Stop()
            Starting_Page.Visible = True
            Label3.Visible = False
            End_Label2.Visible = False
            Error_Label.Text = ""
            AddNew_Button.Visible = False
        Else
            ''else if just subtracts 1 from the timer every second
            timer = timer - 1
        End If
    End Sub

    Private Sub Create_Entry(connection As OleDbConnection, Tardy As Boolean)
        connection.Open()
        Dim command_string As String
        Dim Reason As String = "N/A"
        Dim Tardy_N As Integer = 0
        Dim TardyType As String = "N/A"
        If (Not Tardy) Then
            For Each radio In Early_Options.Controls
                If TypeOf radio Is RadioButton Then
                    If radio.checked Then
                        If (radio.name = "Other_RadioButton") Then
                            mError = False
                            Reason = Other_Textbox.Text
                            Exit For
                        Else
                            mError = False
                            Reason = radio.Text
                            Exit For
                        End If
                    Else
                        mError = True
                    End If
                End If
            Next
        ElseIf Tardy Then
            For Each radio In Tardy_Options_Group.Controls
                If TypeOf radio Is RadioButton Then
                    If radio.checked Then
                        mError = False
                        TardyType = radio.Text
                        Exit For
                    Else
                        mError = True
                    End If
                End If
            Next
            Tardy_Num += 1
            Dim Tableconnect As New OleDbConnection
            ''connects the variable to the data base (accel)
            Tableconnect.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student Database/Student_Data.accdb")
            ''opens data base (accel) with all the students in it
            Tableconnect.Open()

            Dim UpdateStr As String
            ''gets the information where the ID number matches what the user put in
            UpdateStr = " Update [Student_Info] set [Tardy_Num] = '" & Val(Tardy_Num) & "' where [Student_ID] = '" & Student_ID & "'"
            ''creates the command grab the info connected to the ID
            Dim cmd As New OleDbCommand(UpdateStr, Tableconnect)
            cmd.ExecuteNonQuery()
            Tableconnect.Close()
            If Tardy_Num Mod 3 = 0 Then
                Tardy_N = Tardy_Num
            End If
        End If
        If (mError) Then
            Early_error.Visible = True
            Tardy_Errror_Label.Visible = True
            Print = False
        Else
            ''command string to insert data into the table
            command_string = "Insert Into Report ([Student_ID], [Student_Name], [Current_Date], [Current_TimeIn], [Current_TimeOut], [Reason(Early Dismissal)], [Tardy/3], [Tardy_Type]) Values (?,?,?,?,?,?,?,?)"
            Dim cmd2 As New OleDbCommand(command_string, connection)
            ''adds the values to the certain values in the table
            cmd2.Parameters.AddWithValue("Student_ID", Student_ID)
            cmd2.Parameters.AddWithValue("Student_Name", Student_Name)
            cmd2.Parameters.AddWithValue("Current_Date", DateString)
            If (Tardy) Then
                cmd2.Parameters.AddWithValue("Current_TimeIn", Time)
                cmd2.Parameters.AddWithValue("Current_TimeOut", "N/A")
                cmd2.Parameters.AddWithValue("Reason(Early Dismissal)", Reason)
                cmd2.Parameters.AddWithValue("Tardy/3", Tardy_N)
                cmd2.Parameters.AddWithValue("Tardy_Type", TardyType)
                Print = True
            Else
                cmd2.Parameters.AddWithValue("Current_TimeIn", "N/A")
                cmd2.Parameters.AddWithValue("Current_TimeOut", Time)
                cmd2.Parameters.AddWithValue("Reason(Early Dismissal)", Reason)
                cmd2.Parameters.AddWithValue("Tardy/3", Tardy_N)
                cmd2.Parameters.AddWithValue("Tardy_Type", TardyType)
                Student_Name = ""
                Student_ID = ""
            End If
            Early_error.Visible = False
            Tardy_Errror_Label.Visible = False
            cmd2.ExecuteNonQuery()
        End If
        connection.Close()
    End Sub
    ''makes it so the window cannot be moved
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_NCLBUTTONDOWN As Integer = 161
        Const WM_SYSCOMMAND As Integer = 274
        Const HTCAPTION As Integer = 2
        Const SC_MOVE As Integer = 61456

        If (m.Msg = WM_SYSCOMMAND) And (m.WParam.ToInt32() = SC_MOVE) Then
            Return
        End If

        If (m.Msg = WM_NCLBUTTONDOWN) And (m.WParam.ToInt32() = HTCAPTION) Then
            Return
        End If

        MyBase.WndProc(m)
    End Sub

    Private Sub Early_Dismissal_Button_Click(sender As Object, e As EventArgs) Handles Early_Dismissal_Button.Click
        Student_ID = User_Input_box.Text
        Early_dismissal()
    End Sub

    Private Sub Submit_Button_Click(sender As Object, e As EventArgs) Handles Submit_Button.Click
        Submit_Entry(False)
    End Sub

    Private Sub Early_dismissal()
        If (check_database()) Then
            Early_Dismissal_Page.Visible = True
            Starting_Page.Visible = False
            Early_StudentID_Label2.Text = Student_ID
            Early_Student_Name_Label2.Text = Student_Name
            Time_Label2.Text = Time
            Error_Label.Text = ""
        End If
    End Sub

    Private Function check_database() As Boolean
        ''Checks to see if the user entered an ID number
        If Student_ID = "" Then
            Error_Label.Text = "No ID Entered. Please Enter an ID."
            AddNew_Button.Visible = False
            Return False
        Else
            Dim Tableconnect As New OleDbConnection
            ''connects the variable to the data base (accel)
            Tableconnect.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student Database/Student_Data.accdb")
            ''opens data base (accel) with all the students in it
            Tableconnect.Open()

            Dim selectstr As String
            ''gets the information where the ID number matches what the user put in
            selectstr = " Select * From Student_Info where Student_ID= ?"
            ''creates the command grab the info connected to the ID
            Dim cmd As New OleDbCommand(selectstr, Tableconnect)
            ''adds the ID number to command so it knows what to look for
            cmd.Parameters.AddWithValue("Student_ID", Student_ID)

            Dim Read As OleDbDataReader
            ''Initialize reader
            Read = cmd.ExecuteReader
            ''Checks to see if the data base has any data in it
            If (Read.HasRows) Then
                ''Reads the data base for the ID value and returns all the information connect with that ID
                Read.Read()
                ''Gets the First and Last name associated with the ID andd sets the public variable Student_Name to that value
                Student_Name = Read("First_Name").ToString + " " + Read("Last_Name").ToString
                If (IsDBNull(Read("Tardy_Num"))) Then
                    Tardy_Num = 0
                Else
                    Tardy_Num = Read("Tardy_Num")
                End If
                User_Input_box.Text = ""
                Time = DateTime.Now.ToString("hh:mm tt")
                Return True
            Else
                ''If it doesnt find the ID
                Error_Label.Text = "ID Does not exist." & vbNewLine & "Please try again or click add to add a new student to database."
                User_Input_box.Text = ""
                AddNew_Button.Visible = True
                Return False
            End If
            ''Closes the connection to the data base
            Tableconnect.Close()
        End If
    End Function

    Private Sub Other_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles Other_RadioButton.CheckedChanged
        If Other_RadioButton.Checked Then
            Other_Textbox.Visible = True
        Else
            Other_Textbox.Visible = False
        End If
    End Sub

    Private Sub AddNew_Button_Click(sender As Object, e As EventArgs) Handles AddNew_Button.Click
        AddNew_Panel.Visible = True
        Starting_Page.Visible = False
    End Sub

    Private Sub AddNewAddButton_Click(sender As Object, e As EventArgs) Handles AddNewAddButton.Click
        Dim command_string As String
        If AddNewFirstBox.Text = "" Or AddNewIDBox.Text = "" Or AddNewLastBox.Text = "" Then
            AddNewErrorLabel.Visible = True
            AddNewErrorLabel.Text = "Error. Missing data. Please fill in the missing data."
        Else
            Dim Tableconnect As New OleDbConnection
            ''connects the variable to the data base (accel)
            Tableconnect.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Student Database/Student_Data.accdb")
            ''opens data base (accel) with all the students in it
            Tableconnect.Open()

            Dim selectstr As String
            ''gets the information where the ID number matches what the user put in
            selectstr = " Select * From Student_Info where Student_ID= ?"
            ''creates the command grab the info connected to the ID
            Dim cmd As New OleDbCommand(selectstr, Tableconnect)
            ''adds the ID number to command so it knows what to look for
            cmd.Parameters.AddWithValue("Student_ID", AddNewIDBox.Text)

            Dim Read As OleDbDataReader
            ''Initialize reader
            Read = cmd.ExecuteReader
            ''Checks to see if the data base has any data in it
            If (Read.HasRows) Then
                AddNewErrorLabel.Text = "Error. ID already exists." & vbNewLine & "Please enter a new ID or hit cancel to go back."
                AddNewErrorLabel.Visible = True
            Else
                ''command string to insert data into the table
                command_string = "Insert Into Student_Info ([Student_ID], [Last_Name], [First_Name]) Values (?,?,?)"
                Dim cmd2 As New OleDbCommand(command_string, Tableconnect)
                ''adds the values to the certain values in the table
                cmd2.Parameters.AddWithValue("Student_ID", AddNewIDBox.Text)
                cmd2.Parameters.AddWithValue("Last_Name", AddNewLastBox.Text)
                cmd2.Parameters.AddWithValue("First_Name", AddNewFirstBox.Text)
                cmd2.ExecuteNonQuery()
                Tableconnect.Close()
                Starting_Page.Visible = True
                Error_Label.Text = "Data Added"
                AddNewErrorLabel.Visible = False
                AddNew_Panel.Visible = False
                AddNewIDBox.Text = ""
                AddNewLastBox.Text = ""
                AddNewFirstBox.Text = ""
                AddNew_Button.Visible = False
            End If
            Tableconnect.Close()
        End If
    End Sub

    Private Sub AddNewReturnButton_Click(sender As Object, e As EventArgs) Handles AddNewReturnButton.Click
        Starting_Page.Visible = True
        AddNewErrorLabel.Visible = False
        AddNew_Panel.Visible = False
        AddNewIDBox.Text = ""
        AddNewLastBox.Text = ""
        AddNewFirstBox.Text = ""
        Error_Label.Text = ""
        AddNew_Button.Visible = False
    End Sub

    Private Sub ExitTardy_Click(sender As Object, e As EventArgs) Handles ExitTardy.Click
        Starting_Page.Visible = True
        Tardy_Page.Visible = False
        Tardy_Errror_Label.Visible = False
        Student_ID = ""
        Student_Name = ""
    End Sub

    Private Sub ExitEarly_Click(sender As Object, e As EventArgs) Handles ExitEarly.Click
        Starting_Page.Visible = True
        Early_Dismissal_Page.Visible = False
        Early_error.Visible = False
        Student_ID = ""
        Student_Name = ""
    End Sub
End Class