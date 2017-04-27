Imports System.IO
Public Class Details

    Private Structure LeisureData
        Public Type As String
        Public Session As String
        Public Day As String
        Public Location As String                  'Creating the structure that will hold the  data.
        Public Level As String
    End Structure

    Private Sub cmdCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCount.Click
        Dim CountNeeded As Integer
        Dim CountGot As Integer
        Dim I As Integer
        Dim SessionCount As Integer
        SessionCount = 0
        CountNeeded = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If txtType.Text.Length > "30" Then MsgBox("Too many characters in type") : Exit Sub
        If txtSession.Text.Length > "5" Then MsgBox("Too many characters in session") : Exit Sub
        If txtDay.Text.Length > "10" Then MsgBox("Too many characters in day") : Exit Sub ' Making sure that if somebody enters too many characters, instead of returning no results, it warns them.
        If txtLocation.Text.Length > "30" Then MsgBox("Too many characters in location") : Exit Sub
        If TxtLevel.Text.Length > "15" Then MsgBox("Too many characters in level") : Exit Sub
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not txtType.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtSession.Text = "" Then CountNeeded = CountNeeded + 1 ' This section counts how many text boxes have data in them.
        If Not txtDay.Text = "" Then CountNeeded = CountNeeded + 1
        If Not txtLocation.Text = "" Then CountNeeded = CountNeeded + 1
        If Not TxtLevel.Text = "" Then CountNeeded = CountNeeded + 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If CountNeeded = 0 Then MsgBox("Please enter something to count!") : Exit Sub

        Dim Leisuredata() As String = File.ReadAllLines(Dir$("Leisure.txt"))
        For I = 0 To UBound(Leisuredata)
            CountGot = 0
            If Trim(Mid(Leisuredata(I), 1, 30)) = txtType.Text And Not txtType.Text = "" Then CountGot = CountGot + 1
            If Trim(Mid(Leisuredata(I), 31, 5)) = txtSession.Text And Not txtSession.Text = "" Then CountGot = CountGot + 1 'Counting how many attributes follow the search
            If Trim(Mid(Leisuredata(I), 36, 10)) = txtDay.Text And Not txtDay.Text = "" Then CountGot = CountGot + 1 ' . . .
            If Trim(Mid(Leisuredata(I), 46, 30)) = txtLocation.Text And Not txtLocation.Text = "" Then CountGot = CountGot + 1 ' . . .
            If Trim(Mid(Leisuredata(I), 76, 15)) = TxtLevel.Text And Not TxtLevel.Text = "" Then CountGot = CountGot + 1 ' . . .
            If CountGot = CountNeeded Then SessionCount = SessionCount + 1 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''If attributes match, add to the count.
        Next I
        MsgBox(SessionCount & " Sessions have been found! Contact the centre for more information.")
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim LeisureTimeData As New LeisureData
        Dim sw As New System.IO.StreamWriter(Dir$("Leisure.txt"), True)
        LeisureTimeData.Type = LSet(txtType.Text, 30)
        LeisureTimeData.Session = LSet(txtSession.Text, 5)
        LeisureTimeData.Day = LSet(txtDay.Text, 10)                'Filling the structure with data.
        LeisureTimeData.Location = LSet(txtLocation.Text, 30)
        LeisureTimeData.Level = LSet(TxtLevel.Text, 15)

        sw.WriteLine(LeisureTimeData.Type & LeisureTimeData.Session & LeisureTimeData.Day & LeisureTimeData.Location & LeisureTimeData.Level)
        sw.Close()                                                                  'Always need to close afterwards
    End Sub

    Private Sub LeisureSessionsComplete_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Dir$("Leisure.txt") = "" Then
            Dim sw As New StreamWriter(Application.StartupPath & "\Leisure.txt", True)    'This makes sure there is actually a database to enter/read data. If not, it creates a new blank one.
            sw.WriteLine("                                                                                                                                                                                                                                                                                                                                    ")
            sw.Close()
            MsgBox("A new database has been created", vbExclamation, "Warning!")
        End If
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub btnSwapFormCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSwapFormCustomers.Click
        'This button is used to navigate between forms, it will create and show the other form,
        'While hiding this one
        Dim form As New Customers
        form.Show()
        Me.Hide()

    End Sub
End Class
