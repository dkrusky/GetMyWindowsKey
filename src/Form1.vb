Public Class Form1
    Const Key As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\"

    Public Overloads ReadOnly Property ProductName As String
        Get
            Return DirectCast(My.Computer.Registry.LocalMachine.OpenSubKey(Key).GetValue("ProductName"), String)
        End Get
    End Property

    Public ReadOnly Property InstallDate As DateTime
        Get
            Dim UnixDate As Integer = DirectCast(My.Computer.Registry.LocalMachine.OpenSubKey(Key).GetValue("InstallDate"), Int32)
            Return New DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc).AddSeconds(UnixDate).ToLocalTime()
        End Get
    End Property

    Public ReadOnly Property DigitalID As Byte()
        Get
            Return DirectCast(My.Computer.Registry.LocalMachine.OpenSubKey(Key).GetValue("DigitalProductId"), Byte())
        End Get
    End Property

    Public ReadOnly Property PID As String
        Get
            Return DirectCast(My.Computer.Registry.LocalMachine.OpenSubKey(Key).GetValue("ProductID"), String)
        End Get
    End Property

    Public ReadOnly Property ProductKey As String
        Get
            Try
                Dim pKey As Byte() = DigitalID
                Dim Chars As String = "BCDFGHJKMPQRTVWXY2346789"
                Dim i As Integer = 24
                Dim isWin8 As Integer = (pKey(66) \ 6) And 1
                Dim Cur As Integer = 0
                Dim x As Integer = 14
                Dim Last As Integer = 0
                Dim keypart1 As String = ""

                Dim insert As String = ""
                Dim a As String = ""
                Dim b As String = ""
                Dim c As String = ""
                Dim d As String = ""
                Dim e As String = ""
                Dim KeyOutput As String = ""

                Const KeyOffset = 52
                pKey(66) = (pKey(66) And &HF7) Or ((isWin8 And 2) * 4)
                i = 24
                Chars = "BCDFGHJKMPQRTVWXY2346789"
                Do
                    Cur = 0
                    x = 14
                    Do
                        Cur = Cur * 256
                        Cur = pKey(x + KeyOffset) + Cur
                        pKey(x + KeyOffset) = (Cur \ 24)
                        Cur = Cur Mod 24
                        x = x - 1
                    Loop While x >= 0
                    i = i - 1
                    KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
                    Last = Cur
                Loop While i >= 0
                If (isWin8 = 1) Then
                    keypart1 = Mid(KeyOutput, 2, Last)
                    insert = "N"
                    KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
                    If Last = 0 Then KeyOutput = insert & KeyOutput
                End If
                a = Mid(KeyOutput, 1, 5)
                b = Mid(KeyOutput, 6, 5)
                c = Mid(KeyOutput, 11, 5)
                d = Mid(KeyOutput, 16, 5)
                e = Mid(KeyOutput, 21, 5)
                Return a & "-" & b & "-" & c & "-" & d & "-" & e
            Catch er As Exception
                Return er.Message
            End Try
        End Get
    End Property

    Private Sub Donate(Email As String, Title As string)
			Dim url As string="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business={0}&lc=US&item_name={1}&no_note=0&cn=Feedback&no_shipping=1&currency_code=USD&bn=PP-DonationsBF:btn_donate_LG.gif:NonHosted"
			System.Diagnostics.Process.Start(String.Format(url, Email, Title))
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = Label2.Text
        TextBox1.Text = ProductKey
        TextBox2.Text = ProductName
        TextBox3.Text = PID
        Label6.Text = "Windows was installed on " & InstallDate.ToString("MMMM dd, yyyy")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Donate("45KFL2X2YESL6","GetMyWindowsKey Donation")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Application.Exit()
        'Button1.PerformClick()
    End Sub

    Dim IsDraggingForm As Boolean = False
    Private MousePos As New System.Drawing.Point(0, 0)

    Private Sub title_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Panel1.MouseDown, Label2.MouseDown
        If e.Button = MouseButtons.Left Then
            IsDraggingForm = True
            MousePos = e.Location
        End If
    End Sub

    Private Sub title_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Panel1.MouseUp, Label2.MouseUp
        If e.Button = MouseButtons.Left Then IsDraggingForm = False
    End Sub

    Private Sub title_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Panel1.MouseMove, Label2.MouseUp
        If IsDraggingForm Then
            Dim temp As Point = New Point(Me.Location + (e.Location - MousePos))
            Me.Location = temp
            temp = Nothing
        End If
    End Sub
End Class
