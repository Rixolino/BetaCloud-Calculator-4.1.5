Imports Microsoft.VisualBasic
Imports System.OverflowException


Public Class Inequalities
    Public MoveForm As Boolean
    Public MoveForm_MousePosition As Point
    Dim m, f, t
    Public Sub Perf(ByRef f As Size)

    End Sub
    Public Sub New()
        InitializeComponent()

        Me.DoubleBuffered = True
        Me.SetStyle(ControlStyles.ResizeRedraw, True)

    End Sub

    Private Const cGrip As Integer = 16
    Private Const cCaption As Integer = 32

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H84 Then
            Dim pos As Point = New Point(m.LParam.ToInt32())
            pos = Me.PointToClient(pos)

            If pos.Y < cCaption Then
                m.Result = CType(2, IntPtr)
                Return
            End If

            If pos.X >= Me.ClientSize.Width - cGrip AndAlso pos.Y >= Me.ClientSize.Height - cGrip Then
                m.Result = CType(17, IntPtr)
                Return
            End If
        End If



        MyBase.WndProc(m)
    End Sub


    Public Sub MoveForm_MouseDown(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseDown ' Add more handles here (Example: PictureBox1.MouseDown)

        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.NoMove2D
            MoveForm_MousePosition = e.Location
        End If

    End Sub

    Public Sub MoveForm_MouseMove(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseMove ' Add more handles here (Example: PictureBox1.MouseMove)

        If MoveForm Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If

    End Sub

    Public Sub MoveForm_MouseUp(sender As Object, e As MouseEventArgs) Handles _
    MyBase.MouseUp ' Add more handles here (Example: PictureBox1.MouseUp)

        If e.Button = MouseButtons.Left Then
            MoveForm = False
            Me.Cursor = Cursors.Default
        End If

    End Sub

    Private Sub Inequalities_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximumSize = Screen.FromRectangle(Me.Bounds).WorkingArea.Size
    End Sub

    Function dec2frac(ByVal dblDecimal As Double) As String
        Dim intNumerator, intDenominator, intNegative As Integer
        Dim dblFraction, dblAccuracy As Double
        Dim txtDecimal As String

        dblAccuracy = 0.1                                       ' Set the initial Accuracy level.
        txtDecimal = dblDecimal.ToString                        ' Get a  string representation of the input number.

        For i As Integer = 0 To (txtDecimal.Length - 1)                                 ' Check each character to see if it's a decimal point...
            If txtDecimal.Substring(i, 1) = "." Then                    ' if it is then we get the number of digits behind the decimal
                dblAccuracy = 1 / 10 ^ (txtDecimal.Length - i)    '   assign the new accuracy level, and
                Exit For                                            '   exit the for loop.
            End If
        Next
        intNumerator = 0                                ' Set the initial numerator value to 0.
        intDenominator = 1                              ' Set the initial denominator value to 1.
        intNegative = 1                                 ' Set the negative value flag to positive.
        If dblDecimal < 0 Then
            intNegative = -1 ' If the desired decimal value is negative,
        End If

        dblFraction = 0                                 ' Set the fraction value to be 0/1.

        Do While Math.Abs(dblFraction - dblDecimal) > dblAccuracy   ' As long as we're still outside the
            '   desired accuracy, then...
            If Math.Abs(dblFraction) > Math.Abs(dblDecimal) Then      ' If our fraction is too big,
                intDenominator += 1         '   increase the denominator
            Else                                            ' Otherwise
                intNumerator += intNegative   '   increase the numerator.
            End If

            dblFraction = intNumerator / intDenominator     ' Set the new value of the fraction.

        Loop

        Return intNumerator.ToString & "/" & intDenominator.ToString ' Display the numerator and denominator
    End Function
    Function num(ByVal dblDecimal As Double) As String
        Dim intNumerator, intDenominator, intNegative As Integer
        Dim dblFraction, dblAccuracy As Double
        Dim txtDecimal As String

        dblAccuracy = 0.1                                       ' Set the initial Accuracy level.
        txtDecimal = dblDecimal.ToString                        ' Get a  string representation of the input number.

        For i As Integer = 0 To (txtDecimal.Length - 1)                                 ' Check each character to see if it's a decimal point...
            If txtDecimal.Substring(i, 1) = "." Then                    ' if it is then we get the number of digits behind the decimal
                dblAccuracy = 1 / 10 ^ (txtDecimal.Length - i)    '   assign the new accuracy level, and
                Exit For                                            '   exit the for loop.
            End If
        Next
        intNumerator = 0                                ' Set the initial numerator value to 0.
        intDenominator = 1                              ' Set the initial denominator value to 1.
        intNegative = 1                                 ' Set the negative value flag to positive.
        If dblDecimal < 0 Then
            intNegative = -1 ' If the desired decimal value is negative,
        End If

        dblFraction = 0                                 ' Set the fraction value to be 0/1.

        Do While Math.Abs(dblFraction - dblDecimal) > dblAccuracy   ' As long as we're still outside the
            '   desired accuracy, then...
            If Math.Abs(dblFraction) > Math.Abs(dblDecimal) Then      ' If our fraction is too big,
                intDenominator += 1         '   increase the denominator
            Else                                            ' Otherwise
                intNumerator += intNegative   '   increase the numerator.
            End If

            dblFraction = intNumerator / intDenominator     ' Set the new value of the fraction.

        Loop

        Return intNumerator.ToString
    End Function
    Function den(ByVal dblDecimal As Double) As String
        Dim intNumerator, intDenominator, intNegative As Integer
        Dim dblFraction, dblAccuracy As Double
        Dim txtDecimal As String

        dblAccuracy = 0.1                                       ' Set the initial Accuracy level.
        txtDecimal = dblDecimal.ToString                        ' Get a  string representation of the input number.

        For i As Integer = 0 To (txtDecimal.Length - 1)                                 ' Check each character to see if it's a decimal point...
            If txtDecimal.Substring(i, 1) = "." Then                    ' if it is then we get the number of digits behind the decimal
                dblAccuracy = 1 / 10 ^ (txtDecimal.Length - i)    '   assign the new accuracy level, and
                Exit For                                            '   exit the for loop.
            End If
        Next
        intNumerator = 0                                ' Set the initial numerator value to 0.
        intDenominator = 1                              ' Set the initial denominator value to 1.
        intNegative = 1                                 ' Set the negative value flag to positive.
        If dblDecimal < 0 Then
            intNegative = -1 ' If the desired decimal value is negative,
        End If

        dblFraction = 0                                 ' Set the fraction value to be 0/1.

        Do While Math.Abs(dblFraction - dblDecimal) > dblAccuracy   ' As long as we're still outside the
            '   desired accuracy, then...
            If Math.Abs(dblFraction) > Math.Abs(dblDecimal) Then      ' If our fraction is too big,
                intDenominator += 1         '   increase the denominator
            Else                                            ' Otherwise
                intNumerator += intNegative   '   increase the numerator.
            End If

            dblFraction = intNumerator / intDenominator     ' Set the new value of the fraction.

        Loop

        Return intDenominator.ToString
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label6.Text = "Δ = " & (b.Text ^ 2) & "-" & "(" & 4 * (a.Text * c.Text) & ") = "

        copies.Show()

        delta.Text = (b.Text ^ 2) - 4 * (a.Text * c.Text)
        'ex: a>0 (concordant)
        If (a.Text > 0) Then If (sign.Text = ">") Then info.Text = "x < " & r1.Text & " v x > " & r2.Text
        'ex: +a<0 (discordant)
        If (a.Text > 0) Then If (sign.Text = "<") Then info.Text = r1.Text & "< x <" & r2.Text Else If (delta.Text < 0) Then info.Text = "∀x∈R"
        'ex: -a<0 (concordant)
        If (a.Text < 0) Then If (sign.Text = "<") Then info.Text = "x < " & r1.Text & " v x > " & r2.Text
        'ex: -a>0 (discordant)
        If (a.Text < 0) Then If (sign.Text = ">") Then info.Text = r1.Text & " < x < " & r2.Text Else If (delta.Text < 0) Then info.Text = "∀x∈R"

        If (delta.Text = 0) Then info.Text = "x≠" & r1.Text
        If (info.Text = "NaN< x <NaN") Then
            info.Text = "∄x∈R"
        End If

        n1.Text = b.Text * -1

        dt.Text = delta.Text

        dn.Text = 2 * a.Text

        r1.Text = (((b.Text * -1) - Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text))) / (2 * a.Text))

        r2.Text = ((b.Text * -1) + Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text))) / (2 * a.Text)

        If (info.Text = "< x <") Then
            info.Text = ""
            MsgBox("Click again the Calculate button")
        End If
        If (info.Text = "x < " & " v x > ") Then
            info.Text = ""
            MsgBox("Click again the Calculate button")
        End If
        If (info.Text = "x < NaN v x > NaN") Then
            info.Text = ""
            MsgBox("Click again the Calculate button")
        End If
        If ((b.Text * -1) + Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text))) / (2 * a.Text) < ((b.Text * -1) - Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text) / (2 * a.Text))) Then
            r1.Text = ((b.Text * -1) + Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text))) / (2 * a.Text)
            r2.Text = ((b.Text * -1) - Math.Sqrt(b.Text ^ 2 - 4 * (a.Text * c.Text))) / (2 * a.Text)
        End If
    End Sub

    Private Sub Closea_Click(sender As Object, e As EventArgs) Handles Closea.Click
        Application.Exit()
    End Sub

    Private Sub Minimizea_Click(sender As Object, e As EventArgs) Handles Minimizea.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.Show()
    End Sub

    Private Sub NormalActualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalActualToolStripMenuItem.Click
        Me.Close()
        Normal.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        copies.Hide()
        a.Text = ""
        b.Text = ""
        c.Text = ""
        delta.Text = ""
        n1.Text = "-b"
        dt.Text = "b^2 - 4ac"
        dn.Text = "2a"
        r1.Text = ""
        r2.Text = ""
        Label6.Text = "Δ = b^2 - 4ac = "
        info.Text = ""
        retta.Text = ""

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        sign.Text = "<"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        sign.Text = ">"
    End Sub

    Private Sub ProgramerToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()
        Programmer.Show()
    End Sub

    Private Sub BinaryAdditionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()
        BinaryAdd.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub copies_Click(sender As Object, e As EventArgs) Handles copies.Click
        My.Computer.Clipboard.SetText(a.Text & "x²+(" & b.Text & ")x +(" & c.Text & ")" & sign.Text & "0" & vbCrLf & vbCrLf & info.Text)
        MsgBox("Copied with success!")
    End Sub

    Private Sub RulesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RulesToolStripMenuItem.Click
        rInequalities.Show()
    End Sub

    Private Sub FastMenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FastMenuToolStripMenuItem.Click
        FastMenu.Show()
    End Sub

    Private Sub info_Click(sender As Object, e As EventArgs) Handles info.Click
        If (info.Text = "∄x∈R") Then
            MsgBox("This equation won't have real solutions")
        End If
        If (info.Text = "∀x∈R") Then
            MsgBox("This equation will have real solutions")
        End If
        If (info.Text = (r1.Text & "< x <" & r2.Text)) Then
            MsgBox("Discordant signs: internal solutions")
        End If
        If (info.Text = ("x < " & r1.Text & " v x > " & r2.Text)) Then
            MsgBox("Concordant signs: external solutions")
        End If
        If (info.Text = ("x≠" & r1.Text)) Then
            MsgBox("All real solutions except " & r1.Text)
        End If
    End Sub

    Private Sub ScientificToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScientificToolStripMenuItem.Click
        Me.Close()
        Scientific.Show()
    End Sub
End Class