Public Class Form1
    Dim EcoSeat(99), BusiSeat(19) As RichTextBox
    Dim EcoSeatClicked As Boolean = False
    Dim BusiSeatClicked As Boolean = False
    Dim EcoRowColInput As Boolean = False
    Dim BusiRowColInput As Boolean = False
    Dim IndexNumOfEcoAry, IndexNumOfBusiAry As Integer
    Dim EcoSeatRowNum, EcoSeatColNum, BusiSeatRowNum, BusiSeatColNum As Integer
    Dim Success As Boolean = False
    Dim ConnOBJ As New ADODB.Connection
    Dim RecSetEco, RecSetBusi As New ADODB.Recordset

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConnOBJ.Provider = "Microsoft.jet.oledb.4.0"
        ConnOBJ.ConnectionString = "C:\Users\jlwan\source\repos\vb\WindowsFormApp.NetFramework.vb\AirLineDB\database\airlinedb.mdb"
        ConnOBJ.Open()
        RecSetEco.Open("Select * FROM EcoSeat", ConnOBJ,
             ADODB.CursorTypeEnum.adOpenDynamic,
             ADODB.LockTypeEnum.adLockOptimistic)
        RecSetBusi.Open("Select * FROM BusiSeat", ConnOBJ,
             ADODB.CursorTypeEnum.adOpenDynamic,
             ADODB.LockTypeEnum.adLockOptimistic)
        Call EcoSeatForm()
        Call BusiSeatForm()
        Call GrabDataFromDB()
    End Sub
    Private Sub GrabDataFromDB()

        Do While Not RecSetEco.EOF
            IndexNumOfEcoAry = RecSetEco.Fields("IndexNumOfEcoAry").Value
            Dim FirstName As String = RecSetEco.Fields("FirstName").Value
            Dim LastName As String = RecSetEco.Fields("LastName").Value
            EcoSeat(IndexNumOfEcoAry).Text = FirstName & " " & LastName
            EcoSeat(IndexNumOfEcoAry).ScrollBars = False
            EcoSeat(IndexNumOfEcoAry).ZoomFactor = 0.8
            EcoSeat(IndexNumOfEcoAry).BackColor = Color.Red
            RecSetEco.MoveNext()
        Loop

        Do While Not RecSetBusi.EOF
            IndexNumOfBusiAry = RecSetBusi.Fields("IndexNumOfBusiAry").Value
            Dim FirstName As String = RecSetBusi.Fields("FirstName").Value
            Dim LastName As String = RecSetBusi.Fields("LastName").Value
            BusiSeat(IndexNumOfBusiAry).Text = FirstName & " " & LastName
            BusiSeat(IndexNumOfBusiAry).ScrollBars = False
            BusiSeat(IndexNumOfBusiAry).ZoomFactor = 0.8
            BusiSeat(IndexNumOfBusiAry).BackColor = Color.Red
            RecSetBusi.MoveNext()
        Loop
    End Sub
    Private Sub EcoSeatForm()
        Dim EX As Integer = 0
        Dim EY As Integer = 0
        For Index As Integer = 0 To 99
            EcoSeat(Index) = New RichTextBox
            EcoSeat(Index).Size = New Size(30, 30)
            EcoSeat(Index).BackColor = Color.White
            EcoSeat(Index).Visible = True
            EcoSeat(Index).Text = "E"
            EcoSeat(Index).Name = "EcoSeat" & Index.ToString
            Dim PxOfEcoSeat As Integer = (EX + 1) * 35
            Dim PyOfEcoSeat As Integer = (EY + 1) * 35
            PyOfEcoSeat = (EY + 1) * 35
            EcoSeat(Index).Location = New Point(PxOfEcoSeat, PyOfEcoSeat)
            EX += 1
            If EX > 9 Then
                EX = 0
                EY += 1
            End If
            Eco.Controls.Add(EcoSeat(Index))
            AddHandler EcoSeat(Index).Click, AddressOf EcoSeat_Click

            Dim EcoSeatRowSign(9), EcoSeatColSign(9) As RichTextBox
            If Index < 10 Then
                EcoSeatColSign(Index) = New RichTextBox
                EcoSeatColSign(Index).Size = New Size(30, 30)
                EcoSeatColSign(Index).BackColor = Color.AntiqueWhite
                EcoSeatColSign(Index).Location = New Point((Index + 1) * 35, 1)
                Dim Letter As String = "ABCDEFGHIJ"
                Dim ColLetter = Letter.Replace(Letter, Letter.Substring(Index, 1))
                EcoSeatColSign(Index).Text = ColLetter
                EcoSeatColSign(Index).Visible = True
                Eco.Controls.Add(EcoSeatColSign(Index))
                EcoSeatRowSign(Index) = New RichTextBox
                EcoSeatRowSign(Index).Size = New Size(30, 30)
                EcoSeatRowSign(Index).BackColor = Color.AntiqueWhite
                EcoSeatRowSign(Index).Location = New Point(1, (Index + 1) * 35)
                EcoSeatRowSign(Index).Visible = True
                EcoSeatRowSign(Index).Text = Str(Index + 1)
                Eco.Controls.Add(EcoSeatRowSign(Index))
            End If
        Next
    End Sub
    Private Sub BusiSeatForm()
        Dim BX As Integer = 0
        Dim BY As Integer = 0
        For Index As Integer = 0 To 19
            BusiSeat(Index) = New RichTextBox
            BusiSeat(Index).Size = New Size(30, 30)
            BusiSeat(Index).BackColor = Color.White
            BusiSeat(Index).Visible = True
            BusiSeat(Index).Text = "E"
            BusiSeat(Index).Name = "BusiSeat" & Index.ToString
            Dim PxOfBusiSeat As Integer = 235 + 35 * BX
            Dim PyOfBusiSeat As Integer = 135 + 35 * BY
            BusiSeat(Index).Location = New Point(PxOfBusiSeat, PyOfBusiSeat)
            Busi.Controls.Add(BusiSeat(Index))
            AddHandler BusiSeat(Index).Click, AddressOf BusiSeat_Click
            BX += 1
            If BX > 4 Then
                BX = 0
                BY += 1
            End If

            Dim BusiSeatRowSign(3), BusiSeatColSign(4) As RichTextBox
            If Index < 5 Then
                BusiSeatColSign(Index) = New RichTextBox
                BusiSeatColSign(Index).Size = New Size(30, 30)
                BusiSeatColSign(Index).BackColor = Color.AntiqueWhite
                BusiSeatColSign(Index).Location = New Point(235 + 35 * Index, 95)
                Dim Letter As String = "ABCDE"
                Dim ColLetter = Letter.Replace(Letter, Letter.Substring(Index, 1))
                BusiSeatColSign(Index).Text = ColLetter
                BusiSeatColSign(Index).Visible = True
                Busi.Controls.Add(BusiSeatColSign(Index))
            End If
            If Index < 4 Then
                BusiSeatRowSign(Index) = New RichTextBox
                BusiSeatRowSign(Index).Size = New Size(30, 30)
                BusiSeatRowSign(Index).BackColor = Color.AntiqueWhite
                BusiSeatRowSign(Index).Location = New Point(200, 135 + Index * 35)
                BusiSeatRowSign(Index).Visible = True
                BusiSeatRowSign(Index).Text = Str(Index + 1)
                Busi.Controls.Add(BusiSeatRowSign(Index))
            End If
        Next
    End Sub
    Private Sub SaveInEcoDB()

        RecSetEco.AddNew()
        RecSetEco.Fields("IndexNumOfEcoAry").Value = IndexNumOfEcoAry
        RecSetEco.Fields("FirstName").Value = TextBox1.Text
        RecSetEco.Fields("LastName").Value = TextBox2.Text
        EcoSeatRowNum = Math.Truncate(IndexNumOfEcoAry / 10) + 1
        EcoSeatColNum = IndexNumOfEcoAry Mod 10 + 1
        Dim EcoSeatColCha As String
        Dim Letter As String = "ABCDEFGHIJ"
        EcoSeatColCha = Letter.Replace(Letter, Letter.Substring(EcoSeatColNum - 1, 1))
        RecSetEco.Fields("EcoSeatNum").Value = EcoSeatRowNum & EcoSeatColCha
        RecSetEco.Update()
        RecSetEco.MoveNext()
    End Sub
    Private Sub SaveInBusiDB()
        RecSetBusi.AddNew()
        RecSetBusi.Fields("IndexNumOfBusiAry").Value = IndexNumOfBusiAry
        RecSetBusi.Fields("FirstName").Value = TextBox8.Text
        RecSetBusi.Fields("LastName").Value = TextBox7.Text
        BusiSeatRowNum = Math.Truncate(IndexNumOfBusiAry / 5) + 1
        BusiSeatColNum = IndexNumOfBusiAry Mod 5 + 1
        Dim BusiSeatColCha As String
        Dim Letter As String = "ABCDE"
        BusiSeatColCha = Letter.Replace(Letter, Letter.Substring(BusiSeatColNum - 1, 1))
        RecSetBusi.Fields("BusiSeatNum").Value = BusiSeatRowNum & BusiSeatColCha
        RecSetBusi.Update()
        RecSetBusi.MoveNext()
    End Sub

    Private Sub EcoSeat_Click(sender As Object, e As System.EventArgs)
        Dim Clicked As String = "EcoClicked"
        EcoSeatClicked = True
        IndexNumOfEcoAry = Array.IndexOf(EcoSeat, sender)
        If RecSetEco.BOF = False Then
            RecSetEco.MoveFirst()
        End If
        Dim Criteria As String = "IndexNumOfEcoAry = " & Str(IndexNumOfEcoAry)
        RecSetEco.Find(Criteria)
        If RecSetEco.EOF = True Then
            Call SaveInEcoDB()
            Call PopuForm(Clicked)
        Else
            RecSetEco.Delete()
            RecSetEco.Update()
            Call FormClear(Clicked)
        End If
    End Sub
    Private Sub PopuForm(ClickOrInput)
        Dim SeatSelected As RichTextBox
        If ClickOrInput = "EcoClicked" Or ClickOrInput = "EcoBooked" Then
            SeatSelected = EcoSeat(IndexNumOfEcoAry)
            SeatSelected.Text = TextBox1.Text & " " & TextBox2.Text
        ElseIf ClickOrInput = "BusiClicked" Or ClickOrInput = "BusiBooked" Then
            SeatSelected = BusiSeat(IndexNumOfBusiAry)
            SeatSelected.Text = TextBox8.Text & " " & TextBox7.Text
        Else
            Exit Sub
        End If
        SeatSelected.ScrollBars = False
        SeatSelected.ZoomFactor = 0.8
        SeatSelected.BackColor = Color.Red
    End Sub
    Private Sub FormClear(ClickOrInput)
        Dim SeatSelected As RichTextBox
        If ClickOrInput = "EcoClicked" Or ClickOrInput = "EcoCancel" Then
            SeatSelected = EcoSeat(IndexNumOfEcoAry)
        ElseIf ClickOrInput = "BusiClicked" Or ClickOrInput = "BusiCancel" Then
            SeatSelected = BusiSeat(IndexNumOfBusiAry)
        Else
            Exit Sub
        End If
        SeatSelected.Text = "E"
        SeatSelected.ScrollBars = False
        SeatSelected.ZoomFactor = 1
        SeatSelected.BackColor = Color.White
    End Sub
    Private Sub BusiSeat_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Clicked As String = "BusiClicked"
        BusiSeatClicked = True
        IndexNumOfBusiAry = Array.IndexOf(BusiSeat, sender)
        If RecSetBusi.BOF = False Then
            RecSetBusi.MoveFirst()
        End If
        Dim Criteria As String = "IndexNumOfBusiAry = " & Str(IndexNumOfBusiAry)
        RecSetBusi.Find(Criteria)
        If RecSetBusi.EOF Then
            Call SaveInBusiDB()
            Call PopuForm(Clicked)
        Else
            RecSetBusi.Delete()
            RecSetBusi.Update()
            Call FormClear(Clicked)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        EcoRowColInput = True
        Dim LettList As New List(Of Char)({"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"})
        For Index As Integer = 0 To 9
            If LettList(Index) = TextBox4.Text Or LCase(LettList(Index)) = TextBox4.Text Then
                EcoSeatColNum = Index + 1
                Exit For
            Else
                If Index = 9 Then
                    MsgBox("wrong input")
                    Exit Sub
                End If
            End If
        Next
        Success = Int32.TryParse(TextBox3.Text, EcoSeatRowNum) And
                  EcoSeatRowNum > 0 AndAlso EcoSeatColNum > 0 AndAlso
                  EcoSeatRowNum < 11 AndAlso EcoSeatColNum < 11
        If Success Then
            IndexNumOfEcoAry = (EcoSeatRowNum - 1) * 10 + EcoSeatColNum - 1
        Else
            MsgBox("sorry you enter wrong row or column number, please re-enter")
            Exit Sub
        End If
        If EcoSeat(IndexNumOfEcoAry).Text = "E" Then
            Call SaveInEcoDB()
            Dim Booked As String = "EcoBooked"
            Call PopuForm(Booked)
        Else
            MsgBox("sorry, the seat has been taken, please select other seats")
            Exit Sub
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Cancel As String = "EcoCancel"
        If EcoSeatClicked = True Then
            Call FormClear(Cancel)
            Call CancelEcoDB()

        End If
        Dim LettList As New List(Of Char)({"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"})
        For Index As Integer = 0 To 9
            If LettList(Index) = TextBox4.Text Or LCase(LettList(Index)) = TextBox4.Text Then
                EcoSeatColNum = Index + 1
                Exit For
            Else
                If Index = 9 Then
                    MsgBox("wrong input")
                    Exit Sub
                End If
            End If
        Next
        Success = Int32.TryParse(TextBox3.Text, EcoSeatRowNum) And
                  EcoSeatRowNum > 0 AndAlso EcoSeatColNum > 0 AndAlso
                  EcoSeatRowNum < 11 AndAlso EcoSeatColNum < 11
        If Success Then
            IndexNumOfEcoAry = (EcoSeatRowNum - 1) * 10 + EcoSeatColNum - 1
        Else
            Exit Sub
        End If

        Call FormClear(Cancel)
        Call CancelEcoDB()

    End Sub
    Private Sub CancelEcoDB()
        If RecSetEco.BOF = False Then
            RecSetEco.MoveFirst()
        End If
        Dim Criteria2 As String = "IndexNumOfEcoAry = " & Str(IndexNumOfEcoAry)
        RecSetEco.Find(Criteria2)
        If Not RecSetEco.EOF Then
            RecSetEco.Delete()
            RecSetEco.Update()
        Else
            MsgBox("wrong input")
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        BusiRowColInput = True
        Dim LettList As New List(Of Char)({"A", "B", "C", "D", "E"})
        For Index As Integer = 0 To 4
            If LettList(Index) = TextBox5.Text Or LCase(LettList(Index)) = TextBox5.Text Then
                BusiSeatColNum = Index + 1
                Exit For
            Else
                If Index = 4 Then
                    MsgBox("wrong input")
                    Exit Sub
                End If
            End If
        Next
        Success = Int32.TryParse(TextBox6.Text, BusiSeatRowNum) And
                  BusiSeatRowNum > 0 AndAlso BusiSeatColNum > 0 AndAlso
                  BusiSeatRowNum < 5 AndAlso BusiSeatColNum < 6
        If Success Then
            IndexNumOfBusiAry = (BusiSeatRowNum - 1) * 5 + BusiSeatColNum - 1
        Else
            MsgBox("sorry you enter wrong row or column number, please re-enter")
            Exit Sub
        End If
        If BusiSeat(IndexNumOfBusiAry).Text = "E" Then
            Call SaveInBusiDB()
            Dim Booked As String = "BusiBooked"
            Call PopuForm(Booked)
        Else
            MsgBox("sorry, the seat has been taken, please select other seats")
            Exit Sub
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Cancel As String = "BusiCancel"
        If BusiSeatClicked = True Then
            Call FormClear(Cancel)
            Call CancelBusiDB()
        End If
        Dim LettList As New List(Of Char)({"A", "B", "C", "D", "E"})
        For Index As Integer = 0 To 4
            If LettList(Index) = TextBox5.Text Or LCase(LettList(Index)) = TextBox5.Text Then
                BusiSeatColNum = Index + 1
                Exit For
            Else
                If Index = 4 Then
                    MsgBox("wrong input")
                    Exit Sub
                End If
            End If
        Next
        Success = Int32.TryParse(TextBox6.Text, BusiSeatRowNum) And
                  BusiSeatRowNum > 0 AndAlso BusiSeatColNum > 0 AndAlso
                  BusiSeatRowNum < 5 AndAlso BusiSeatColNum < 6
        If Success Then
            IndexNumOfBusiAry = (BusiSeatRowNum - 1) * 5 + BusiSeatColNum - 1
        Else
            Exit Sub
        End If
        Call FormClear(Cancel)
        Call CancelBusiDB()
    End Sub
    Private Sub CancelBusiDB()
        If RecSetBusi.BOF = False Then
            RecSetBusi.MoveFirst()
        End If
        Dim Criteria As String = "IndexNumOfBusiAry = " & Str(IndexNumOfBusiAry)
        RecSetBusi.Find(Criteria)
        If Not RecSetBusi.EOF Then
            RecSetBusi.Delete()
            RecSetBusi.Update()
        Else
            MsgBox("wrong input")
        End If
    End Sub
End Class