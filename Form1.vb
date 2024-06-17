Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Numerics

Public Class Form1

    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Cashflow_Main\adminpaneldb.accdb;Persist Security Info=False;"
    Dim connection As New OleDbConnection(connectionString)

    Dim SD_Cost As Decimal = 0
    Dim SD_Cashflow As Decimal = 0
    Dim SD_Downpayment As Decimal = 0
    Dim BD_Cost As Decimal = 0
    Dim BD_Cashflow As Decimal = 0
    Dim BD_Downpayment As Decimal = 0
    Dim Chance_Val As Decimal = 0
    Dim Penalty_Val As Decimal = 0
    Dim Ploanint As Decimal = 12
    Dim Bloanint As Decimal = 10
    Dim Proploan As Decimal = 6.5

    Dim Current_Player As String = ""
    Dim PC As Integer = 0

    Dim PayDayCount As Integer = 0

    Public Function Insert_Asset(Name As String, Amount As Decimal, Player As String) As Integer
        Try
            Dim insertQuery As String = "INSERT INTO Assets (AssetName,Amount,Player) VALUES (@Value1, @Value2, @Value3);"
            Dim command As New OleDbCommand(insertQuery, connection)

            command.Parameters.AddWithValue("@Value1", Name)
            command.Parameters.AddWithValue("@Value2", Amount)
            command.Parameters.AddWithValue("@Value3", Player)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Insert_Liabilities(Name As String, Amount As Decimal, Player As String) As Integer
        Try
            Dim insertQuery As String = "INSERT INTO Liabilities (LiabilitiesName,Amount,Player) VALUES (@Value1, @Value2, @Value3);"
            Dim command As New OleDbCommand(insertQuery, connection)

            command.Parameters.AddWithValue("@Value1", Name)
            command.Parameters.AddWithValue("@Value2", Amount)
            command.Parameters.AddWithValue("@Value3", Player)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Update_Liabilities(Playerd As String) As Integer
        Try
            Dim insertQuery As String = "DELETE FROM Liabilities WHERE Player = @Value1;"
            Dim command As New OleDbCommand(insertQuery, connection)

            command.Parameters.AddWithValue("@Value1", Playerd)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)
        End Try

        Return 0
    End Function

    Public Function Insert_Expenses(Name As String, Amount As Decimal, Player As String) As Integer
        Try
            Dim insertQuery As String = "INSERT INTO Expenses (ExpensesName,Amount,Player) VALUES (@Value1, @Value2, @Value3);"
            Dim command As New OleDbCommand(insertQuery, connection)

            command.Parameters.AddWithValue("@Value1", Name)
            command.Parameters.AddWithValue("@Value2", Amount)
            command.Parameters.AddWithValue("@Value3", Player)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Insert_Income(Name As String, Amount As Decimal, Player As String) As Integer
        Try
            Dim insertQuery As String = "INSERT INTO Income (IncomeName,Amount,Player) VALUES (@Value1, @Value2, @Value3);"
            Dim command As New OleDbCommand(insertQuery, connection)

            command.Parameters.AddWithValue("@Value1", Name)
            command.Parameters.AddWithValue("@Value2", Amount)
            command.Parameters.AddWithValue("@Value3", Player)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Check_Current_Balance(Player As String) As Decimal
        Dim Cash_Available As Decimal = 0

        Try
            Dim selectquery As String = "SELECT cash FROM [Player Details] WHERE [PlayerName] = @Value1"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Player)
            Dim reader As OleDbDataReader = command.ExecuteReader()


            While reader.Read()
                Cash_Available = reader(0)
            End While
            reader.Close()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return Cash_Available
    End Function

    Public Function Check_Loan_Balance(Liability_Type As String, Player As String) As Decimal
        Dim Loan_Available As Decimal = 0

        Try
            Dim selectquery As String = "SELECT Amount FROM Liabilities WHERE Player = @Value1 AND LiabilitiesName = @Value2;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Player)
            command.Parameters.AddWithValue("@Value2", Liability_Type)
            Dim reader As OleDbDataReader = command.ExecuteReader()


            While reader.Read()
                Loan_Available = reader(0)
            End While
            reader.Close()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return Loan_Available
    End Function


    Public Function Update_Cash(Amount As Decimal, Player As String) As Integer
        Try
            Dim selectquery As String = "UPDATE [Player Details] SET cash = @Value1 WHERE [PlayerName] = @Value2"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Amount)
            command.Parameters.AddWithValue("@Value2", Player)

            command.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function


    Public Function Update_Buy(Cost As Decimal, Cashflow As Decimal, Player As String) As Integer
        Try
            Dim Asset_Available = 0
            Dim Income_Available = 0
            Dim Payday_Available = 0
            Dim PassInc_Available = 0
            Dim Cash_Available = 0

            Dim selectquery As String = "SELECT Assets, Income, Payday, [Passive Income], cash FROM [Player Details] WHERE [PlayerName] = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Player)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                Asset_Available = reader(0)
                Income_Available = reader(1)
                Payday_Available = reader(2)
                PassInc_Available = reader(3)
                Cash_Available = reader(4)
            End While

            reader.Close()

            Dim updatequery As String = "UPDATE [Player Details] SET Assets = @Value1, Income = @Value2,Payday = @Value3,  [Passive Income] = @Value4, cash=@Value5 WHERE [PlayerName] = @Value6"
            Dim command2 As New OleDbCommand(updatequery, connection)

            command2.Parameters.AddWithValue("@Value1", Asset_Available + Cost)
            command2.Parameters.AddWithValue("@Value2", Income_Available + Cashflow)
            command2.Parameters.AddWithValue("@Value3", Payday_Available + Cashflow)
            command2.Parameters.AddWithValue("@Value4", PassInc_Available + Cashflow)
            command2.Parameters.AddWithValue("@Value5", Cash_Available - Cost)
            command2.Parameters.AddWithValue("@Value6", Player)

            command2.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function
    Public Function Update_Borrow(Amount As Decimal, Roi As Decimal, Player As String) As Integer
        Try
            Dim Liability_Available = 0
            Dim Expense_Available = 0
            Dim Payday_Available = 0
            Dim PassInc_Available = 0
            Dim Cash_Available = 0
            Dim selectquery As String = "SELECT Liabilities, Expenses, Payday, [Passive Income], cash FROM [Player Details] WHERE [PlayerName] = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Player)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                Liability_Available = reader(0)
                Expense_Available = reader(1)
                Payday_Available = reader(2)
                PassInc_Available = reader(3)
                Cash_Available = reader(4)
            End While
            reader.Close()

            Dim updatequery As String = "UPDATE [Player Details] SET Liabilities = @Value1, Expenses = @Value2,Payday = @Value3,cash=@Value5 WHERE [PlayerName] = @Value6;"
            Dim command2 As New OleDbCommand(updatequery, connection)

            command2.Parameters.AddWithValue("@Value1", Liability_Available + Amount)
            command2.Parameters.AddWithValue("@Value2", Expense_Available + ((Roi * Amount) / 100))
            command2.Parameters.AddWithValue("@Value3", Payday_Available - ((Roi * Amount) / 100))
            command2.Parameters.AddWithValue("@Value5", Cash_Available + Amount)
            command2.Parameters.AddWithValue("@Value6", Player)

            command2.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.stocks' table. You can move, or remove it, as needed.
        Me.StocksTableAdapter.Fill(Me.AdminpaneldbDataSet.stocks)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Penalty' table. You can move, or remove it, as needed.
        Me.PenaltyTableAdapter.Fill(Me.AdminpaneldbDataSet.Penalty)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Chance' table. You can move, or remove it, as needed.
        Me.ChanceTableAdapter.Fill(Me.AdminpaneldbDataSet.Chance)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.BigDeals' table. You can move, or remove it, as needed.
        Me.BigDealsTableAdapter.Fill(Me.AdminpaneldbDataSet.BigDeals)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.SmallDeals' table. You can move, or remove it, as needed.
        Me.SmallDealsTableAdapter.Fill(Me.AdminpaneldbDataSet.SmallDeals)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.stockdetails' table. You can move, or remove it, as needed.
        Me.StockdetailsTableAdapter.Fill(Me.AdminpaneldbDataSet.stockdetails)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Expenses' table. You can move, or remove it, as needed.
        Me.ExpensesTableAdapter.Fill(Me.AdminpaneldbDataSet.Expenses)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Income' table. You can move, or remove it, as needed.
        Me.IncomeTableAdapter.Fill(Me.AdminpaneldbDataSet.Income)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Liabilities' table. You can move, or remove it, as needed.
        Me.LiabilitiesTableAdapter.Fill(Me.AdminpaneldbDataSet.Liabilities)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Assets' table. You can move, or remove it, as needed.
        Me.AssetsTableAdapter.Fill(Me.AdminpaneldbDataSet.Assets)
        'TODO: This line of code loads data into the 'AdminpaneldbDataSet.Player_Details' table. You can move, or remove it, as needed.
        Me.Player_DetailsTableAdapter.Fill(Me.AdminpaneldbDataSet.Player_Details)

        Try
            connection.Open()
        Catch ex As Exception
            MsgBox("Error in Connection: " + ex.Message)

        End Try
        Panel1.Enabled = False

        Panel3.Visible = True



    End Sub
    Public Function Set_SD_panel() As Integer
        Try
            Dim selectquery As String = "SELECT * FROM SmallDeals WHERE Name_SD = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", ComboBox1.Text)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                SD_Cost = reader(1)
                SD_Downpayment = reader(2)
                SD_Cashflow = reader(3)
            End While

            reader.Close()


            Label16.Text = SD_Cost.ToString
            Label17.Text = SD_Cashflow.ToString
            Label18.Text = SD_Downpayment.ToString


        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Set_BD_panel() As Integer
        Try
            Dim selectquery As String = "SELECT * FROM BigDeals WHERE Name_BD = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", ComboBox2.Text)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                BD_Cost = reader(1)
                BD_Downpayment = reader(2)
                BD_Cashflow = reader(3)
            End While

            reader.Close()

            Label19.Text = BD_Downpayment.ToString
            Label20.Text = BD_Cashflow.ToString
            Label21.Text = BD_Cost.ToString

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Set_Chance_panel() As Integer
        Try
            Dim selectquery As String = "SELECT * FROM Chance WHERE Name_C = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", ComboBox3.Text)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                Chance_Val = reader(1)
            End While

            reader.Close()

            Label29.Text = Chance_Val.ToString

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function

    Public Function Set_Penalty_panel() As Integer
        Try
            Dim selectquery As String = "SELECT * FROM Penalty WHERE Name_P = @Value1;"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", ComboBox4.Text)
            Dim reader As OleDbDataReader = command.ExecuteReader()

            While reader.Read()
                Penalty_Val = reader(1)
            End While
            reader.Close()

            Label27.Text = Penalty_Val.ToString

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try

        Return 0
    End Function
    Public Function Active_Panel() As Integer
        SDPanel.Visible = False
        BDPanel.Visible = False
        ChancePanel.Visible = False
        PenaltyPanel.Visible = False
        BorrowPanel.Visible = False
        RepayPanel.Visible = False
        TradePanel.Visible = False
        SD_Cost = 0
        SD_Cashflow = 0
        SD_Downpayment = 0
        BD_Cost = 0
        BD_Cashflow = 0
        BD_Downpayment = 0
        Chance_Val = 0
        Penalty_Val = 0

        Try
            Dim DataTable1 = New DataTable()
            Dim adapter1 = New OleDbDataAdapter()

            Dim DataTable2 = New DataTable()
            Dim adapter2 = New OleDbDataAdapter()
            Dim DataTable3 = New DataTable()
            Dim adapter3 = New OleDbDataAdapter()
            Dim DataTable4 = New DataTable()
            Dim adapter4 = New OleDbDataAdapter()
            Dim DataTable5 = New DataTable()
            Dim adapter5 = New OleDbDataAdapter()
            Dim DataTable6 = New DataTable()
            Dim adapter6 = New OleDbDataAdapter()

            DataGridView2.DataSource = DataTable1
            DataGridView3.DataSource = DataTable2
            DataGridView4.DataSource = DataTable3
            DataGridView5.DataSource = DataTable4
            DataGridView6.DataSource = DataTable5
            DataGridView1.DataSource = DataTable6


            Dim Query1 = "Select AssetName,Amount from Assets where Player = @Player1"
            Dim Query2 = "Select LiabilitiesName,Amount from Liabilities where Player = @Player1"
            Dim Query3 = "Select IncomeName,Amount from Income where Player = @Player1"
            Dim Query4 = "Select ExpensesName,Amount from Expenses where Player = @Player1"
            Dim Query5 = "Select StockName,Quantity_In_hand,Stock_Value from stockdetails where Player = @Player1"
            Dim Query6 = "Select * FROM [Player Details] WHERE [PlayerName] = @Play1 OR [PlayerName] = @Play2 OR [PlayerName] = @Play3"

            adapter1.SelectCommand = New OleDbCommand(Query1, connection)
            adapter1.SelectCommand.Parameters.AddWithValue("@Player1", Current_Player)
            adapter2.SelectCommand = New OleDbCommand(Query2, connection)
            adapter2.SelectCommand.Parameters.AddWithValue("@Player1", Current_Player)
            adapter3.SelectCommand = New OleDbCommand(Query3, connection)
            adapter3.SelectCommand.Parameters.AddWithValue("@Player1", Current_Player)
            adapter4.SelectCommand = New OleDbCommand(Query4, connection)
            adapter4.SelectCommand.Parameters.AddWithValue("@Player1", Current_Player)
            adapter5.SelectCommand = New OleDbCommand(Query5, connection)
            adapter5.SelectCommand.Parameters.AddWithValue("@Player1", Current_Player)
            adapter6.SelectCommand = New OleDbCommand(Query6, connection)
            adapter6.SelectCommand.Parameters.AddWithValue("@Play1", BTT1.Text)
            adapter6.SelectCommand.Parameters.AddWithValue("@Play2", BTT2.Text)
            adapter6.SelectCommand.Parameters.AddWithValue("@Play3", BTT3.Text)

            DataTable1.Clear()
            DataTable2.Clear()
            DataTable3.Clear()
            DataTable4.Clear()
            DataTable5.Clear()
            DataTable6.Clear()



            adapter1.Fill(DataTable1)
            adapter2.Fill(DataTable2)
            adapter3.Fill(DataTable3)
            adapter4.Fill(DataTable4)
            adapter5.Fill(DataTable5)
            adapter6.Fill(DataTable6)
        Catch ex As Exception
            MsgBox("Transaction Error :" + ex.Message)

        End Try

        Return 0
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Active_Panel()
        SDPanel.Visible = True
        ComboBox1.DataSource = SmallDealsBindingSource
        ComboBox1.DisplayMember = "Name_SD"
        Set_SD_panel()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Active_Panel()
        BDPanel.Visible = True
        ComboBox2.DataSource = BigDealsBindingSource
        ComboBox2.DisplayMember = "Name_BD"
        Set_BD_panel()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Active_Panel()
        PenaltyPanel.Visible = True
        ComboBox4.DataSource = PenaltyBindingSource
        ComboBox4.DisplayMember = "Name_P"
        Set_Penalty_panel()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Active_Panel()
        ChancePanel.Visible = True
        ComboBox3.DataSource = ChanceBindingSource
        ComboBox3.DisplayMember = "Name_C"
        Set_Chance_panel()
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        Set_SD_panel()
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Set_BD_panel()
    End Sub
    Private Sub ComboBox3_TextChanged(sender As Object, e As EventArgs) Handles ComboBox3.TextChanged
        Set_Chance_panel()
    End Sub
    Private Sub ComboBox4_TextChanged(sender As Object, e As EventArgs) Handles ComboBox4.TextChanged
        Set_Penalty_panel()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim Cash_Avaialble As Decimal = Check_Current_Balance(Current_Player)
        If (Cash_Avaialble < SD_Downpayment) Then
            MsgBox("Sorry Not Enough Cash")
        Else
            If (Cash_Avaialble < SD_Cost) Then
                Dim Borrow_Amount = SD_Cost - Cash_Avaialble
                Dim New_Amount = Check_Loan_Balance("Loan @10", Current_Player) + Borrow_Amount
                Dim Amount1 = Check_Loan_Balance("Loan @12", Current_Player)
                Dim Amount2 = Check_Loan_Balance("Loan @6.5", Current_Player)
                Update_Liabilities(Current_Player)
                Insert_Liabilities("Loan @10", New_Amount, Current_Player)
                Insert_Liabilities("Loan @12", Amount1, Current_Player)
                Insert_Liabilities("Loan @6.5", Amount2, Current_Player)
                Insert_Expenses("Loan 10%", (Bloanint * Borrow_Amount) / 100, Current_Player)
                Update_Borrow(Borrow_Amount, Bloanint, Current_Player)
            End If
            Insert_Asset(ComboBox1.Text, SD_Cost, Current_Player)
            Insert_Income(ComboBox1.Text + " Income", SD_Cashflow, Current_Player)
            Update_Buy(SD_Cost, SD_Cashflow, Current_Player)
        End If
        Active_Panel()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim Cash_Avaialble As Decimal = Check_Current_Balance(Current_Player)
        If (Cash_Avaialble < BD_Downpayment) Then
            MsgBox("Sorry Not Enough Cash")
        Else
            If (Cash_Avaialble < BD_Cost) Then
                Dim Borrow_Amount = BD_Cost - Cash_Avaialble
                Dim New_Amount = Check_Loan_Balance("Loan @6.5", Current_Player) + Borrow_Amount
                Dim Amount1 = Check_Loan_Balance("Loan @12", Current_Player)
                Dim Amount2 = Check_Loan_Balance("Loan @10", Current_Player)
                Update_Liabilities(Current_Player)
                Insert_Liabilities("Loan @6.5", New_Amount, Current_Player)
                Insert_Liabilities("Loan @12", Amount1, Current_Player)
                Insert_Liabilities("Loan @10", Amount2, Current_Player)
                Insert_Expenses("Loan 6.5%", (Proploan * Borrow_Amount) / 100, Current_Player)
                Update_Borrow(Borrow_Amount, Proploan, Current_Player)
            End If
            Insert_Asset(ComboBox2.Text, BD_Cost, Current_Player)
            Insert_Income(ComboBox2.Text + " Income", BD_Cashflow, Current_Player)
            Update_Buy(BD_Cost, BD_Cashflow, Current_Player)
        End If
        Active_Panel()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim Cash_Avaialble As Decimal = Check_Current_Balance(Current_Player)

        If (Cash_Avaialble < Penalty_Val) Then
            Dim Borrow_Amount = Penalty_Val - Cash_Avaialble
            Dim New_Amount = Check_Loan_Balance("Loan @12", Current_Player) + Borrow_Amount
            Dim Amount1 = Check_Loan_Balance("Loan @10", Current_Player)
            Dim Amount2 = Check_Loan_Balance("Loan @6.5", Current_Player)
            Update_Liabilities(Current_Player)
            Insert_Liabilities("Loan @12", New_Amount, Current_Player)
            Insert_Liabilities("Loan @10", Amount1, Current_Player)
            Insert_Liabilities("Loan @6.5", Amount2, Current_Player)
            Insert_Expenses("Loan 12%", (Ploanint * Borrow_Amount) / 100, Current_Player)
            Update_Borrow(Borrow_Amount, Ploanint, Current_Player)
        End If
        Update_Cash(Check_Current_Balance(Current_Player) - Penalty_Val, Current_Player)
        Active_Panel()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim Cash_Avaialble As Decimal = Check_Current_Balance(Current_Player)
        Update_Cash(Cash_Avaialble + Chance_Val, Current_Player)
        Active_Panel()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If (PC = 1) Then
            BTT2.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT2.Text
            Current_Player_Label.Text = Current_Player
            PC = 2
        ElseIf (PC = 2) Then
            BTT3.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT2.BackColor = Color.WhiteSmoke
            Current_Player = BTT3.Text
            Current_Player_Label.Text = Current_Player
            PC = 3
        ElseIf (PC = 3) Then
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        Else
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        End If
        PayDayCount = 0
        Active_Panel()
    End Sub


    Public Function Insert_Into_Stock(StName As String, Player As String) As Integer
        Try
            Dim selectquery As String = "INSERT INTO `stockdetails` (`StockName`, `Quantity_In_Hand`, `Stock_Value`, `Player`) VALUES (@SName, @SQty,@SVal,@Player);"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@SName", StName)
            command.Parameters.AddWithValue("@SQty", 0)
            command.Parameters.AddWithValue("@SVal", 0.00)
            command.Parameters.AddWithValue("@Player", Player)

            command.ExecuteNonQuery()


        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)

        End Try
        Return 0
    End Function
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim Income_Start As Decimal = 0
        Dim Expense_Start As Decimal = 0
        Dim Cash_Start As Decimal = 0
        Dim PayDay_Start As Decimal = 0

        If (TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "") Then
            MsgBox("Empty Inputs Not Allowed")
        ElseIf (TextBox1.Text = TextBox2.Text Or TextBox2.Text = TextBox3.Text Or TextBox3.Text = TextBox1.Text) Then
            MsgBox("Enter Unique Names")
        Else
            Try

                Dim selectquery1 As String = "SELECT Val FROM Start_values WHERE Property =@Val"

                Dim command1 As New OleDbCommand(selectquery1, connection)
                command1.Parameters.AddWithValue("@Val", "Income")
                Dim reader As OleDbDataReader = command1.ExecuteReader()
                While reader.Read()
                    Income_Start = reader(0)
                End While
                reader.Close()

                Dim command2 As New OleDbCommand(selectquery1, connection)
                command2.Parameters.AddWithValue("@Val", "Expenses")
                Dim reader2 As OleDbDataReader = command2.ExecuteReader()
                While reader2.Read()
                    Expense_Start = reader2(0)
                End While
                reader2.Close()

                Dim command3 As New OleDbCommand(selectquery1, connection)
                command3.Parameters.AddWithValue("@Val", "Cash")
                Dim reader3 As OleDbDataReader = command3.ExecuteReader()
                While reader3.Read()
                    Cash_Start = reader3(0)
                End While
                reader3.Close()

                Dim command4 As New OleDbCommand(selectquery1, connection)
                command4.Parameters.AddWithValue("@Val", "Payday")
                Dim reader4 As OleDbDataReader = command4.ExecuteReader()
                While reader4.Read()
                    PayDay_Start = reader4(0)
                End While
                reader4.Close()


                Dim InsertIntoPlayerDetailsQuery = "INSERT INTO [Player Details] (PlayerName, Assets, Liabilities, Income, Expenses,Payday, [Passive Income],cash) VALUES (@Value1,0,0, @Value2, @Value3, @Value4,@Value5,@Value6);"

                Dim Command_insert_1 As New OleDbCommand(InsertIntoPlayerDetailsQuery, connection)
                Command_insert_1.Parameters.AddWithValue("@Value1", TextBox1.Text)
                Command_insert_1.Parameters.AddWithValue("@Value2", Income_Start)
                Command_insert_1.Parameters.AddWithValue("@Value3", Expense_Start)
                Command_insert_1.Parameters.AddWithValue("@Value4", PayDay_Start)
                Command_insert_1.Parameters.AddWithValue("@Value5", 0)
                Command_insert_1.Parameters.AddWithValue("@Value6", Cash_Start)
                Command_insert_1.ExecuteNonQuery()

                Dim Command_insert_2 As New OleDbCommand(InsertIntoPlayerDetailsQuery, connection)
                Command_insert_2.Parameters.AddWithValue("@Value1", TextBox2.Text)
                Command_insert_2.Parameters.AddWithValue("@Value2", Income_Start)
                Command_insert_2.Parameters.AddWithValue("@Value3", Expense_Start)
                Command_insert_2.Parameters.AddWithValue("@Value4", PayDay_Start)
                Command_insert_2.Parameters.AddWithValue("@Value5", 0)
                Command_insert_2.Parameters.AddWithValue("@Value6", Cash_Start)

                Command_insert_2.ExecuteNonQuery()

                Dim Command_insert_3 As New OleDbCommand(InsertIntoPlayerDetailsQuery, connection)
                Command_insert_3.Parameters.AddWithValue("@Value3", TextBox3.Text)
                Command_insert_3.Parameters.AddWithValue("@Value2", Income_Start)
                Command_insert_3.Parameters.AddWithValue("@Value3", Expense_Start)
                Command_insert_3.Parameters.AddWithValue("@Value4", PayDay_Start)
                Command_insert_3.Parameters.AddWithValue("@Value5", 0)
                Command_insert_3.Parameters.AddWithValue("@Value6", Cash_Start)

                Command_insert_3.ExecuteNonQuery()

                Dim QuerySEL = "SELECT * FROM stocks;"
                Dim Commadsel As New OleDbCommand(QuerySEL, connection)
                Dim readersel As OleDbDataReader = Commadsel.ExecuteReader()

                While readersel.Read()
                    Insert_Into_Stock(readersel(0), TextBox1.Text)
                    Insert_Into_Stock(readersel(0), TextBox2.Text)
                    Insert_Into_Stock(readersel(0), TextBox3.Text)
                End While

            Catch ex As Exception
                MessageBox.Show("Transaction Error: " & ex.Message)

            End Try

            Insert_Income("Monthly Salary", Income_Start, TextBox1.Text)
            Insert_Expenses("Monthly Expenses", Expense_Start, TextBox1.Text)
            Insert_Income("Monthly Salary", Income_Start, TextBox2.Text)
            Insert_Expenses("Monthly Expenses", Expense_Start, TextBox2.Text)
            Insert_Income("Monthly Salary", Income_Start, TextBox3.Text)
            Insert_Expenses("Monthly Expenses", Expense_Start, TextBox3.Text)

            Insert_Liabilities("Loan @6.5", 0, TextBox1.Text)
            Insert_Liabilities("Loan @10", 0, TextBox1.Text)
            Insert_Liabilities("Loan @12", 0, TextBox1.Text)
            Insert_Liabilities("Loan @6.5", 0, TextBox2.Text)
            Insert_Liabilities("Loan @10", 0, TextBox2.Text)
            Insert_Liabilities("Loan @12", 0, TextBox2.Text)
            Insert_Liabilities("Loan @6.5", 0, TextBox3.Text)
            Insert_Liabilities("Loan @10", 0, TextBox3.Text)
            Insert_Liabilities("Loan @12", 0, TextBox3.Text)


            Current_Player = TextBox1.Text

            BTT1.Text = TextBox1.Text
            BTT2.Text = TextBox2.Text
            BTT3.Text = TextBox3.Text
            Panel3.Visible = False
            Panel1.Enabled = True

            If (PC = 1) Then
                BTT2.BackColor = Color.Chartreuse
                BTT1.BackColor = Color.WhiteSmoke
                BTT3.BackColor = Color.WhiteSmoke
                Current_Player = BTT2.Text
                Current_Player_Label.Text = Current_Player
                PC = 2
            ElseIf (PC = 2) Then
                BTT3.BackColor = Color.Chartreuse
                BTT1.BackColor = Color.WhiteSmoke
                BTT2.BackColor = Color.WhiteSmoke
                Current_Player = BTT3.Text
                Current_Player_Label.Text = Current_Player
                PC = 3
            ElseIf (PC = 3) Then
                BTT1.BackColor = Color.Chartreuse
                BTT2.BackColor = Color.WhiteSmoke
                BTT3.BackColor = Color.WhiteSmoke
                Current_Player = BTT1.Text
                Current_Player_Label.Text = Current_Player
                PC = 1
            Else
                BTT1.BackColor = Color.Chartreuse
                BTT2.BackColor = Color.WhiteSmoke
                BTT3.BackColor = Color.WhiteSmoke
                Current_Player = BTT1.Text
                Current_Player_Label.Text = Current_Player
                PC = 1
            End If

            Active_Panel()


        End If


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If (PC = 1) Then
            BTT2.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT2.Text
            Current_Player_Label.Text = Current_Player
            PC = 2
        ElseIf (PC = 2) Then
            BTT3.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT2.BackColor = Color.WhiteSmoke
            Current_Player = BTT3.Text
            Current_Player_Label.Text = Current_Player
            PC = 3
        ElseIf (PC = 3) Then
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        Else
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        End If

        Active_Panel()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If (PC = 1) Then
            BTT2.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT2.Text
            Current_Player_Label.Text = Current_Player
            PC = 2
        ElseIf (PC = 2) Then
            BTT3.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT2.BackColor = Color.WhiteSmoke
            Current_Player = BTT3.Text
            Current_Player_Label.Text = Current_Player
            PC = 3
        ElseIf (PC = 3) Then
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        Else
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        End If

        Active_Panel()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Current_Player = TextBox1.Text

        BTT1.Text = TextBox1.Text
        BTT2.Text = TextBox2.Text
        BTT3.Text = TextBox3.Text
        Panel3.Visible = False
        Panel1.Enabled = True

        If (PC = 1) Then
            BTT2.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT2.Text
            Current_Player_Label.Text = Current_Player
            PC = 2
        ElseIf (PC = 2) Then
            BTT3.BackColor = Color.Chartreuse
            BTT1.BackColor = Color.WhiteSmoke
            BTT2.BackColor = Color.WhiteSmoke
            Current_Player = BTT3.Text
            Current_Player_Label.Text = Current_Player
            PC = 3
        ElseIf (PC = 3) Then
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        Else
            BTT1.BackColor = Color.Chartreuse
            BTT2.BackColor = Color.WhiteSmoke
            BTT3.BackColor = Color.WhiteSmoke
            Current_Player = BTT1.Text
            Current_Player_Label.Text = Current_Player
            PC = 1
        End If

        Active_Panel()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim Payday_Available As Decimal = 0
        If (PayDayCount = 0) Then
            Try
                Dim selectquery As String = "SELECT Payday FROM [Player Details] WHERE [PlayerName] = @Value1"
                Dim command As New OleDbCommand(selectquery, connection)

                command.Parameters.AddWithValue("@Value1", Current_Player)
                Dim reader As OleDbDataReader = command.ExecuteReader()


                While reader.Read()
                    Payday_Available = reader(0)
                End While
                reader.Close()

            Catch ex As Exception
                MessageBox.Show("Transaction Error: " & ex.Message)

            End Try

            Payday_Available = Payday_Available + Check_Current_Balance(Current_Player)
            If (Payday_Available < 0) Then
                MsgBox("Negative Cash Not Allowed - You Should Sell Your Stocks If available")
            Else
                Update_Cash(Payday_Available, Current_Player)
                MsgBox("Payday Sanctioned")
                PayDayCount = PayDayCount + 1
                Active_Panel()
            End If
        Else
            MsgBox("Single Payday Per Round")
        End If

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If (TextBox6.Text = "") Then
            MsgBox("Please Enter Some Amount")
        ElseIf (Decimal.Parse(TextBox6.Text) > 175000) Then
            MsgBox("Maximum Personal Loan Amount is 1,75,000")
        Else
            Try
                Dim New_Amount As Decimal = Check_Loan_Balance("Loan @12", Current_Player) + Decimal.Parse(TextBox6.Text)
                Dim Amount1 = Check_Loan_Balance("Loan @10", Current_Player)
                Dim Amount2 = Check_Loan_Balance("Loan @6.5", Current_Player)
                Update_Liabilities(Current_Player)
                Insert_Liabilities("Loan @12", New_Amount, Current_Player)
                Insert_Liabilities("Loan @10", Amount1, Current_Player)
                Insert_Liabilities("Loan @6.5", Amount2, Current_Player)
                Insert_Expenses("Loan 12%", (Ploanint * Decimal.Parse(TextBox6.Text)) / 100, Current_Player)
                Update_Borrow(Decimal.Parse(TextBox6.Text), Ploanint, Current_Player)
            Catch ex As Exception
                MsgBox("Enter Correct Amount")
            End Try
            Active_Panel()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Active_Panel()
        BorrowPanel.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Active_Panel()
        RepayPanel.Visible = True
    End Sub


    ''' Repayment
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click

        If (TextBox4.Text = "") Then
            MsgBox("Please Enter Some Amount")
        ElseIf (Decimal.Parse(TextBox4.Text) > Check_Current_Balance(Current_Player)) Then
            MsgBox("Insufficient Funds")
        Else
            Try
                Dim New_Amount As Decimal = Check_Loan_Balance("Loan @" + ComboBox6.Text, Current_Player) - Decimal.Parse(TextBox4.Text)

                If (ComboBox6.Text = "6.5") Then
                    Dim Amount1 = Check_Loan_Balance("Loan @12", Current_Player)
                    Dim Amount2 = Check_Loan_Balance("Loan @10", Current_Player)
                    Update_Liabilities(Current_Player)
                    Insert_Liabilities("Loan @6.5", New_Amount, Current_Player)
                    Insert_Liabilities("Loan @12", Amount1, Current_Player)
                    Insert_Liabilities("Loan @10", Amount2, Current_Player)
                ElseIf (ComboBox6.Text = "12") Then
                    Dim Amount1 = Check_Loan_Balance("Loan @6.5", Current_Player)
                    Dim Amount2 = Check_Loan_Balance("Loan @10", Current_Player)
                    Update_Liabilities(Current_Player)
                    Insert_Liabilities("Loan @12", New_Amount, Current_Player)
                    Insert_Liabilities("Loan @6.5", Amount1, Current_Player)
                    Insert_Liabilities("Loan @10", Amount2, Current_Player)
                ElseIf (ComboBox6.Text = "10") Then
                    Dim Amount1 = Check_Loan_Balance("Loan @6.5", Current_Player)
                    Dim Amount2 = Check_Loan_Balance("Loan @12", Current_Player)
                    Update_Liabilities(Current_Player)
                    Insert_Liabilities("Loan @10", New_Amount, Current_Player)
                    Insert_Liabilities("Loan @6.5", Amount1, Current_Player)
                    Insert_Liabilities("Loan @12", Amount2, Current_Player)
                End If
                Insert_Expenses("Loan Repayment", -(Decimal.Parse(ComboBox6.Text) * Decimal.Parse(TextBox4.Text)) / 100, Current_Player)
                Update_Borrow(-Decimal.Parse(TextBox4.Text), Decimal.Parse(ComboBox6.Text), Current_Player)
            Catch ex As Exception
                MsgBox("Enter Correct Amount")
            End Try
            Active_Panel()
        End If
    End Sub
    Public Function Val() As Integer
        Try
            Sto_Val.Text = (Decimal.Parse(TQty.Text) * Decimal.Parse(TPrice.Text)).ToString
        Catch ex As Exception
            Sto_Val.Text = "0.00"
        End Try
        Return 0
    End Function
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TPrice.TextChanged
        Val()
    End Sub

    Private Sub TQty_TextChanged(sender As Object, e As EventArgs) Handles TQty.TextChanged
        Val()
    End Sub

    Public Function Get_Stock_Qty(Player As String) As Integer
        Dim Stock_Qty As Integer

        Try
            Dim selectquery As String = "SELECT Quantity_In_Hand FROM stockdetails WHERE [Player] = @Value1 AND [StockName] = @Value2"
            Dim command As New OleDbCommand(selectquery, connection)

            command.Parameters.AddWithValue("@Value1", Player)
            command.Parameters.AddWithValue("@Value2", ComboBox5.Text)
            Dim reader As OleDbDataReader = command.ExecuteReader()


            While reader.Read()
                Stock_Qty = reader(0)
            End While
            reader.Close()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)
        End Try

        Return Stock_Qty
    End Function

    Public Function Update_Qty(NewQty As Integer, NewValue As Decimal, Player As String) As Integer
        Try

            Dim Query2 = "update stockdetails set Quantity_In_Hand = @QtyVal, Stock_Value = @StockVal where StockName = @SName AND Player = @Player;"

            Dim command2 As New OleDbCommand(Query2, connection)
            command2.Parameters.AddWithValue("@QtyVal", NewQty)
            command2.Parameters.AddWithValue("@StockVal", NewValue)
            command2.Parameters.AddWithValue("@SName", ComboBox5.Text)
            command2.Parameters.AddWithValue("@Player", Player)
            command2.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Transaction Error: " & ex.Message)
        Finally

        End Try
        Return 0
    End Function

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim StockValue As Decimal = Decimal.Parse(Sto_Val.Text)
        Dim Stock_Qty As Integer = Integer.Parse(TQty.Text)
        If (Check_Current_Balance(Current_Player) < StockValue) Then
            MsgBox("You Do Not Have enough Cash" & vbNewLine & "Please Borrow To Trade")
        Else
            Update_Cash(Check_Current_Balance(Current_Player) - StockValue, Current_Player)
            Update_Qty(Get_Stock_Qty(Current_Player) + Stock_Qty, (Get_Stock_Qty(Current_Player) + Stock_Qty) * Decimal.Parse(TPrice.Text), Current_Player)
            Active_Panel()
        End If
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim StockValue As Decimal = Decimal.Parse(Sto_Val.Text)
        Dim Stock_Qty As Integer = Integer.Parse(TQty.Text)
        If (Get_Stock_Qty(Current_Player) < Stock_Qty) Then
            MsgBox("Insufficient Qty of Stocks Available")
        Else
            Update_Cash(Check_Current_Balance(Current_Player) + StockValue, Current_Player)
            Update_Qty(Get_Stock_Qty(Current_Player) - Stock_Qty, (Get_Stock_Qty(Current_Player) - Stock_Qty) * Decimal.Parse(TPrice.Text), Current_Player)
            Active_Panel()
        End If
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Active_Panel()
        TradePanel.Visible = True
    End Sub

End Class
