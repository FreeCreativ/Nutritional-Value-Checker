Option Explicit On
Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class FrmHome
	Dim array_getdata(29) As Integer
	Dim displayTable As New DataTable
	Dim TotalTable As New DataTable
	Dim arrayTotal(0 To 28) As Object
	Private Sub FrmHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		'TODO: This line of code loads data into the 'NutValueCheckerDataSet1.Item' table. You can move, or remove it, as needed.
		Me.ItemTableAdapter.Fill(Me.NutValueCheckerDataSet1.Item)
		'TODO: This line of code loads data into the 'NutValueCheckerDataSet.FoodItems' table. You can move, or remove it, as needed.
		Me.FoodItemsTableAdapter.Fill(Me.NutValueCheckerDataSet.FoodItems)


	End Sub
	Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
		Dim query As String = "INSERT INTO [Item] ([Itemname], [EdibleConversionFactor], [Energy], [Water], [Protein], [Fat], [Carbohydrate], [Fibre], [Ash], [Ca], [Fe], [Mg], [P], [K], [Na], [Zn], [Cu], [Vit_A_Rae], [Retinol], [B_caroteneEquivalent], [Vit_D], [Vit_E], [Thiamin], [Riboflavin], [Niacin], [Vit_B6], [Folate], [Vit_B12], [Vit_C]) VALUES ('" & txtIname.Text & "', '" & txtEDF.Text & "', '" & txtEnergy.Text & " ',' " & txtWater.Text & "',' " & txtProtein.Text & "', '" & txtFat.Text & "', '" & txtCarb.Text & "', '" & txtFibre.Text & "','" & txtAsh.Text & "', '" & txtCa.Text & "', '" & txtFe.Text & "', '" & txtMg.Text & "', '" & txtP.Text & "', '" & txtk.Text & "', '" & txtNa.Text & "', '" & txtZn.Text & "',' " & txtCu.Text & "', '" & txtVitAR.Text & "',' " & txtRetinol.Text & "','" & txtCarotene.Text & "',' " & txtVitD.Text & "',' " & txtVitE.Text & "',' " & txtThiamin.Text & "',' " & txtRiboflavin.Text & "', '" & txtNiacin.Text & "', '" & txtVitB6.Text & "', '" & txtFo.Text & "', '" & txtVitB12.Text & "', '" & txtVitC.Text & "')"
		Dim conn As New OleDbConnection
		Dim comm As New OleDbCommand(query, conn)
		conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\NutValueChecker.mdb"

		conn.Open()
		comm.ExecuteNonQuery()
		conn.Close()
		txtAsh.Text = ""
		txtCa.Text = ""
		txtCarb.Text = ""
		txtCarotene.Text = ""
		txtCu.Text = ""
		txtEDF.Text = ""
		txtEnergy.Text = ""
		txtFat.Text = ""
		txtFe.Text = ""
		txtFibre.Text = ""
		txtFo.Text = ""
		txtIname.Text = ""
		txtMg.Text = ""
		txtNa.Text = ""
		txtNiacin.Text = ""
		txtProtein.Text = ""
		txtP.Text = ""
		txtk.Text = ""
		txtRetinol.Text = ""
		txtRiboflavin.Text = ""
		txtThiamin.Text = ""
		txtVitAR.Text = ""
		txtVitB12.Text = ""
		txtVitB6.Text = ""
		txtVitC.Text = ""
		txtVitD.Text = ""
		txtVitE.Text = ""
		txtWater.Text = ""
		txtZn.Text = ""



	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		txtAsh.Text = ""
		txtCa.Text = ""
		txtCarb.Text = ""
		txtCarotene.Text = ""
		txtCu.Text = ""
		txtEDF.Text = ""
		txtEnergy.Text = ""
		txtFat.Text = ""
		txtFe.Text = ""
		txtFibre.Text = ""
		txtFo.Text = ""
		txtIname.Text = ""
		txtMg.Text = ""
		txtNa.Text = ""
		txtNiacin.Text = ""
		txtProtein.Text = ""
		txtP.Text = ""
		txtk.Text = ""
		txtRetinol.Text = ""
		txtRiboflavin.Text = ""
		txtThiamin.Text = ""
		txtVitAR.Text = ""
		txtVitB12.Text = ""
		txtVitB6.Text = ""
		txtVitC.Text = ""
		txtVitD.Text = ""
		txtVitE.Text = ""
		txtWater.Text = ""
		txtZn.Text = ""
	End Sub
	Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
		Dim connection As New OleDbConnection()
		connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\NutValueChecker.mdb"
		connection.Open()
		Dim adp As OleDbDataAdapter = New OleDbDataAdapter("select * from Item", connection)
		Dim ds As DataSet = New DataSet()
		adp.Fill(ds)
		DataGridView1.DataSource = ds.Tables(0)
		connection.Close()
	End Sub

	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
		Me.FoodItemsTableAdapter.Fill(Me.NutValueCheckerDataSet.FoodItems)
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		Dim Pro As Double
		Dim carb As Double
		Dim fat As Double
		Dim Kcal As Double
		Dim pro2 As Double
		Dim Carb2 As Double
		Dim fat2 As Double
		Dim sum As Double
		Dim fat3 As Double
		Dim fat4 As Integer

		Kcal = Val(txtKcal.Text)
		Pro = Val(inputProtein.Text)
		carb = Val(inputCarb.Text)
		fat = Val(inputFat.Text)
		sum = Pro + carb + fat

		If sum = 100 Then

			fat2 = fat / 100
			fat3 = fat2 * Kcal
			fat4 = CInt(fat3 / 9)

			pro2 = ((Pro / 100) * Kcal) / 4
			Carb2 = ((carb / 100) * Kcal) / 4


			txtpro2.Text = CType((pro2), String)
			TextBox3.Text = CType(fat4, String)
			txtcarb2.Text = CType(Carb2, String)

		Else
			MsgBox(" Distribution percentage not equal to 100", vbOKCancel, "percentage error")

		End If


	End Sub

	Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
		If ComboBox1.SelectedIndex = 0 Then
			TextBox1.Text = "4-8% of daily Kcal from Protein, or 0.6-0.8g Protein/Kg Body weight"
		ElseIf ComboBox1.SelectedIndex = 1 Then
			TextBox1.Text = "16-25% Protein in the daily calorie distribution; and /or 1-2g Protein/Kg Body weight"
		ElseIf ComboBox1.SelectedIndex = 2 Then
			TextBox1.Text = "Greater than 14g/1000Kcal/day"
		ElseIf ComboBox1.SelectedIndex = 3 Then
			TextBox1.Text = "Lesser Than 10g/day"
		ElseIf ComboBox1.SelectedIndex = 4 Then
			TextBox1.Text = "20-29% of daily Kcal from Fat"
		ElseIf ComboBox1.SelectedIndex = 5 Then
			TextBox1.Text = "1500mg-4000mg of sodium is required"
		ElseIf ComboBox1.SelectedIndex = 6 Then
			TextBox1.Text = "40-49% of daily Kcal from Carbohydrate "
		End If
	End Sub


	Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
		Dim connection As New OleDbConnection()
		connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\NutValueChecker.mdb"
		Dim ds As New DataSet
		Dim da As OleDbDataAdapter
		Dim sql As String
		Dim ar(0 To 28) As Object
		Dim ans(0 To 28) As Object
		Dim dt As New DataTable
		Dim n As Integer
		Dim m As Integer
		Dim x As Double
		Dim Y As Double
		Dim Z As Double

		Dim relation As DataRow
		getinput()
		If txtname.Text = "" Then
			MsgBox("input the name of the Food Item", vbOK)
			txtname.Focus()
		Else
			Try

				connection.Open()
				sql = "SELECT * FROM Item where Itemname='" & txtname.Text & "'"
				da = New OleDbDataAdapter(sql, connection)
				da.Fill(ds, "ITEMS")
				dt = ds.Tables(0)
				If displayTable.Rows.Count = 0 Then
					displayTable = dt.Clone
				End If



				For n = 0 To 28
					ar(n) = CType(dt.Rows(0).Item(n), String)
				Next n
				For m = 0 To 28
					If m = 0 Then
						ar(0) = CInt(0)
					End If
					x = ar(m)
					Y = array_getdata(m)
					Z = x * Y
					ans(m) = Z / 100

				Next m
				ans(0) = CType(txtname.Text, String)


				'MsgBox(ds.Tables(0).Rows(0).Item(1))
				relation = displayTable.NewRow()
				relation.ItemArray = ans
				displayTable.Rows.Add(relation)
				displayTable.AcceptChanges()

				For n = 0 To 28
					ans(0) = 0
					arrayTotal(n) = CInt(arrayTotal(n) + ans(n))
				Next
				arrayTotal(0) = "Total"

				TotalTable = displayTable.Copy

				relation = TotalTable.NewRow()
				relation.ItemArray = arrayTotal
				TotalTable.Rows.Add(relation)
				TotalTable.AcceptChanges()
				arrayTotal(0) = 0

				DataGridView2.DataSource = TotalTable

			Catch ex As Exception
				Throw ex
			Finally
				'close connection
				connection.Close()

			End Try
		End If
	End Sub

	Public Sub getinput()
		array_getdata(0) = 0
		If txtiEDF.Text = "" Then
			array_getdata(1) = 0
		Else
			array_getdata(1) = CInt(txtiEDF.Text)
		End If
		If txtiEnergy.Text = "" Then
			array_getdata(2) = 0
		Else
			array_getdata(2) = CInt(txtiEnergy.Text)
		End If
		If txtiWater.Text = "" Then
			array_getdata(3) = 0
		Else
			array_getdata(3) = CInt(txtiWater.Text)
		End If
		If txtiProtein.Text = "" Then
			array_getdata(4) = 0
		Else
			array_getdata(4) = CInt(txtiProtein.Text)
		End If
		If txtiFat.Text = "" Then
			array_getdata(5) = 0
		Else
			array_getdata(5) = CInt(txtiFat.Text)
		End If
		If txtiCarb.Text = "" Then
			array_getdata(6) = 0
		Else
			array_getdata(6) = CInt(txtiCarb.Text)
		End If
		If txtiFibre.Text = "" Then
			array_getdata(7) = 0
		Else
			array_getdata(7) = CInt(txtiFibre.Text)
		End If
		If txtiAsh.Text = "" Then
			array_getdata(8) = 0
		Else
			array_getdata(8) = CInt(txtiAsh.Text)
		End If
		If txtiCa.Text = "" Then
			array_getdata(9) = 0
		Else
			array_getdata(9) = CInt(txtiCa.Text)
		End If
		If txtiFe.Text = "" Then
			array_getdata(10) = 0
		Else
			array_getdata(10) = CInt(txtiFe.Text)
		End If
		If txtiMg.Text = "" Then
			array_getdata(11) = 0
		Else
			array_getdata(11) = CInt(txtiMg.Text)
		End If
		If txtiP.Text = "" Then
			array_getdata(12) = 0
		Else
			array_getdata(12) = CInt(txtiP.Text)
		End If
		If txtiK.Text = "" Then
			array_getdata(13) = 0
		Else
			array_getdata(13) = CInt(txtiK.Text)
		End If
		If txtiNa.Text = "" Then
			array_getdata(14) = 0
		Else
			array_getdata(14) = CInt(txtiNa.Text)
		End If
		If txtiZn.Text = "" Then
			array_getdata(15) = 0
		Else
			array_getdata(15) = CInt(txtiZn.Text)
		End If
		If txtiCu.Text = "" Then
			array_getdata(16) = 0
		Else
			array_getdata(16) = CInt(txtiCu.Text)
		End If
		If txtiVitARae.Text = "" Then
			array_getdata(17) = 0
		Else
			array_getdata(17) = CInt(txtiVitARae.Text)
		End If
		If txtiRetinol.Text = "" Then
			array_getdata(18) = 0
		Else
			array_getdata(18) = CInt(txtiRetinol.Text)
		End If
		If txtiCarotene.Text = "" Then
			array_getdata(19) = 0
		Else
			array_getdata(19) = CInt(txtiCarotene.Text)
		End If
		If txtiVitD.Text = "" Then
			array_getdata(20) = 0
		Else
			array_getdata(20) = CInt(txtiVitD.Text)
		End If
		If txtiVitE.Text = "" Then
			array_getdata(21) = 0
		Else
			array_getdata(21) = CInt(txtiVitE.Text)
		End If
		If txtiThiamin.Text = "" Then
			array_getdata(21) = 0
		Else
			array_getdata(21) = CInt(txtiThiamin.Text)
		End If
		If txtiRiboflavin.Text = "" Then
			array_getdata(23) = 0
		Else
			array_getdata(23) = CInt(txtiRiboflavin.Text)
		End If
		If txtiNiacin.Text = "" Then
			array_getdata(24) = 0
		Else
			array_getdata(24) = CInt(txtiNiacin.Text)
		End If
		If txtiVitB6.Text = "" Then
			array_getdata(25) = 0
		Else
			array_getdata(25) = CInt(txtiVitB6.Text)
		End If
		If txtiFolate.Text = "" Then
			array_getdata(26) = 0
		Else
			array_getdata(26) = CInt(txtiFolate.Text)
		End If
		If txtiVitB12.Text = "" Then
			array_getdata(27) = 0
		Else
			array_getdata(27) = CInt(txtiVitB12.Text)
		End If
		If txtiVitC.Text = "" Then
			array_getdata(28) = 0
		Else
			array_getdata(28) = CInt(txtiVitC.Text)
		End If
	End Sub
End Class
