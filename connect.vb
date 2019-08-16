Option Explicit On
Option Strict On
Imports System.Data.SqlClient
Module runconnect



	Public Sub filltb(query As String)
		Dim connection As SqlConnection = New SqlConnection()
		connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\NutValueChecker.mdb"
		connection.Open()
		Dim adp As SqlDataAdapter = New SqlDataAdapter(query, connection)
		Dim ds As DataSet = New DataSet()
		adp.Fill(ds)
	End Sub
End Module


