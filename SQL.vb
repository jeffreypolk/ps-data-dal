Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Text

Public Class SQL
    Implements IDisposable

#Region " Parameter Collection "

    Public Class DALParameterCollection
        Inherits CollectionBase
        Default Public Property Item(ByVal Index As Integer) As SqlParameter
            Get
                Return CType(List.Item(Index), SqlParameter)
            End Get
            Set(ByVal value As SqlParameter)
                List.Item(Index) = value
            End Set
        End Property
        Default Public Property Item(ByVal Name As String) As SqlParameter
            Get
                Dim ReturnValue As SqlParameter = Nothing
                For Each sp As SqlParameter In List
                    If sp.ParameterName = Name Then
                        ReturnValue = sp
                    End If
                Next
                Return ReturnValue
            End Get
            Set(ByVal value As SqlParameter)
                For Each sp As SqlParameter In List
                    If sp.ParameterName = value.ParameterName Then
                        sp = value
                    End If
                Next
            End Set
        End Property
        Public Function AddWithValue(ByVal Name As String, ByVal Value As Object) As SqlParameter
            Dim NewParameter As New SqlParameter
            NewParameter.ParameterName = Name
            NewParameter.Value = Value
            List.Add(NewParameter)
            Return NewParameter
        End Function
        Public Function AddWithValue(ByVal Name As String, ByVal value As Object, ByVal Direction As ParameterDirection, ByVal ParameterDBType As System.Data.DbType) As SqlParameter
            Dim NewParameter As New SqlParameter
            With NewParameter
                .ParameterName = Name
                .Value = value
                .Direction = Direction
                .DbType = ParameterDBType
            End With
            List.Add(NewParameter)
            Return NewParameter
        End Function
        Public Function AddWithValue(ByVal Name As String, ByVal value As Object, ByVal Direction As ParameterDirection, ByVal ParameterDBType As System.Data.DbType, ByVal Size As Integer) As SqlParameter
            Dim NewParameter As New SqlParameter
            With NewParameter
                .ParameterName = Name
                .Value = value
                .Direction = Direction
                .DbType = ParameterDBType
                .Size = Size
            End With
            List.Add(NewParameter)
            Return NewParameter
        End Function

        Public Overloads Sub Clear()
            MyBase.Clear()
        End Sub
    End Class
#End Region

    Private _CnnStr As String = String.Empty
    Private _Connection As SqlConnection = Nothing
    Private _Transaction As SqlTransaction = Nothing
    Private _Parameters As New DALParameterCollection

    Public ReadOnly Property Parameters() As DALParameterCollection
        Get
            Return _Parameters
        End Get
    End Property

    Public Sub Execute(ByVal SQL As String, ByRef DataTableParameter As DataTable)
        Dim ds As New DataSet
        Try
            Using cmd As SqlCommand = BuildCommand(SQL)
                Using da As New SqlDataAdapter
                    da.SelectCommand = cmd
                    da.Fill(ds)
                    If ds.Tables.Count > 0 Then
                        DataTableParameter = ds.Tables(0)
                    End If
                End Using
            End Using
        Catch ex As Exception
            Throw
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try
    End Sub

    Public Sub Execute(ByVal SQL As String, ByRef StringParameter As String)
        Try
            Using cmd As SqlCommand = BuildCommand(SQL)
                StringParameter = Convert.ToString(cmd.ExecuteScalar)
            End Using
        Catch ex As Exception
            Throw
        Finally

        End Try
    End Sub

    Public Sub Execute(ByVal SQL As String, ByRef BooleanParameter As Boolean)
        Try
            Using cmd As SqlCommand = BuildCommand(SQL)
                BooleanParameter = Convert.ToBoolean(cmd.ExecuteScalar)
            End Using
        Catch ex As Exception
            Throw
        Finally

        End Try
    End Sub

    Public Sub Execute(ByVal SQL As String)
        Try
            Using cmd As SqlCommand = BuildCommand(SQL)
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    Public Sub Execute(ByVal SQL As String, ByRef IntegerParameter As Integer)
        Try
            Using cmd As SqlCommand = BuildCommand(SQL)
                Dim SQLResult As String = cmd.ExecuteScalar
                If Not Integer.TryParse(SQLResult, IntegerParameter) Then
                    Throw New Exception("Dal.Execute() was unable to convert '" & SQLResult & "' to an Integer.")
                End If
            End Using
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    Private Function BuildCommand(ByVal SQL As String) As SqlCommand

        If IsNothing(_Connection) Then
            If _CnnStr = String.Empty Then
                Throw New Exception("No connection string defined")
            End If
            _Connection = New SqlConnection(_CnnStr)
            _Connection.Open()
        End If

        Using Command As New SqlCommand(SQL, _Connection)
            If _Transaction IsNot Nothing Then
                Command.Transaction = _Transaction
            End If

            Command.CommandType = GetCommandType(SQL)

            Command.Parameters.Clear()
            If _Parameters.Count > 0 Then
                For Each sp As SqlParameter In _Parameters
                    Command.Parameters.Add(sp)
                Next
            End If

            Return Command
        End Using
    End Function

    Private Function GetCommandType(ByVal SQL As String) As CommandType
        If SQL.IndexOf(" ") >= 0 Then
            Return CommandType.Text
        Else
            Return CommandType.StoredProcedure
        End If
    End Function

    Public Sub New()
        _CnnStr = System.Configuration.ConfigurationManager.AppSettings("cnnstr")
    End Sub
    Public Sub New(ByVal ConnectionString As String)
        _CnnStr = ConnectionString
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If _Connection IsNot Nothing Then
                    If _Connection.State <> ConnectionState.Closed Then _Connection.Close()
                    _Connection.Dispose()
                    _Connection = Nothing
                End If
            End If
        End If
        disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
    End Sub
#End Region

End Class
