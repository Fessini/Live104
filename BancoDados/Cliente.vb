Imports System.Data.SqlClient

Public Class Cliente
    'ATRIBUTOS
    Private conexao As SqlConnection
    Private comando As SqlCommand
    Private da As SqlDataAdapter
    Private dr As SqlDataReader
    '
    'CONSTRUTOR DA CLASSE
    Sub New()
        conexao = New SqlConnection("Server=CODE-DESENV\SQLEXPRESS;Database=Live104;User Id=sa;Password=@mar147;")
    End Sub
    '
    'PROPRIEDADES
    Public Property IdCliente As Integer
    Public Property Nome As String
    Public Property Endereco As String
    Public Property Numero As String
    '
    'MÉTODOS
    Public Function NovoCliente() As Boolean
        Dim retorno As Boolean

        Try
            'PREPARANDO O COMANDO
            comando = New SqlCommand
            comando.Connection = conexao
            comando.CommandType = CommandType.StoredProcedure
            comando.CommandText = "PR_NOVO_CLIENTE"
            comando.Parameters.AddWithValue("@NOME", Nome)
            If Endereco.Length > 0 Then comando.Parameters.AddWithValue("@ENDERECO", Endereco)
            If Numero.Length > 0 Then comando.Parameters.AddWithValue("@NUMERO", Numero)
            '
            'ABRE O BANCO
            conexao.Open()
            '
            'EXECUTA O COMANDO
            retorno = comando.ExecuteNonQuery
            '
            'FECHA CONEXÃO
            conexao.Close()
            '
            'RETORNO
            Return retorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function AtualizaCliente() As Boolean
        Dim retorno As Boolean

        Try
            'PREPARANDO O COMANDO
            comando = New SqlCommand
            comando.Connection = conexao
            comando.CommandType = CommandType.StoredProcedure
            comando.CommandText = "PR_ATUALIZA_CLIENTE"
            comando.Parameters.AddWithValue("@ID", IdCliente)
            comando.Parameters.AddWithValue("@NOME", Nome)
            If Endereco.Length > 0 Then comando.Parameters.AddWithValue("@ENDERECO", Endereco)
            If Numero.Length > 0 Then comando.Parameters.AddWithValue("@NUMERO", Numero)
            '
            'ABRE O BANCO
            conexao.Open()
            '
            'EXECUTA O COMANDO
            retorno = comando.ExecuteNonQuery
            '
            'FECHA CONEXÃO
            conexao.Close()
            '
            'RETORNO
            Return retorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function DeleteCliente() As Boolean
        Dim retorno As Boolean

        Try
            'PREPARANDO O COMANDO
            comando = New SqlCommand
            comando.Connection = conexao
            comando.CommandType = CommandType.StoredProcedure
            comando.CommandText = "PR_DELETE_CLIENTE"
            comando.Parameters.AddWithValue("@ID", IdCliente)
            '
            'ABRE O BANCO
            conexao.Open()
            '
            'EXECUTA O COMANDO
            retorno = comando.ExecuteNonQuery
            '
            'FECHA CONEXÃO
            conexao.Close()
            '
            'RETORNO
            Return retorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function BuscaCliente() As Boolean
        Dim retorno As Boolean = False

        Try
            'PREPARANDO O COMANDO
            comando = New SqlCommand
            comando.Connection = conexao
            comando.CommandType = CommandType.StoredProcedure
            comando.CommandText = "PR_BUSCA_CLIENTE"
            comando.Parameters.AddWithValue("@ID", IdCliente)
            '
            'ABRE O BANCO
            conexao.Open()
            '
            'EXECUTA O COMANDO
            dr = comando.ExecuteReader
            Do While dr.Read
                Nome = dr("NOME")
                If Not IsDBNull(dr("ENDERECO")) Then Endereco = dr("ENDERECO")
                If Not IsDBNull(dr("NUMERO")) Then Numero = dr("NUMERO")
                retorno = True
            Loop
            '
            'FECHA CONEXÃO
            conexao.Close()
            '
            'RETORNO
            Return retorno
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function ListaCliente() As DataTable
        Dim tabela As New DataTable

        Try
            'PREPARANDO O COMANDO
            comando = New SqlCommand
            comando.Connection = conexao
            comando.CommandType = CommandType.StoredProcedure
            comando.CommandText = "PR_LISTA_CLIENTE"
            '
            'EXECUTA O COMANDO
            da = New SqlDataAdapter(comando)
            da.Fill(tabela)
            '
            '
            Return tabela
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
