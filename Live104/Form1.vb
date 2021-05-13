Imports BancoDados

Public Class Form1
    Private objCliente As New Cliente

    Private Sub btnIncluir_Click(sender As Object, e As EventArgs) Handles btnIncluir.Click
        Try
            objCliente.Nome = txtNome.Text
            objCliente.Endereco = txtendereco.Text
            objCliente.Numero = txtNumero.Text
            If objCliente.NovoCliente = True Then
                MessageBox.Show("Cliente cadastrado.")
            Else
                MessageBox.Show("Cliente não cadastrado.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub btnEditar_Click(sender As Object, e As EventArgs) Handles btnEditar.Click
        Try
            objCliente.IdCliente = txtID.Text
            objCliente.Nome = txtNome.Text
            objCliente.Endereco = txtendereco.Text
            objCliente.Numero = txtNumero.Text
            If objCliente.AtualizaCliente = True Then
                MessageBox.Show("Cliente alterado.")
            Else
                MessageBox.Show("Cliente não alterado.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            objCliente.IdCliente = txtID.Text
            If objCliente.DeleteCliente = True Then
                MessageBox.Show("Cliente deletado.")
            Else
                MessageBox.Show("Cliente não deletado.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnConsultar_Click(sender As Object, e As EventArgs) Handles btnConsultar.Click
        Try
            objCliente.IdCliente = txtID.Text
            If objCliente.BuscaCliente = True Then
                txtNome.Text = objCliente.Nome
                txtendereco.Text = objCliente.Endereco
                txtNumero.Text = objCliente.Numero
            Else
                MessageBox.Show("Cliente não cadastrado.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnListar_Click(sender As Object, e As EventArgs) Handles btnListar.Click
        Try
            Dim tabela As DataTable = objCliente.ListaCliente
            dgvDados.DataSource = tabela
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
