Imports System.Data

Namespace TegService
    '''---------------------------------------------------------------------------------------------------
    '''Autore.......: Antonello Lo Bianco
    '''Data ........: 02/05/2014
    '''Descrizione..: Classe di accesso al database.
    '''Progetto ....: ReminderServiceMail
    '''Revisioni....: Indicare qui di seguito le modifiche apportate
    '''---------------------------------+-------------------------+---------------------------------------
    '''Nome e Cognome                   Data                      Descrizione revisione
    '''---------------------------------+-------------------------+---------------------------------------
    '''Antonello Lo Bianco                  02/05/2014                Creazione classe.
    '''---------------------------------+-------------------------+---------------------------------------
    Public Class DB
        Public Shared cnn As System.Data.Common.DbConnection
        Public Shared cmd As System.Data.Common.DbCommand
        Public Shared da As System.Data.Common.DataAdapter
        Public Shared ds As System.Data.DataSet
        Public Shared exc As System.Exception
        Public Shared IsOleDb As Boolean = False

        Public Shared ReadOnly Property HasError() As Boolean
            Get
                If exc Is Nothing Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public Shared Sub ApriConnessione()
            exc = Nothing
            If cnn Is Nothing Then
                Dim cs As String = My.Settings.DataDB
                If cs.Contains("Provider=") Then
                    IsOleDb = True
                    cnn = New System.Data.OleDb.OleDbConnection(cs)
                Else
                    IsOleDb = False
                    cnn = New System.Data.SqlClient.SqlConnection(cs)
                End If
            End If
            If cnn.State = ConnectionState.Closed Then
                Try
                    cnn.Open()
                Catch ex As Exception
                    exc = ex
                End Try
            End If
        End Sub

        Public Shared Function CaricaDV(ByVal sqlString As String) As DataView
            Dim dv As DataView = Nothing

            exc = Nothing
            Try
                'Carica i dati da una tabella su database
                ApriConnessione()
                If cnn.State = ConnectionState.Open Then
                    Dim cmd As System.Data.Common.DbCommand = cnn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sqlString
                    If IsOleDb Then
                        da = New System.Data.OleDb.OleDbDataAdapter(cmd)
                    Else
                        da = New System.Data.SqlClient.SqlDataAdapter(cmd)
                    End If
                    ds = New DataSet
                    da.Fill(ds)
                    dv = New DataView(ds.Tables(0))
                    da.Dispose()
                    da = Nothing
                    cmd.Dispose()
                    cmd = Nothing
                End If
            Catch ex As Exception
                exc = ex
                dv = Nothing
            End Try
            Return dv
        End Function

        Public Shared Sub ChiudiConnessione()
            exc = Nothing
            If cnn IsNot Nothing Then
                Try
                    If ds IsNot Nothing Then
                        ds.Dispose()
                        ds = Nothing
                    End If
                    If da IsNot Nothing Then
                        da.Dispose()
                        da = Nothing
                    End If
                    If cmd IsNot Nothing Then
                        cmd.Dispose()
                        cmd = Nothing
                    End If
                    cnn.Close()
                    cnn.Dispose()
                Catch ex As Exception
                    exc = ex
                Finally
                    cnn = Nothing
                End Try
            End If
        End Sub

        Public Shared Function EseguiDML(ByVal sqlString As String) As Boolean
            Dim ret As Boolean = False

            exc = Nothing
            Try
                ApriConnessione()
                If cnn.State = ConnectionState.Open Then
                    Dim cmd As System.Data.Common.DbCommand = cnn.CreateCommand()
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = sqlString
                    If cmd.ExecuteNonQuery() > 0 Then
                        ret = True
                    End If
                    cmd.Dispose()
                    cmd = Nothing
                End If
            Catch ex As Exception
                exc = ex
            End Try
            Return ret
        End Function
    End Class
End Namespace