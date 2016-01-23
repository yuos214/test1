Imports Microsoft.VisualBasic.ControlChars

Module Module2

    Sub Test()

        Dim cn As New System.Data.OleDb.OleDbConnection

        Dim command As System.Data.OleDb.OleDbCommand

        Dim reader As System.Data.OleDb.OleDbDataReader



        'Accessファイルの保管場所

        Dim FilePath As String = "D:\stocks\stockshistory1.accdb"



        Try

            cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & Quote & FilePath & Quote & ";"



            command = cn.CreateCommand



            command.CommandText = "SELECT * FROM 1301"



            cn.Open()



            reader = command.ExecuteReader()

            '読み込んだデータを表示

            While reader.Read() = True

                Debug.WriteLine(reader(0))

            End While



            cn.Close()

            command.Dispose()

            cn.Dispose()

        Catch ex As Exception

            MessageBox.Show(ex.Message, "エラー")

        End Try

    End Sub

End Module