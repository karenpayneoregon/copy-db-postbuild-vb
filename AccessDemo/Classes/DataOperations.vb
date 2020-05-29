Imports System.Data.OleDb
Imports AccessDemo.HelperClasses

Namespace Classes
    ''' <summary>
    ''' Code to validate the copy database from sub folder to
    ''' application folder works.
    ''' </summary>
    Public Class DataOperations
        Inherits DatabaseException

        Private Shared ConnectionString As String =
                           "Provider=Microsoft.ACE.OLEDB.12.0;" &
                           "Data Source=Database1.accdb"
        ''' <summary>
        ''' Standard read joined tables
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function LoadCustomers() As DataTable


            Dim dt As New DataTable

            Dim selectStatement =
                    <SQL>
                    SELECT 
                        Cust.CustomerIdentifier, 
                        Cust.CompanyName, 
                        Cust.ContactId, 
                        CT.ContactTitle, 
                        Con.FirstName, 
                        Con.LastName, 
                        Cust.City, 
                        Cust.PostalCode, 
                        Cust.CountryIdentifier, 
                        Cust.ContactTypeIdentifier, 
                        CO.Name AS CountryName
                    FROM ((Customers AS Cust 
                        INNER JOIN Contacts AS Con ON Cust.ContactId = Con.ContactId) 
                        INNER JOIN ContactType AS CT ON Cust.ContactTypeIdentifier = CT.ContactTypeIdentifier) 
                        INNER JOIN Countries AS CO ON Cust.CountryIdentifier = CO.CountryIdentifier;
                    </SQL>.Value

            Using cn As New OleDbConnection With {.ConnectionString = ConnectionString}
                Using cmd As New OleDbCommand With {.CommandText = selectStatement, .Connection = cn}
                    cn.Open()

                    dt.Load(cmd.ExecuteReader())
                    dt.Columns("CustomerIdentifier").ColumnMapping = MappingType.Hidden
                    dt.Columns("ContactId").ColumnMapping = MappingType.Hidden
                    dt.Columns("CountryIdentifier").ColumnMapping = MappingType.Hidden
                    dt.Columns("ContactTypeIdentifier").ColumnMapping = MappingType.Hidden
                End Using
            End Using

            Return dt

        End Function
        ''' <summary>
        ''' Simple example for creating a table.
        ''' Note inside of the access table the boolean column
        ''' will not show true/false but numeric values yet
        ''' does when read into a DataTable.
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function CreateTable() As Boolean
            mHasException = False

            Dim tableName = "Person"

            Using cn = New OleDbConnection(ConnectionString)
                Using cmd = New OleDbCommand("", cn)
                    cmd.CommandText =
                        $"CREATE TABLE {tableName} ([Id] COUNTER, [FirstName] TEXT(25)," &
                         "[LastName] TEXT(255), [ActiveAccount] YESNO, [JoinYear] INT)"

                    Try

                        cn.Open()
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = $"INSERT INTO {tableName} " &
                                          "(FirstName,LastName,ActiveAccount,JoinYear) VALUES ('Karen','Payne',?,?)"

                        cmd.Parameters.AddWithValue("?", 0)
                        cmd.Parameters.AddWithValue("?", Now.Year)
                        cmd.ExecuteNonQuery()

                        '
                        ' Code that can be used to verify all is well,
                        ' uncomment, place a breakpoint on the return statement,
                        ' when breakpoint is hit hover over dt to see the row
                        '
                        'cmd.CommandText = $"SELECT FirstName,LastName,ActiveAccount,JoinYear FROM {tableName}"
                        'Dim dt As New DataTable
                        'dt.Load(cmd.ExecuteReader())

                        Return True

                    Catch e1 As Exception
                        mHasException = True
                        mLastException = e1
                        Return False
                    End Try

                End Using
            End Using
        End Function
    End Class
End Namespace