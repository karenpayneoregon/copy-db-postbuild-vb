Imports System.IO

Namespace HelperClasses

    Public Class FileHelper
        Inherits DatabaseException

        Private Shared databaseFile As New FileInfo(FileName)
        ''' <summary>
        ''' Safe copy of database
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property FileName() As String
            Get
                Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database", "Database1.accdb")
            End Get
        End Property
        ''' <summary>
        ''' Database name to use in application folder
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property ProductionDatabaseFileName() As String
            Get
                Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, databaseFile.Name)
            End Get
        End Property
        ''' <summary>
        ''' Determines if the database exists in the application folder
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property Exists() As Boolean
            Get
                Return File.Exists(ProductionDatabaseFileName)
            End Get
        End Property
        ''' <summary>
        ''' Copy clean database to application folder
        ''' </summary>
        ''' <param name="overwrite"></param>
        Public Shared Sub CopyDatabase(Optional overwrite As Boolean = False)
            mHasException = True

            'Dim destinationPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, databaseFile.Name)

            Try

                If File.Exists(ProductionDatabaseFileName) Then
                    If overwrite Then
                        File.Delete(ProductionDatabaseFileName)
                    Else
                        mHasException = False
                        Exit Sub
                    End If

                End If

                databaseFile.CopyTo(ProductionDatabaseFileName)

                mHasException = False

            Catch ex As Exception
                mLastException = ex
                mHasException = True
            End Try

        End Sub

    End Class
End Namespace