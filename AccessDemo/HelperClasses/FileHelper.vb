Imports System.IO

Namespace HelperClasses

    Public Class FileHelper
        Inherits DatabaseException

        Private Shared databaseFile As New FileInfo(FileName)
        Public Shared ReadOnly Property FileName() As String
            Get
                Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database", "Database1.accdb")
            End Get
        End Property
        ''' <summary>
        ''' Copy clean database to application folder
        ''' </summary>
        ''' <param name="overwrite"></param>
        Public Shared Sub CopyDatabase(Optional overwrite As Boolean = False)
            mHasException = True

            Dim destinationPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, databaseFile.Name)

            Try

                If File.Exists(destinationPath) Then
                    If overwrite Then
                        File.Delete(destinationPath)
                    Else
                        mHasException = False
                        Exit Sub
                    End If

                End If

                databaseFile.CopyTo(destinationPath)

                mHasException = False

            Catch ex As Exception
                mLastException = ex
                mHasException = True
            End Try

        End Sub

    End Class
End Namespace