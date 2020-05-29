Namespace HelperClasses
    Public Class DatabaseException
        Protected Shared mHasException As Boolean
        ''' <summary>
        ''' Indicates if there was an exception
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property HasException() As Boolean
            Get
                Return mHasException
            End Get
        End Property
        Protected Shared mLastException As Exception
        ''' <summary>
        ''' Last exception thrown if any
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property LastException() As Exception
            Get
                Return mLastException
            End Get
        End Property
        ''' <summary>
        ''' Last exception message
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property LastExceptionMessage As String
            Get
                If mLastException Is Nothing Then
                    Return ""
                Else
                    Return mLastException.Message
                End If
            End Get
        End Property
        ''' <summary>
        ''' Indicates if the application database is currently open or closed
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property DatbaseInuse() As Boolean
            Get
                If mLastException Is Nothing Then
                    Return False
                Else
                    Return mLastException.Message.Contains("The process cannot access the file")
                End If
            End Get
        End Property
        ''' <summary>
        ''' Indicates success or failure
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property IsSuccessFul As Boolean
            Get
                Return Not mHasException
            End Get
        End Property
    End Class
End Namespace