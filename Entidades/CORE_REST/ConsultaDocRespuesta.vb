
Public Class ConsultaDocRespuesta
    Public Property result As ResultDocRespuesta
    Public Property accessKey As String
    Public Property authorizationDate As String
    Public Property authorizationNumber As String
    Public Property state As String
    Public Property statusMsg As String
    Public Property pdf As String
    Public Property xml As String
End Class

Public Class ResultDocRespuesta
    Public Property code As String
    Public Property message As String
End Class

