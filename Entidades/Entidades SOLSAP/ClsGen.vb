Public Class authenticationRequest
    Public Property usuario As String
    Public Property password As String
End Class

Public Class authenticationResponse
    Public Property token As String
    Public Property user As user
End Class

Public Class user
    Public Property created_by As String
    Public Property usuario As String
    Public Property created As String
    Public Property active As Boolean
    Public Property profile_id As String
    Public Property password As String
    Public Property updated_by As String
    Public Property created_by_name As String
    Public Property identificacion As String
    Public Property full_name As String
    Public Property id As String
    Public Property updated As String
    Public Property updated_by_name As String
    Public Property email As String
End Class

Public Class ResponseDocuments
    Public Property log As List(Of String)
    Public Property msg As String
    Public Property type As String
End Class
