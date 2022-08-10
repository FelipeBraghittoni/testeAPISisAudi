Option Explicit On
<ComClass(DirCadastros.ClassId, DirCadastros.InterfaceId, DirCadastros.EventsId)>
Public Class DirCadastros

    'Public DirDoctosDir As DirDoctosDir
    Private mvarDirDoctosDir As DirDoctosDir
    Public Property DirDoctosDir() As DirDoctosDir
        Get
            Return mvarDirDoctosDir
        End Get
        Set(ByVal value As DirDoctosDir)
            mvarDirDoctosDir = value
        End Set
    End Property

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850103"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320103"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0103"
#End Region

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        DirDoctosDir = New DirDoctosDir
    End Sub

    Public Sub NaoFazNada()
    End Sub

End Class
