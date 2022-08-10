Option Explicit On
<ComClass(TabelasAuxiliares.ClassId, TabelasAuxiliares.InterfaceId, TabelasAuxiliares.EventsId)>
Public Class TabelasAuxiliares

    ' DirTpDiretriz As DirTpDiretriz
    Private mvarDirTpDiretriz As DirTpDiretriz
    Public Property DirTpDiretriz() As DirTpDiretriz
        Get
            Return mvarDirTpDiretriz
        End Get
        Set(ByVal value As DirTpDiretriz)
            mvarDirTpDiretriz = value
        End Set
    End Property
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850104"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320104"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0104"
#End Region

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        DirTpDiretriz = New DirTpDiretriz
    End Sub

    Public Sub NaoFazNada()
    End Sub

End Class
