Option Explicit On
Public Class Empresas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850311"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320311"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0311"
#End Region

#Region "Instancias de classe"
    Private mvarEmprEmpresas As EmprEmpresas
    Public Property EmprEmpresas() As EmprEmpresas
        Get
            Return mvarEmprEmpresas
        End Get
        Set(value As EmprEmpresas)
            mvarEmprEmpresas = value
        End Set
    End Property


    Private mvarEmprCargos As EmprCargos
    Public Property EmprCargos() As EmprCargos
        Get
            Return mvarEmprCargos
        End Get
        Set(value As EmprCargos)
            mvarEmprCargos = value
        End Set
    End Property


    Private mvarEmprAvaliacClie As EmprAvaliacClie
    Public Property EmprAvaliacClie() As EmprAvaliacClie
        Get
            Return mvarEmprAvaliacClie
        End Get
        Set(value As EmprCargos)
            mvarEmprAvaliacClie = value
        End Set
    End Property


    Private mvarEmprAvaliacForn As EmprAvaliacForn
    Public Property EmprAvaliacForn() As EmprAvaliacForn
        Get
            Return mvarEmprAvaliacClie
        End Get
        Set(value As EmprAvaliacForn)
            mvarEmprAvaliacForn = value
        End Set
    End Property


    Private mvarEmprAlocDeptoSite As EmprAlocDeptoSite
    Public Property EmprAlocDeptoSite() As EmprAlocDeptoSite
        Get
            Return mvarEmprAlocDeptoSite
        End Get
        Set(value As EmprAlocDeptoSite)
            mvarEmprAlocDeptoSite = value
        End Set
    End Property

#End Region

    Private Sub Class_Initialize()

        'create the mEmprEmpresas object when the Empresa class is created
        mvarEmprEmpresas = New EmprEmpresas
        'create the mEmprCargos object when the Cadastros class is created
        mvarEmprCargos = New EmprCargos
        'create the mEmprAvaliacClie object when the EmprEmpresas class is created
        mvarEmprAvaliacClie = New EmprAvaliacClie
        'create the mEmprAvaliacForn object when the EmprEmpresas class is created
        mvarEmprAvaliacForn = New EmprAvaliacForn
        'create the mEmprAlocDeptoSite object when the Empresa class is created
        mvarEmprAlocDeptoSite = New EmprAlocDeptoSite

    End Sub
End Class
