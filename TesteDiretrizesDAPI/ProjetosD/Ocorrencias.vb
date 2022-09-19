Option Explicit On
Public Class Ocorrencias
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850703"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320703"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0703"
#End Region

#Region "Instancias de classes"
    Private mvarProjFluxo As ProjFluxo
    Public Property ProjFluxo() As ProjFluxo
        Get
            Return mvarProjFluxo
        End Get
        Set(value As ProjFluxo)
            mvarProjFluxo = value
        End Set
    End Property

    Private mvarProjFatosRelev As ProjFatosRelev
    Public Property ProjFatosRelev() As ProjFatosRelev
        Get
            Return mvarProjFatosRelev
        End Get
        Set(value As ProjFatosRelev)
            mvarProjFatosRelev = value
        End Set
    End Property

    Private mvarProjAlocAtiv As ProjAlocAtiv
    Public Property ProjAlocAtiv() As ProjAlocAtiv
        Get
            Return mvarProjAlocAtiv
        End Get
        Set(value As ProjAlocAtiv)
            mvarProjAlocAtiv = value
        End Set
    End Property


    Private mvarProjAlocSAtiv As ProjAlocSAtiv
    Public Property ProjAlocSAtiv() As ProjAlocSAtiv
        Get
            Return mvarProjAlocSAtiv
        End Get
        Set(value As ProjAlocSAtiv)
            mvarProjAlocSAtiv = value
        End Set
    End Property

    Private mvarProjAlocDetAtiv As ProjAlocDetAtiv
    Public Property ProjAlocDetAtiv() As ProjAlocDetAtiv
        Get
            Return mvarProjAlocDetAtiv
        End Get
        Set(value As ProjAlocDetAtiv)
            mvarProjAlocDetAtiv = value
        End Set
    End Property

    Private mvarProjAlocGerentes As ProjAlocGerentes
    Public Property ProjAlocGerentes() As ProjAlocGerentes
        Get
            Return mvarProjAlocGerentes
        End Get
        Set(value As ProjAlocGerentes)
            mvarProjAlocGerentes = value
        End Set
    End Property

#End Region

    Private Sub Class_Initialize()

        'create the mProjFluxo object when the Ocorrencias class is created
        mvarProjFluxo = New ProjFluxo
        'create the mProjFatosRelev object when the Ocorrencias class is created
        mvarProjFatosRelev = New ProjFatosRelev
        'create the mProjAlocAtiv object when the Ocorrencias class is created
        mvarProjAlocAtiv = New ProjAlocAtiv
        'create the mProjAlocSAtiv object when the Ocorrencias class is created
        mvarProjAlocSAtiv = New ProjAlocSAtiv
        'create the mProjAlocDetAtiv object when the Ocorrencias class is created
        mvarProjAlocDetAtiv = New ProjAlocDetAtiv
        'create the mProjAlocGerentes object when the Ocorrencias class is created
        mvarProjAlocGerentes = New ProjAlocGerentes

    End Sub

End Class
