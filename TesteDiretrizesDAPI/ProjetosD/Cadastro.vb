Option Explicit On
Public Class Cadastro
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850702"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320702"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0702"
#End Region

#Region "Instancias de classes"
    Private mvarProjProjetos As ProjProjetos
    Public Property ProjProjetos() As ProjProjetos
        Get
            Return mvarProjProjetos
        End Get
        Set(value As ProjProjetos)
            mvarProjProjetos = value
        End Set
    End Property

    Private mvarProjDoctosProj As ProjDoctosProj
    Public Property ProjDoctosProj() As ProjDoctosProj
        Get
            Return mvarProjDoctosProj
        End Get
        Set(value As ProjDoctosProj)
            mvarProjDoctosProj = value
        End Set
    End Property

    Private mvarProjAtividades As ProjAtividades
    Public Property ProjAtividades() As ProjAtividades
        Get
            Return mvarProjAtividades
        End Get
        Set(value As ProjAtividades)
            mvarProjAtividades = value
        End Set
    End Property

    Private mvarProjSAtividades As ProjSAtividades
    Public Property ProjSAtividades() As ProjSAtividades
        Get
            Return mvarProjSAtividades
        End Get
        Set(value As ProjSAtividades)
            mvarProjSAtividades = value
        End Set
    End Property


    Private mvarProjDetAtiv As ProjDetAtiv
    Public Property ProjDetAtiv() As ProjDetAtiv
        Get
            Return mvarProjDetAtiv
        End Get
        Set(value As ProjDetAtiv)
            mvarProjDetAtiv = value
        End Set
    End Property

#End Region

    Private Sub Class_Initialize()

        'create the mProjProjetos object when the Cadastros class is created
        mvarProjProjetos = New ProjProjetos
        'create the mProjDoctosProj object when the Cadastros class is created
        mvarProjDoctosProj = New ProjDoctosProj
        'create the mProjAtividades object when the Cadastros class is created
        mvarProjAtividades = New ProjAtividades
        'create the mProjSAtividades object when the Cadastros class is created
        mvarProjSAtividades = New ProjSAtividades
        'create the mProjDetAtiv object when the Cadastros class is created
        mvarProjDetAtiv = New ProjDetAtiv

    End Sub
End Class
