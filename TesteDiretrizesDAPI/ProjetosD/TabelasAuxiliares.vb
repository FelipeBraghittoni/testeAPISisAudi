Option Explicit On
Public Class TabelasAuxiliares
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850738"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320738"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0738"
#End Region

#Region "instancias de classes"
    Private mvarProjTipoStatus As ProjTipoStatus
    Public Property ProjTipoStatus() As ProjTipoStatus
        Get
            Return mvarProjTipoStatus
        End Get
        Set(value As ProjTipoStatus)
            mvarProjTipoStatus = value
        End Set
    End Property

    Private mvarProjTpFatos As ProjTpFatos
    Public Property ProjTpFatos() As ProjTpFatos
        Get
            Return mvarProjTpFatos
        End Get
        Set(value As ProjTpFatos)
            mvarProjTpFatos = value
        End Set
    End Property

    Private mvarProjTpModalidade As ProjTpModalidade
    Public Property ProjTpModalidade() As ProjTpModalidade
        Get
            Return mvarProjTpModalidade
        End Get
        Set(value As ProjTpModalidade)
            mvarProjTpModalidade = value
        End Set
    End Property

#End Region

    Private Sub Class_Initialize()

        'create the mProjTipoStatus object when the TabelasAuxiliares class is created
        mvarProjTipoStatus = New ProjTipoStatus
        'create the mProjTpFatos object when the TabelasAuxiliares class is created
        mvarProjTpFatos = New ProjTpFatos
        'create the mProjTpModalidade object when the TabelasAuxiliares class is created
        mvarProjTpModalidade = New ProjTpModalidade

    End Sub


End Class
