Option Explicit On
Public Class repEmprCompSoc
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850322"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320322"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0322"
#End Region

#Region "Variaveis de Ambiente"

    Private mvaridEmpresa As Short
    Public Property idEmpresa() As Short
        Get
            Return mvaridEmpresa
        End Get
        Set(value As Short)
            mvaridEmpresa = value
        End Set
    End Property


    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvarnomeEmpresa
        End Get
        Set(value As String)
            mvarnomeEmpresa = value
        End Set
    End Property


    Private mvarDtpSocio As String
    Public Property DtpSocio() As String
        Get
            Return mvarDtpSocio
        End Get
        Set(value As String)
            mvarDtpSocio = value
        End Set
    End Property


    Private mvaridSocio As Single
    Public Property idSocio() As Single
        Get
            Return mvaridSocio
        End Get
        Set(value As Single)
            mvaridSocio = value
        End Set
    End Property


    Private mvarnomeSocio As String
    Public Property nomeSocio() As String
        Get
            Return mvarnomeSocio
        End Get
        Set(value As String)
            mvarnomeSocio = value
        End Set
    End Property


    Private mvarenderecoSocio As String
    Public Property enderecoSocio() As String
        Get
            Return mvarenderecoSocio
        End Get
        Set(value As String)
            mvarenderecoSocio = value
        End Set
    End Property


    Private mvarcomplemSocio As String
    Public Property complemSocio() As String
        Get
            Return mvarcomplemSocio
        End Get
        Set(value As String)
            mvarcomplemSocio = value
        End Set
    End Property


    Private mvarbairroSocio As String
    Public Property bairroSocio() As String
        Get
            Return mvarbairroSocio
        End Get
        Set(value As String)
            mvarbairroSocio = value
        End Set
    End Property


    Private mvarpaisSocio As String
    Public Property paisSocio() As String
        Get
            Return mvarpaisSocio
        End Get
        Set(value As String)
            mvarpaisSocio = value
        End Set
    End Property


    Private mvarestadoSocio As String
    Public Property estadoSocio() As String
        Get
            Return mvarestadoSocio
        End Get
        Set(value As String)
            mvarestadoSocio = value
        End Set
    End Property


    Private mvarcidadeSocio As String
    Public Property cidadeSocio() As String
        Get
            Return mvarcidadeSocio
        End Get
        Set(value As String)
            mvarcidadeSocio = value
        End Set
    End Property


    Private mvarDcepSocio As String
    Public Property DcepSocio() As String
        Get
            Return mvarDcepSocio
        End Get
        Set(value As String)
            mvarDcepSocio = value
        End Set
    End Property


    Private mvareMailSocio As String
    Public Property eMailSocio() As String
        Get
            Return mvareMailSocio
        End Get
        Set(value As String)
            mvareMailSocio = value
        End Set
    End Property


    Private mvarDfoneSocio As String
    Public Property DfoneSocio() As String
        Get
            Return mvarDfoneSocio
        End Get
        Set(value As String)
            mvarDfoneSocio = value
        End Set
    End Property


    Private mvarDcelSocio As String
    Public Property DcelSocio() As String
        Get
            Return mvarDcelSocio
        End Get
        Set(value As String)
            mvarDcelSocio = value
        End Set
    End Property


    Private mvarCPFSocio As String
    Public Property CPFSocio() As String
        Get
            Return mvarCPFSocio
        End Get
        Set(value As String)
            mvarCPFSocio = value
        End Set
    End Property


    Private mvarRGSocio As String
    Public Property RGSocio() As String
        Get
            Return mvarRGSocio
        End Get
        Set(value As String)
            mvarRGSocio = value
        End Set
    End Property


    Private mvaremissorRGSocio As String
    Public Property emissorRGSocio() As String
        Get
            Return mvaremissorRGSocio
        End Get
        Set(value As String)
            mvaremissorRGSocio = value
        End Set
    End Property


    Private mvarCNPJSocio As String
    Public Property CNPJSocio() As String
        Get
            Return mvarCNPJSocio
        End Get
        Set(value As String)
            mvarCNPJSocio = value
        End Set
    End Property


    Private mvartpParticip As String
    Public Property tpParticip() As String
        Get
            Return mvartpParticip
        End Get
        Set(value As String)
            mvartpParticip = value
        End Set
    End Property


    Private mvarparticipacao As String
    Public Property participacao() As String
        Get
            Return mvarparticipacao
        End Get
        Set(value As String)
            mvarparticipacao = value
        End Set
    End Property


    Private mvarnomeMoeda As String
    Public Property nomeMoeda() As String
        Get
            Return mvarnomeMoeda
        End Get
        Set(value As String)
            mvarnomeMoeda = value
        End Set
    End Property


    Private mvardtInicio As Date
    Public Property dtInicio() As Date
        Get
            Return mvardtInicio
        End Get
        Set(value As Date)
            mvardtInicio = value
        End Set
    End Property


    Private mvardDtFim As String
    Public Property dDtFim() As String
        Get
            Return mvardDtFim
        End Get
        Set(value As String)
            mvardDtFim = value
        End Set
    End Property

#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public trepEmprCompSoc As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable, índice ChavePrimaria  *
        '* * 1 = OpenDynaset                      *
        '* ****************************************

        On Error GoTo EdbConecta

        dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            '* ********************************************
            '* Cria uma instância de ADODB.Connection:    *
            '* ********************************************
            Dim strConnect As String
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        trepEmprCompSoc = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            trepEmprCompSoc.Open("repEmprCompSoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                trepEmprCompSoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprCompSoc.Close()

                If tipo = 1 Then
                    trepEmprCompSoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprCompSoc.Open("repEmprCompSoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprCompSoc - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Short) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not trepEmprCompSoc.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprCompSoc.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprCompSoc.EOF Then
            idEmpresa = trepEmprCompSoc.Fields("idEmpresa").Value
            DtpSocio = trepEmprCompSoc.Fields("DtpSocio").Value
            idSocio = trepEmprCompSoc.Fields("idSocio").Value
            nomeSocio = trepEmprCompSoc.Fields("nomeSocio").Value
            enderecoSocio = trepEmprCompSoc.Fields("enderecoSocio").Value
            complemSocio = trepEmprCompSoc.Fields("complemSocio").Value
            bairroSocio = trepEmprCompSoc.Fields("bairroSocio").Value
            paisSocio = trepEmprCompSoc.Fields("paisSocio").Value
            estadoSocio = trepEmprCompSoc.Fields("estadoSocio").Value
            cidadeSocio = trepEmprCompSoc.Fields("cidadeSocio").Value
            DcepSocio = trepEmprCompSoc.Fields("DcepSocio").Value
            eMailSocio = trepEmprCompSoc.Fields("eMailSocio").Value
            DfoneSocio = trepEmprCompSoc.Fields("DfoneSocio").Value
            DcelSocio = trepEmprCompSoc.Fields("DcelSocio").Value
            CPFSocio = trepEmprCompSoc.Fields("CPFSocio").Value
            RGSocio = trepEmprCompSoc.Fields("RGSocio").Value
            emissorRGSocio = trepEmprCompSoc.Fields("emissorRGSocio").Value
            CNPJSocio = trepEmprCompSoc.Fields("CNPJSocio").Value
            tpParticip = trepEmprCompSoc.Fields("tpParticip").Value
            participacao = trepEmprCompSoc.Fields("participacao").Value
            nomeMoeda = trepEmprCompSoc.Fields("nomeMoeda").Value
            dtInicio = trepEmprCompSoc.Fields("dtInicio").Value
            dDtFim = trepEmprCompSoc.Fields("dDtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprCompSoc - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function extraiMemo() As String
        'Extrai o conteúdo do campo memo txtDescric

        extraiMemo = trepEmprCompSoc.Fields("descric").Value

    End Function

End Class
