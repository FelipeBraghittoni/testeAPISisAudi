Option Explicit On
Public Class repEmprSites
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850328"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320328"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0328"
#End Region

#Region "Variaveis de Ambiente"
    Private mvarEntidade As String
    Public Property Entidade() As String
        Get
            Return mvarEntidade
        End Get
        Set(value As String)
            mvarEntidade = value
        End Set
    End Property

    Private mvaridSite As Short
    Public Property idSite() As Short
        Get
            Return mvaridSite
        End Get
        Set(value As Short)
            mvaridSite = value
        End Set
    End Property

    Private mvarnomeSite As String
    Public Property nomeSite() As String
        Get
            Return mvarnomeSite
        End Get
        Set(value As String)
            mvarnomeSite = value
        End Set
    End Property

    Private mvarfone1Site As String
    Public Property fone1Site() As String
        Get
            Return mvarfone1Site
        End Get
        Set(value As String)
            mvarfone1Site = value
        End Set
    End Property

    Private mvarramal1Site As Short
    Public Property ramal1Site() As Short
        Get
            Return mvarramal1Site
        End Get
        Set(value As Short)
            mvarramal1Site = value
        End Set
    End Property

    Private mvarfone2Site As String
    Public Property fone2Site() As String
        Get
            Return mvarfone2Site
        End Get
        Set(value As String)
            mvarfone2Site = value
        End Set
    End Property

    Private mvarramal2Site As Short
    Public Property ramal2Site() As Short
        Get
            Return mvarramal2Site
        End Get
        Set(value As Short)
            mvarramal2Site = value
        End Set
    End Property

    Private mvarfaxSite As String
    Public Property faxSite() As String
        Get
            Return mvarfaxSite
        End Get
        Set(value As String)
            mvarfaxSite = value
        End Set
    End Property

    Private mvarenderecoSite As String
    Public Property enderecoSite() As String
        Get
            Return mvarenderecoSite
        End Get
        Set(value As String)
            mvarenderecoSite = value
        End Set
    End Property

    Private mvarcomplemSite As String
    Public Property complemSite() As String
        Get
            Return mvarcomplemSite
        End Get
        Set(value As String)
            mvarcomplemSite = value
        End Set
    End Property

    Private mvarbairroSite As String
    Public Property bairroSite() As String
        Get
            Return mvarbairroSite
        End Get
        Set(value As String)
            mvarbairroSite = value
        End Set
    End Property

    Private mvarpaisSite As String
    Public Property paisSite() As String
        Get
            Return mvarpaisSite
        End Get
        Set(value As String)
            mvarpaisSite = value
        End Set
    End Property

    Private mvarufSite As String
    Public Property ufSite() As String
        Get
            Return mvarufSite
        End Get
        Set(value As String)
            mvarufSite = value
        End Set
    End Property

    Private mvarcidadeSite As String
    Public Property cidadeSite() As String
        Get
            Return mvarcidadeSite
        End Get
        Set(value As String)
            mvarcidadeSite = value
        End Set
    End Property

    Private mvarcepSite As String
    Public Property cepSite() As String
        Get
            Return mvarcepSite
        End Get
        Set(value As String)
            mvarcepSite = value
        End Set
    End Property

    Private mvarregimeTrabSite As String
    Public Property regimeTrabSite() As String
        Get
            Return mvarregimeTrabSite
        End Get
        Set(value As String)
            mvarregimeTrabSite = value
        End Set
    End Property

    Private mvardiaInicTrabSite As String
    Public Property diainicTrabSite() As String
        Get
            Return mvardiaInicTrabSite
        End Get
        Set(value As String)
            mvardiaInicTrabSite = value
        End Set
    End Property

    Private mvardiaFimTrabSite As String
    Public Property diaFimTrabSite() As String
        Get
            Return mvardiaFimTrabSite
        End Get
        Set(value As String)
            mvardiaFimTrabSite = value
        End Set
    End Property

    Private mvarrefeitorioSite As String
    Public Property refeitorioSite() As String
        Get
            Return mvarrefeitorioSite
        End Get
        Set(value As String)
            mvarrefeitorioSite = value
        End Set
    End Property

    Private mvarhoraInicTrabSite As String
    Public Property horaInicTrabSite() As String
        Get
            Return mvarhoraInicTrabSite
        End Get
        Set(value As String)
            mvarhoraInicTrabSite = value
        End Set
    End Property

    Private mvarhoraInicDescSite As String
    Public Property horaInicDescSite() As String
        Get
            Return mvarhoraInicDescSite
        End Get
        Set(value As String)
            mvarhoraInicDescSite = value
        End Set
    End Property

    Private mvarhoraFimDescSite As String
    Public Property horaFimDescSite() As String
        Get
            Return mvarhoraFimDescSite
        End Get
        Set(value As String)
            mvarhoraFimDescSite = value
        End Set
    End Property

    Private mvarhoraFimTrabSite As String
    Public Property horaFimTrabSite() As String
        Get
            Return mvarhoraFimTrabSite
        End Get
        Set(value As String)
            mvarhoraFimTrabSite = value
        End Set
    End Property

    Private mvarCNPJ As String
    Public Property CNPJ() As String
        Get
            Return mvarCNPJ
        End Get
        Set(value As String)
            mvarCNPJ = value
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

    Private mvardtFim As String
    Public Property dtFim() As String
        Get
            Return mvardtFim
        End Get
        Set(value As String)
            mvardtFim = value
        End Set
    End Property

    Private mvartipoPessoa As String
    Public Property tipoPessoa() As String
        Get
            Return mvartipoPessoa
        End Get
        Set(value As String)
            mvartipoPessoa = value
        End Set
    End Property
#End Region

#Region "Conexão com banco"

    Public Db As ADODB.Connection
    Public trepEmprSites As ADODB.Recordset
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
        '* * 2 = OpenTable, índice IndiceNome     *
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
        trepEmprSites = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            trepEmprSites.Open("repEmprSites", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM repEmprSites ORDER BY Entidade, nomeSite"
                trepEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprSites.Close()

                If tipo = 1 Or tipo = 2 Then
                    trepEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprSites.Open("repEmprSites", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprSites - dbConecta" & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Short) As Integer
        '* ************************************************
        '* * Lê sequencialmente a tabela                  *
        '* *                                              *
        '* * vPrimVez - Se é a primeira leitura da tabela *
        '* * vIndice - Qual o índice que está acessando   *
        '* ************************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not trepEmprSites.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprSites.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprSites.EOF Then
            Entidade = trepEmprSites.Fields("Entidade").Value
            idSite = trepEmprSites.Fields("idSite").Value
            nomeSite = trepEmprSites.Fields("nomeSite").Value
            fone1Site = trepEmprSites.Fields("fone1Site").Value
            ramal1Site = trepEmprSites.Fields("ramal1Site").Value
            fone2Site = trepEmprSites.Fields("fone2Site").Value
            ramal2Site = trepEmprSites.Fields("ramal2Site").Value
            faxSite = trepEmprSites.Fields("faxSite").Value
            enderecoSite = trepEmprSites.Fields("enderecoSite").Value
            complemSite = trepEmprSites.Fields("complemSite").Value
            bairroSite = trepEmprSites.Fields("bairroSite").Value
            paisSite = trepEmprSites.Fields("paisSite").Value
            ufSite = trepEmprSites.Fields("ufSite").Value
            cidadeSite = trepEmprSites.Fields("cidadeSite").Value
            cepSite = trepEmprSites.Fields("cepSite").Value
            regimeTrabSite = trepEmprSites.Fields("regimeTrabSite").Value
            diainicTrabSite = trepEmprSites.Fields("diaInicTrab").Value
            diaFimTrabSite = trepEmprSites.Fields("diaFimTrab").Value
            refeitorioSite = trepEmprSites.Fields("refeitorioSite").Value
            horaInicTrabSite = trepEmprSites.Fields("horaInicTrabSite").Value
            horaInicDescSite = trepEmprSites.Fields("horaInicDescSite").Value
            horaFimDescSite = trepEmprSites.Fields("horaFimDescSite").Value
            horaFimTrabSite = trepEmprSites.Fields("horaFimTrabSite").Value
            CNPJ = trepEmprSites.Fields("CNPJ").Value
            dtInicio = trepEmprSites.Fields("dtInicio").Value
            dtFim = trepEmprSites.Fields("dtFim").Value
            tipoPessoa = trepEmprSites.Fields("tipoPessoa").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprSites - leSeq" & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
