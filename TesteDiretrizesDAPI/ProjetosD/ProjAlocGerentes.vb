Option Explicit On
Public Class ProjAlocGerentes

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850707"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320707"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0707"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridColaborador As Single
    Public Property idColaborador() As Single
        Get
            Return mvaridColaborador
        End Get
        Set(value As Single)
            mvaridColaborador = value
        End Set
    End Property

    Private mvarlocalTrabColab As Short
    Public Property localTrabColab() As Short
        Get
            Return mvarlocalTrabColab
        End Get
        Set(value As Short)
            mvarlocalTrabColab = value
        End Set
    End Property

    Private mvaridOrganiz As Short
    Public Property idOrganiz() As Short
        Get
            Return mvaridOrganiz
        End Get
        Set(value As Short)
            mvaridOrganiz = value
        End Set
    End Property

    Private mvaridSetorOrg As Short
    Public Property idSetorOrg() As Short
        Get
            Return mvaridSetorOrg
        End Get
        Set(value As Short)
            mvaridSetorOrg = value
        End Set
    End Property

    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
        End Set
    End Property

    Private mvaralocacaoProj As Short
    Public Property alocacaoProj() As Short
        Get
            Return mvaralocacaoProj
        End Get
        Set(value As Short)
            mvaralocacaoProj = value
        End Set
    End Property

    Private mvarhorasalocadas As Short
    Public Property horasalocadas() As Short
        Get
            Return mvarhorasalocadas
        End Get
        Set(value As Short)
            mvarhorasalocadas = value
        End Set
    End Property

    Private mvaridApuracaoHoras As Short
    Public Property idApuracaoHoras() As Short
        Get
            Return mvaridApuracaoHoras
        End Get
        Set(value As Short)
            mvaridApuracaoHoras = value
        End Set
    End Property

    Private mvarhorasAlocadasMax As Short
    Public Property horasAlocadasMax() As Short
        Get
            Return mvarhorasAlocadasMax
        End Get
        Set(value As Short)
            mvarhorasAlocadasMax = value
        End Set
    End Property

    Private mvardtInicioProjeto As Date
    Public Property dtInicioProjeto() As Date
        Get
            Return mvardtInicioProjeto
        End Get
        Set(value As Date)
            mvardtInicioProjeto = value
        End Set
    End Property

    Private mvardtFinalProjeto As Date
    Public Property dtFinalProjeto() As Date
        Get
            Return mvardtFinalProjeto
        End Get
        Set(value As Date)
            mvardtFinalProjeto = value
        End Set
    End Property

    Private mvargerenteCom As String
    Public Property gerenteCom() As Date
        Get
            Return mvargerenteCom
        End Get
        Set(value As Date)
            mvargerenteCom = value
        End Set
    End Property

    Private mvargerenteTec As String
    Public Property gerenteTec() As String
        Get
            Return mvargerenteTec
        End Get
        Set(value As String)
            mvargerenteTec = value
        End Set
    End Property

    Private mvargerenteProj As String
    Public Property gerenteProj() As String
        Get
            Return mvargerenteProj
        End Get
        Set(value As String)
            mvargerenteProj = value
        End Set
    End Property

    Private mvarcrachaCliente As String
    Public Property crachaCliente() As String
        Get
            Return mvarcrachaCliente
        End Get
        Set(value As String)
            mvarcrachaCliente = value
        End Set
    End Property

    Private mvarnumCrachaCliente As Double
    Public Property numCrachaCliente() As Double
        Get
            Return mvarnumCrachaCliente
        End Get
        Set(value As Double)
            mvarnumCrachaCliente = value
        End Set
    End Property

    Private mvarviaCrachaCliente As Short
    Public Property viaCrachaCliente() As Double
        Get
            Return mvarviaCrachaCliente
        End Get
        Set(value As Double)
            mvarviaCrachaCliente = value
        End Set
    End Property

    Private mvardtVencCrachaCliente As Date
    Public Property dtVencCrachaCliente() As Date
        Get
            Return mvardtVencCrachaCliente
        End Get
        Set(value As Date)
            mvardtVencCrachaCliente = value
        End Set
    End Property
#End Region

#Region "Conexao com Banco"
    Public Db As ADODB.Connection
    Public tProjAlocGerentes As ADODB.Recordset
#End Region
    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable                     *
        '* * 1 = OpenDynaset                   *
        '* *************************************

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
        tProjAlocGerentes = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjAlocGerentes.Open("OcorAlocProjetos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjAlocGerentes.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocGerentes.Close()

                If tipo = 1 Then
                    tProjAlocGerentes.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjAlocGerentes.Open("OcorAlocProjetos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjAlocGerentes - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjAlocGerentes.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjAlocGerentes.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjAlocGerentes.EOF Then
            idColaborador = tProjAlocGerentes.Fields("idColaborador").Value
            localTrabColab = tProjAlocGerentes.Fields("localTrabColab").Value
            idOrganiz = tProjAlocGerentes.Fields("idOrganiz").Value
            idSetorOrg = tProjAlocGerentes.Fields("idSetorOrg").Value
            idProjeto = tProjAlocGerentes.Fields("idProjeto").Value
            alocacaoProj = tProjAlocGerentes.Fields("alocacaoProj").Value
            horasalocadas = tProjAlocGerentes.Fields("horasalocadas").Value
            idApuracaoHoras = tProjAlocGerentes.Fields("idApuracaoHoras").Value
            horasAlocadasMax = tProjAlocGerentes.Fields("horasAlocadasMax").Value
            dtInicioProjeto = tProjAlocGerentes.Fields("dtInicioProjeto").Value
            dtFinalProjeto = tProjAlocGerentes.Fields("dtFinalProjeto").Value
            gerenteCom = tProjAlocGerentes.Fields("gerenteCom").Value
            gerenteTec = tProjAlocGerentes.Fields("gerenteTec").Value
            gerenteProj = tProjAlocGerentes.Fields("gerenteProj").Value
            crachaCliente = tProjAlocGerentes.Fields("crachaCliente").Value
            numCrachaCliente = tProjAlocGerentes.Fields("numCrachaCliente").Value
            viaCrachaCliente = tProjAlocGerentes.Fields("viaCrachaCliente").Value
            dtVencCrachaCliente = tProjAlocGerentes.Fields("dtVencCrachaCliente").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjAlocGerentes - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(colab As Single, vLocal As Short, organiz As Short, setor As Short, projeto As Double, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjAlocGerentes.                    *
        '* *                                      *
        '* * colab      = identif. para pesquisa  *
        '* * vlocal     = identif. para pesquisa  *
        '* * organiz    = identif. para pesquisa  *
        '* * setor      = identif. para pesquisa  *
        '* * projeto    = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM OcorAlocProjetos WHERE idColaborador = " & colab
        vSelect = vSelect & " AND localTrabColab = " & vLocal
        vSelect = vSelect & " AND idOrganiz = " & organiz
        vSelect = vSelect & " AND idSetorOrg = " & setor
        vSelect = vSelect & " AND idProjeto = " & projeto
        tProjAlocGerentes.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjAlocGerentes.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idColaborador = tProjAlocGerentes.Fields("idColaborador").Value
            localTrabColab = tProjAlocGerentes.Fields("localTrabColab").Value
            idOrganiz = tProjAlocGerentes.Fields("idOrganiz").Value
            idSetorOrg = tProjAlocGerentes.Fields("idSetorOrg").Value
            idProjeto = tProjAlocGerentes.Fields("idProjeto").Value
            alocacaoProj = tProjAlocGerentes.Fields("alocacaoProj").Value
            horasalocadas = tProjAlocGerentes.Fields("horasalocadas").Value
            idApuracaoHoras = tProjAlocGerentes.Fields("idApuracaoHoras").Value
            horasAlocadasMax = tProjAlocGerentes.Fields("horasAlocadasMax").Value
            dtInicioProjeto = tProjAlocGerentes.Fields("dtInicioProjeto").Value
            dtFinalProjeto = tProjAlocGerentes.Fields("dtFinalProjeto").Value
            gerenteCom = tProjAlocGerentes.Fields("gerenteCom").Value
            gerenteTec = tProjAlocGerentes.Fields("gerenteTec").Value
            gerenteProj = tProjAlocGerentes.Fields("gerenteProj").Value
            crachaCliente = tProjAlocGerentes.Fields("crachaCliente").Value
            numCrachaCliente = tProjAlocGerentes.Fields("numCrachaCliente").Value
            viaCrachaCliente = tProjAlocGerentes.Fields("viaCrachaCliente").Value
            dtVencCrachaCliente = tProjAlocGerentes.Fields("dtVencCrachaCliente").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocGerentes.Close()
                tProjAlocGerentes.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjAlocGerentes - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function
End Class
