Option Explicit On
Public Class repProjHorasConsumidas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850736"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320736"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0736"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
        End Set
    End Property

    Private mvarnomeProjeto As String
    Public Property nomeProjeto() As String
        Get
            Return mvarnomeProjeto
        End Get
        Set(value As String)
            mvarnomeProjeto = value
        End Set
    End Property

    Private mvaridAtividade As Short
    Public Property idAtividade() As Short
        Get
            Return mvaridAtividade
        End Get
        Set(value As Short)
            mvaridAtividade = value
        End Set
    End Property

    Private mvarnomeAtividade As String
    Public Property nomeAtividade() As Short
        Get
            Return mvarnomeAtividade
        End Get
        Set(value As Short)
            mvarnomeAtividade = value
        End Set
    End Property

    Private mvaridSAtividade As Short
    Public Property idSAtividade() As Short
        Get
            Return mvaridSAtividade
        End Get
        Set(value As Short)
            mvaridSAtividade = value
        End Set
    End Property

    Private mvarnomeSAtividade As String
    Public Property nomeSAtividade() As String
        Get
            Return mvarnomeSAtividade
        End Get
        Set(value As String)
            mvarnomeSAtividade = value
        End Set
    End Property

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

    Private mvaridDepto As Short
    Public Property idDepto() As Short
        Get
            Return mvaridDepto
        End Get
        Set(value As Short)
            mvaridDepto = value
        End Set
    End Property

    Private mvarnomeDepto As String
    Public Property nomeDepto() As String
        Get
            Return mvarnomeDepto
        End Get
        Set(value As String)
            mvarnomeDepto = value
        End Set
    End Property

    Private mvarperiodo As Date
    Public Property periodo() As Date
        Get
            Return mvarperiodo
        End Get
        Set(value As Date)
            mvarperiodo = value
        End Set
    End Property

    Private mvarhsAlocadas As Double
    Public Property hsAlocadas() As Double
        Get
            Return mvarhsAlocadas
        End Get
        Set(value As Double)
            mvarhsAlocadas = value
        End Set
    End Property

    Private mvarhsConsumidasMes As Double
    Public Property hsConsumidasMes() As Double
        Get
            Return mvarhsConsumidasMes
        End Get
        Set(value As Double)
            mvarhsConsumidasMes = value
        End Set
    End Property

    Private mvarhsConsumidasTotal As Double
    Public Property hsConsumidasTotal() As Double
        Get
            Return mvarhsConsumidasTotal
        End Get
        Set(value As Double)
            mvarhsConsumidasTotal = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public trepProjHorasConsumidas As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable repProjHorasConsumidas *
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
        trepProjHorasConsumidas = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjHorasConsumidas como OpenTable
                trepProjHorasConsumidas.Open("repProjHorasConsumidas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjHorasConsumidas.Close()

                Select Case tipo
                    Case 0      'Abre repProjHorasConsumidas como OpenTable
                        trepProjHorasConsumidas.Open("repProjHorasConsumidas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjHorasConsumidas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjHorasConsumidas.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjHorasConsumidas.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjHorasConsumidas.EOF Then
            idProjeto = trepProjHorasConsumidas.Fields("idProjeto").Value
            nomeProjeto = trepProjHorasConsumidas.Fields("nomeProjeto").Value
            nomeEmpresa = trepProjHorasConsumidas.Fields("nomeEmpresa").Value
            nomeDepto = trepProjHorasConsumidas.Fields("nomeDepto").Value
            idAtividade = trepProjHorasConsumidas.Fields("idAtividade").Value
            nomeAtividade = trepProjHorasConsumidas.Fields("nomeAtividade").Value
            idSAtividade = trepProjHorasConsumidas.Fields("idSAtividade").Value
            nomeSAtividade = trepProjHorasConsumidas.Fields("nomeSAtividade").Value
            idEmpresa = trepProjHorasConsumidas.Fields("idEmpresa").Value
            idDepto = trepProjHorasConsumidas.Fields("idDepto").Value
            periodo = trepProjHorasConsumidas.Fields("periodo").Value
            hsAlocadas = trepProjHorasConsumidas.Fields("hsAlocadas").Value
            hsConsumidasMes = trepProjHorasConsumidas.Fields("hsConsumidasMes").Value
            hsConsumidasTotal = trepProjHorasConsumidas.Fields("hsConsumidasTotal").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjHorasConsumidas - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaTxt() As String

        CarregaTxt = trepProjHorasConsumidas.Fields("descricAlocHoras").Value

    End Function
End Class
