Option Explicit On
Public Class ProjHorasConsumidas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850716"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320716"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0716"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridEmpresa As Short
    Public Property idEmpresa() As Short
        Get
            Return mvaridEmpresa
        End Get
        Set(value As Short)
            mvaridEmpresa = value
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

    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
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

    Private mvaridSAtividade As Short
    Public Property idSAtividade() As Short
        Get
            Return mvaridSAtividade
        End Get
        Set(value As Short)
            mvaridSAtividade = value
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

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tProjHorasConsumidas As ADODB.Recordset
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
        tProjHorasConsumidas = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjHorasConsumidas.Open("ProjHorasConsumidas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 1 Then    'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjHorasConsumidas.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjHorasConsumidas.Open("ProjHorasConsumidas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjHorasConsumidas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjHorasConsumidas.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjHorasConsumidas.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjHorasConsumidas.EOF Then
            idEmpresa = tProjHorasConsumidas.Fields("idEmpresa").Value
            idDepto = tProjHorasConsumidas.Fields("idDepto").Value
            idProjeto = tProjHorasConsumidas.Fields("idProjeto").Value
            idAtividade = tProjHorasConsumidas.Fields("idAtividade").Value
            idSAtividade = tProjHorasConsumidas.Fields("idSAtividade").Value
            periodo = tProjHorasConsumidas.Fields("periodo").Value
            hsAlocadas = tProjHorasConsumidas.Fields("hsAlocadas").Value
            hsConsumidasMes = tProjHorasConsumidas.Fields("hsConsumidasMes").Value
            hsConsumidasTotal = tProjHorasConsumidas.Fields("hsConsumidasTotal").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjHorasConsumidas - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empr As Short, Depto As Short, proj As Double, ativ As Short, sAtiv As Short, prdo As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjHorasConsumidas.                 *
        '* *                                      *
        '* * empr       = identif. para pesquisa  *
        '* * depto      = identif. para pesquisa  *
        '* * proj       = identif. para pesquisa  *
        '* * ativ       = identif. para pesquisa  *
        '* * sAtiv      = identif. para pesquisa  *
        '* * prdo       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelec As String

        On Error GoTo Elocaliza
        vSelec = "SELECT * FROM ProjHorasConsumidas WHERE idEmpresa = " & empr
        vSelec = vSelec & " AND idDepto = " & Depto
        vSelec = vSelec & " AND idProjeto = " & proj
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        vSelec = vSelec & " AND Convert(datetime, periodo, 112) = '" & String.Format(prdo, "yyyymmdd") & "'"
        tProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjHorasConsumidas.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tProjHorasConsumidas.Fields("idEmpresa").Value
            idDepto = tProjHorasConsumidas.Fields("idDepto").Value
            idProjeto = tProjHorasConsumidas.Fields("idProjeto").Value
            idAtividade = tProjHorasConsumidas.Fields("idAtividade").Value
            idSAtividade = tProjHorasConsumidas.Fields("idSAtividade").Value
            periodo = tProjHorasConsumidas.Fields("periodo").Value
            hsAlocadas = tProjHorasConsumidas.Fields("hsAlocadas").Value
            hsConsumidasMes = tProjHorasConsumidas.Fields("hsConsumidasMes").Value
            hsConsumidasTotal = tProjHorasConsumidas.Fields("hsConsumidasTotal").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjHorasConsumidas.Close()
                tProjHorasConsumidas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjHorasConsumidas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tProjHorasConsumidas.AddNew()
        tProjHorasConsumidas.Fields("idEmpresa").Value = idEmpresa
        tProjHorasConsumidas.Fields("idDepto").Value = idDepto
        tProjHorasConsumidas.Fields("idProjeto").Value = idProjeto
        tProjHorasConsumidas.Fields("idAtividade").Value = idAtividade
        tProjHorasConsumidas.Fields("idSAtividade").Value = idSAtividade
        tProjHorasConsumidas.Fields("periodo").Value = periodo
        tProjHorasConsumidas.Fields("hsAlocadas").Value = hsAlocadas
        tProjHorasConsumidas.Fields("hsConsumidasMes").Value = hsConsumidasMes
        tProjHorasConsumidas.Fields("hsConsumidasTotal").Value = hsConsumidasTotal
        tProjHorasConsumidas.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjHorasConsumidas - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tProjHorasConsumidas.Fields("hsAlocadas").Value = hsAlocadas
        tProjHorasConsumidas.Fields("hsConsumidasMes").Value = hsConsumidasMes
        tProjHorasConsumidas.Fields("hsConsumidasTotal").Value = hsConsumidasTotal
        tProjHorasConsumidas.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjHorasConsumidas - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CalcHorasTotal(vEmpr As Short, vDepto As Short, vProj As Double, vAtiv As Short, vSAtiv As Short, vPeriodo As Date) As Double

        Dim vSelec As String
        Dim RC As Integer

        On Error GoTo ECalcHorasTotal

        'Calcula o total de horas do projeto:
        vSelec = "SELECT * FROM ProjHorasConsumidas WHERE idEmpresa = " & vEmpr
        vSelec = vSelec & " AND idDepto = " & vDepto
        vSelec = vSelec & " AND idProjeto = " & vProj
        vSelec = vSelec & " AND idAtividade = " & vAtiv
        vSelec = vSelec & " AND idSAtividade = " & vSAtiv
        vSelec = vSelec & " AND Convert(datetime, periodo, 112) < '" & String.Format(vPeriodo, "yyyymmdd") & "'"
        RC = dbConecta(0, 1, vSelec)
        RC = leSeq(1)
        Do While RC = 0
            CalcHorasTotal = CalcHorasTotal + hsConsumidasMes
            'Lê o próximo registro:
            RC = leSeq(0)
        Loop
        If RC <> 1016 Then
            vSelec = "Erro no Cálculo de Total de Horas consumida pelo projeto " & vProj & Chr(10)
            MsgBox(vSelec & "Erro " & RC)
            Exit Function
        End If

ECalcHorasTotal:
        If Err.Number Then
            vSelec = "Erro no Cálculo de Total de Horas consumida pelo projeto " & vProj & Chr(10)
            MsgBox(vSelec & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
