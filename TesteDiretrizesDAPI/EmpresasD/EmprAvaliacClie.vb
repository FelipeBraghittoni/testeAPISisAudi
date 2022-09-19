Option Explicit On
Public Class EmprAvaliacClie
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850304"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320304"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0304"
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

    Private mvaridCliente As Short
    Public Property idCliente() As Short
        Get
            Return mvaridCliente
        End Get
        Set(value As Short)
            mvaridCliente = value
        End Set
    End Property


    Private mvarnomeAvaliador As String
    Public Property nomeAvaliador() As String
        Get
            Return mvarnomeAvaliador
        End Get
        Set(value As String)
            mvarnomeAvaliador = value
        End Set
    End Property


    Private mvarcargoAvaliador As String
    Public Property cargoAvaliador() As String
        Get
            Return mvarcargoAvaliador
        End Get
        Set(value As String)
            mvarcargoAvaliador = value
        End Set
    End Property


    Private mvardtAvaliacao As Date
    Public Property dtAvaliacao() As Date
        Get
            Return mvardtAvaliacao
        End Get
        Set(value As Date)
            mvardtAvaliacao = value
        End Set
    End Property


#End Region

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tAvaliacClie As ADODB.Recordset
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
        tAvaliacClie = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tAvaliacClie.Open("EmprAvaliacClie", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tAvaliacClie.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tAvaliacClie.Close()

                If tipo = 1 Then
                    tAvaliacClie.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tAvaliacClie.Open("EmprAvaliacClie", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprAvaliacClie - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tAvaliacClie.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tAvaliacClie.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tAvaliacClie.EOF Then
            idEmpresa = tAvaliacClie.Fields("idEmpresa").Value
            idCliente = tAvaliacClie.Fields("idCliente").Value
            nomeAvaliador = tAvaliacClie.Fields("nomeAvaliador").Value
            cargoAvaliador = tAvaliacClie.Fields("cargoAvaliador").Value
            dtAvaliacao = tAvaliacClie.Fields("dtAvaliacao").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprAvaliacClie - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, cliente As Short, data As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * AvaliacClie.                         *
        '* *                                      *
        '* * empresa   = identif. para pesquisa + *
        '* * cliente   = identif. para pesquisa + *
        '* * data      = identif. para pesquisa   *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprAvaliacClie WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idCliente = " & cliente
        vSelect = vSelect & " AND Convert(datetime, dtAvaliacao, 112) = '" & Format(data, "yyyymmdd") & "'"
        tAvaliacClie.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tAvaliacClie.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tAvaliacClie.Fields("idEmpresa").Value
            idCliente = tAvaliacClie.Fields("idCliente").Value
            nomeAvaliador = tAvaliacClie.Fields("nomeAvaliador").Value
            cargoAvaliador = tAvaliacClie.Fields("cargoAvaliador").Value
            dtAvaliacao = tAvaliacClie.Fields("dtAvaliacao").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tAvaliacClie.Close()
                tAvaliacClie.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprAvaliacClie - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(avaliacao As String) As Integer

        On Error GoTo Einclui
        tAvaliacClie.AddNew()
        tAvaliacClie(0).Value = idEmpresa
        tAvaliacClie(1).Value = idCliente
        tAvaliacClie(2).Value = nomeAvaliador
        tAvaliacClie(3).Value = cargoAvaliador
        tAvaliacClie(4).Value = dtAvaliacao
        tAvaliacClie(5).Value = IIf(avaliacao = "", " ", avaliacao)
        tAvaliacClie.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprAvaliacClie - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(avaliacao As String) As Integer

        On Error GoTo Ealtera

        tAvaliacClie(2).Value = nomeAvaliador
        tAvaliacClie(3).Value = cargoAvaliador
        tAvaliacClie(5).Value = IIf(avaliacao = "", " ", avaliacao)
        tAvaliacClie.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprAvaliacClie - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String

        carregaMemos = tAvaliacClie.Fields("txtAvaliacao").Value

    End Function

    Public Function elimina(empresa As Short, clie As Short, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprAvaliacClie.                     *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * clie       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprAvaliacClie WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idCliente = " & clie
        vSelec = vSelec & " AND Convert(datetime, dtAvaliacao, 112) = '" & Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprAvaliacClie - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
