Option Explicit On
Imports ADODB

Public Class ProjPropostas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850720"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320720"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0720"
#End Region

#Region "variaveis de ambiente"

    Private mvaridProposta As Single
    Public Property idProposta() As Single
        Get
            Return mvaridProposta
        End Get
        Set(value As Single)
            mvaridProposta = value
        End Set
    End Property

    Private mvarversao As Short
    Public Property versao() As Short
        Get
            Return mvarversao
        End Get
        Set(value As Short)
            mvarversao = value
        End Set
    End Property

    Private mvarano As Short
    Public Property ano() As Short
        Get
            Return mvarano
        End Get
        Set(value As Short)
            mvarano = value
        End Set
    End Property

    Private mvartipo As String
    Public Property tipo() As String
        Get
            Return mvartipo
        End Get
        Set(value As String)
            mvartipo = value
        End Set
    End Property

    Private mvaridPropostaExt As String
    Public Property idPropostaExt() As String
        Get
            Return mvaridPropostaExt
        End Get
        Set(value As String)
            mvaridPropostaExt = value
        End Set
    End Property

    Private mvarnomeProposta As String
    Public Property nomeProposta() As String
        Get
            Return mvarnomeProposta
        End Get
        Set(value As String)
            mvarnomeProposta = value
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

    Private mvardtFim As Date
    Public Property dtFim() As Date
        Get
            Return mvardtFim
        End Get
        Set(value As Date)
            mvardtFim = value
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

    Private mvaridCliente As Short
    Public Property idCliente() As Short
        Get
            Return mvaridCliente
        End Get
        Set(value As Short)
            mvaridCliente = value
        End Set
    End Property

    Private mvaridColaborador As Single
    Public Property idColaborador() As Single
        Get
            Return mvaridColaborador
        End Get
        Set(value As Single)
            mvaridColaborador = value
        End Set
    End Property

    Private mvaridServProp As Short
    Public Property idServProp() As Short
        Get
            Return mvaridServProp
        End Get
        Set(value As Short)
            mvaridServProp = value
        End Set
    End Property

    Private mvaridMotivosProp As Short
    Public Property idMotivosProp() As Short
        Get
            Return mvaridMotivosProp
        End Get
        Set(value As Short)
            mvaridMotivosProp = value
        End Set
    End Property

    Private mvaridModeloContrat As Short
    Public Property idModeloContrat() As Short
        Get
            Return mvaridModeloContrat
        End Get
        Set(value As Short)
            mvaridModeloContrat = value
        End Set
    End Property

    Private mvaridResultado As Short
    Public Property idResultado() As Short
        Get
            Return mvaridResultado
        End Get
        Set(value As Short)
            mvaridResultado = value
        End Set
    End Property

    Private mvaridAcoesProp As Short
    Public Property idAcoesProp() As Short
        Get
            Return mvaridAcoesProp
        End Get
        Set(value As Short)
            mvaridAcoesProp = value
        End Set
    End Property

    Private mvarqtdAcoes As Short
    Public Property qtdAcoes() As Short
        Get
            Return mvarqtdAcoes
        End Get
        Set(value As Short)
            mvarqtdAcoes = value
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

    Private mvarduracaoProp As Short
    Public Property duracaoProp() As Short
        Get
            Return mvarduracaoProp
        End Get
        Set(value As Short)
            mvarduracaoProp = value
        End Set
    End Property

    Private mvartpDuracaoProp As Short
    Public Property tpDuracaoProp() As Short
        Get
            Return mvartpDuracaoProp
        End Get
        Set(value As Short)
            mvartpDuracaoProp = value
        End Set
    End Property

    Private mvarvalor1 As Double
    Public Property valor1() As Double
        Get
            Return mvarvalor1
        End Get
        Set(value As Double)
            mvarvalor1 = value
        End Set
    End Property

    Private mvarqtde1 As Short
    Public Property qtde1() As Short
        Get
            Return mvarqtde1
        End Get
        Set(value As Short)
            mvarqtde1 = value
        End Set
    End Property

    Private mvaridTpValor1 As Short
    Public Property idTpValor1() As Short
        Get
            Return mvaridTpValor1
        End Get
        Set(value As Short)
            mvaridTpValor1 = value
        End Set
    End Property

    Private mvarvalor2 As Double
    Public Property valor2() As Double
        Get
            Return mvarvalor2
        End Get
        Set(value As Double)
            mvarvalor2 = value
        End Set
    End Property

    Private mvarqtde2 As Short
    Public Property qtde2() As Short
        Get
            Return mvarqtde2
        End Get
        Set(value As Short)
            mvarqtde2 = value
        End Set
    End Property

    Private mvaridTpValor2 As Short
    Public Property idTpValor2() As Short
        Get
            Return mvaridTpValor2
        End Get
        Set(value As Short)
            mvaridTpValor2 = value
        End Set
    End Property

    Private mvarvalor3 As Double
    Public Property valor3() As Double
        Get
            Return mvarvalor3
        End Get
        Set(value As Double)
            mvarvalor3 = value
        End Set
    End Property

    Private mvarqtde3 As Short
    Public Property qtde3() As Short
        Get
            Return mvarqtde3
        End Get
        Set(value As Short)
            mvarqtde3 = value
        End Set
    End Property

    Private mvaridTpValor3 As Short
    Public Property idTpValor3() As Short
        Get
            Return mvaridTpValor3
        End Get
        Set(value As Short)
            mvaridTpValor3 = value
        End Set
    End Property

    Private mvaridMoeda As Short
    Public Property idMoeda() As Short
        Get
            Return mvaridMoeda
        End Get
        Set(value As Short)
            mvaridMoeda = value
        End Set
    End Property
#End Region

#Region "conexão com banco"

    Public Db As ADODB.Connection
    Public tProjPropostas As ADODB.Recordset
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
        tProjPropostas = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjPropostas.Open("ProjPropostas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjPropostas ORDER BY nomeProposta"
                tProjPropostas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjPropostas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjPropostas.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjPropostas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjPropostas.Open("ProjPropostas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjPropostas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function numReg(vIdProposta As Single) As Short
        '* *************************************************
        '* Retorna a quantidade de Versões de uma Proposta *
        '* *************************************************

        Dim tProjPropostas As New Recordset

        tProjPropostas.Open("SELECT idProposta, versao, ano FROM ProjPropostas WHERE idProposta = " & vIdProposta, Db, ADODB.CursorTypeEnum.adOpenDynamic)
        Do While Not tProjPropostas.EOF
            numReg = numReg + 1
            tProjPropostas.MoveNext()
        Loop

    End Function

    Public Function numVersao(vIdProposta As Single) As Short
        '* *****************************************
        '* Retorna a última Versão de uma Proposta *
        '* *****************************************

        Dim tProjPropostas As New Recordset
        Dim vTipo As String
        Dim ultNumero As Short

        tProjPropostas.Open("SELECT idProposta, versao, tipo FROM ProjPropostas WHERE idProposta = " & vIdProposta & " ORDER BY tipo, versao", Db, ADODB.CursorTypeEnum.adOpenDynamic)
        If Not tProjPropostas.EOF Then
            vTipo = tProjPropostas.Fields("tipo").Value
        End If
        Do While Not tProjPropostas.EOF
            If vTipo <> tProjPropostas.Fields("tipo").Value Then
                ultNumero = numVersao
            End If
            numVersao = tProjPropostas.Fields("versao").Value
            tProjPropostas.MoveNext()
        Loop
        If ultNumero > numVersao Then
            numVersao = ultNumero
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
        If Not tProjPropostas.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjPropostas.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjPropostas.EOF Then
            idProposta = tProjPropostas.Fields("idProposta").Value
            versao = tProjPropostas.Fields("versao").Value
            ano = tProjPropostas.Fields("ano").Value
            tipo = tProjPropostas.Fields("tipo").Value
            idPropostaExt = tProjPropostas.Fields("idPropostaExt").Value
            nomeProposta = tProjPropostas.Fields("nomeProposta").Value
            dtInicio = tProjPropostas.Fields("dtInicio").Value
            dtFim = tProjPropostas.Fields("dtFim").Value
            idProjeto = tProjPropostas.Fields("idProjeto").Value
            idCliente = tProjPropostas.Fields("idCliente").Value
            idColaborador = tProjPropostas.Fields("idColaborador").Value
            idServProp = tProjPropostas.Fields("idServProp").Value
            idMotivosProp = tProjPropostas.Fields("idMotivosProp").Value
            idModeloContrat = tProjPropostas.Fields("idModeloContrat").Value
            idResultado = tProjPropostas.Fields("idResultado").Value
            idAcoesProp = tProjPropostas.Fields("idAcoesProp").Value
            qtdAcoes = tProjPropostas.Fields("qtdAcoes").Value
            idApuracaoHoras = tProjPropostas.Fields("idApuracaoHoras").Value
            duracaoProp = tProjPropostas.Fields("duracaoProp").Value
            tpDuracaoProp = tProjPropostas.Fields("tpDuracaoProp").Value
            valor1 = tProjPropostas.Fields("valor1").Value
            qtde1 = tProjPropostas.Fields("qtde1").Value
            idTpValor1 = tProjPropostas.Fields("idTpValor1").Value
            valor2 = tProjPropostas.Fields("valor2").Value
            qtde2 = tProjPropostas.Fields("qtde2").Value
            idTpValor2 = tProjPropostas.Fields("idTpValor2").Value
            valor3 = tProjPropostas.Fields("valor3").Value
            qtde3 = tProjPropostas.Fields("qtde3").Value
            idTpValor3 = tProjPropostas.Fields("idTpValor3").Value
            idMoeda = tProjPropostas.Fields("idMoeda").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjPropostas - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(proposta As Single, versao As Short, tipo As Short, propostaExt As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjPropostas.                       *
        '* *                                      *
        '* * proposta    = identif. para pesquisa *
        '* * versao      = identif. para pesquisa *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjPropostas WHERE idProposta = " & proposta
        vSelect = vSelect & " AND versao = " & versao
        vSelect = vSelect & " AND tipo = '" & tipo & "'"
        vSelect = vSelect & " AND idPropostaExt = '" & propostaExt & "'"
        tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0        'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjPropostas.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProposta = tProjPropostas.Fields("idProposta").Value
            versao = tProjPropostas.Fields("versao").Value
            ano = tProjPropostas.Fields("ano").Value
            tipo = tProjPropostas.Fields("tipo").Value
            idPropostaExt = tProjPropostas.Fields("idPropostaExt").Value
            nomeProposta = tProjPropostas.Fields("nomeProposta").Value
            dtInicio = tProjPropostas.Fields("dtInicio").Value
            dtFim = tProjPropostas.Fields("dtFim").Value
            idProjeto = tProjPropostas.Fields("idProjeto").Value
            idCliente = tProjPropostas.Fields("idCliente").Value
            idColaborador = tProjPropostas.Fields("idColaborador").Value
            idServProp = tProjPropostas.Fields("idServProp").Value
            idMotivosProp = tProjPropostas.Fields("idMotivosProp").Value
            idModeloContrat = tProjPropostas.Fields("idModeloContrat").Value
            idResultado = tProjPropostas.Fields("idResultado").Value
            idAcoesProp = tProjPropostas.Fields("idAcoesProp").Value
            qtdAcoes = tProjPropostas.Fields("qtdAcoes").Value
            idApuracaoHoras = tProjPropostas.Fields("idApuracaoHoras").Value
            duracaoProp = tProjPropostas.Fields("duracaoProp").Value
            tpDuracaoProp = tProjPropostas.Fields("tpDuracaoProp").Value
            valor1 = tProjPropostas.Fields("valor1").Value
            qtde1 = tProjPropostas.Fields("qtde1").Value
            idTpValor1 = tProjPropostas.Fields("idTpValor1").Value
            valor2 = tProjPropostas.Fields("valor2").Value
            qtde2 = tProjPropostas.Fields("qtde2").Value
            idTpValor2 = tProjPropostas.Fields("idTpValor2").Value
            valor3 = tProjPropostas.Fields("valor3").Value
            qtde3 = tProjPropostas.Fields("qtde3").Value
            idTpValor3 = tProjPropostas.Fields("idTpValor3").Value
            idMoeda = tProjPropostas.Fields("idMoeda").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjPropostas.Close()
                tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjPropostas - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function
    Public Function localizaExt(propostaExt As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjPropostas.                       *
        '* *                                      *
        '* * propostaExt = identif. para pesquisa *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaExt
        vSelect = "SELECT * FROM ProjPropostas WHERE idProposta = " & Mid(propostaExt, 3, 4)
        vSelect = vSelect & " AND versao = " & Mid(propostaExt, 8, 2)
        vSelect = vSelect & " AND tipo = '" & Mid(propostaExt, 1, 2) & "'"
        tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaExt = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjPropostas.EOF Then
            localizaExt = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProposta = tProjPropostas.Fields("idProposta").Value
            versao = tProjPropostas.Fields("versao").Value
            ano = tProjPropostas.Fields("ano").Value
            tipo = tProjPropostas.Fields("tipo").Value
            idPropostaExt = tProjPropostas.Fields("idPropostaExt").Value
            nomeProposta = tProjPropostas.Fields("nomeProposta").Value
            dtInicio = tProjPropostas.Fields("dtInicio").Value
            dtFim = tProjPropostas.Fields("dtFim").Value
            idProjeto = tProjPropostas.Fields("idProjeto").Value
            idCliente = tProjPropostas.Fields("idCliente").Value
            idColaborador = tProjPropostas.Fields("idColaborador").Value
            idServProp = tProjPropostas.Fields("idServProp").Value
            idMotivosProp = tProjPropostas.Fields("idMotivosProp").Value
            idModeloContrat = tProjPropostas.Fields("idModeloContrat").Value
            idResultado = tProjPropostas.Fields("idResultado").Value
            idAcoesProp = tProjPropostas.Fields("idAcoesProp").Value
            qtdAcoes = tProjPropostas.Fields("qtdAcoes").Value
            idApuracaoHoras = tProjPropostas.Fields("idApuracaoHoras").Value
            duracaoProp = tProjPropostas.Fields("duracaoProp").Value
            tpDuracaoProp = tProjPropostas.Fields("tpDuracaoProp").Value
            valor1 = tProjPropostas.Fields("valor1").Value
            qtde1 = tProjPropostas.Fields("qtde1").Value
            idTpValor1 = tProjPropostas.Fields("idTpValor1").Value
            valor2 = tProjPropostas.Fields("valor2").Value
            qtde2 = tProjPropostas.Fields("qtde2").Value
            idTpValor2 = tProjPropostas.Fields("idTpValor2").Value
            valor3 = tProjPropostas.Fields("valor3").Value
            qtde3 = tProjPropostas.Fields("qtde3").Value
            idTpValor3 = tProjPropostas.Fields("idTpValor3").Value
            idMoeda = tProjPropostas.Fields("idMoeda").Value
        End If

ElocalizaExt:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjPropostas.Close()
                tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaExt = Err.Number
                MsgBox("Classe ProjPropostas - localizaExt" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* *******************************************
        '* * Localiza um registro específico,        *
        '* * baseado na Chave da tabela              *
        '* * ProjPropostas.                          *
        '* *                                         *
        '* * descric    = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe    *
        '* *******************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjPropostas WHERE nomeProposta = '" & descric & "'"
        tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0        'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjPropostas.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProposta = tProjPropostas.Fields("idProposta").Value
            versao = tProjPropostas.Fields("versao").Value
            ano = tProjPropostas.Fields("ano").Value
            tipo = tProjPropostas.Fields("tipo").Value
            idPropostaExt = tProjPropostas.Fields("idPropostaExt").Value
            nomeProposta = tProjPropostas.Fields("nomeProposta").Value
            dtInicio = tProjPropostas.Fields("dtInicio").Value
            dtFim = tProjPropostas.Fields("dtFim").Value
            idProjeto = tProjPropostas.Fields("idProjeto").Value
            idCliente = tProjPropostas.Fields("idCliente").Value
            idColaborador = tProjPropostas.Fields("idColaborador").Value
            idServProp = tProjPropostas.Fields("idServProp").Value
            idMotivosProp = tProjPropostas.Fields("idMotivosProp").Value
            idModeloContrat = tProjPropostas.Fields("idModeloContrat").Value
            idResultado = tProjPropostas.Fields("idResultado").Value
            idAcoesProp = tProjPropostas.Fields("idAcoesProp").Value
            qtdAcoes = tProjPropostas.Fields("qtdAcoes").Value
            idApuracaoHoras = tProjPropostas.Fields("idApuracaoHoras").Value
            duracaoProp = tProjPropostas.Fields("duracaoProp").Value
            tpDuracaoProp = tProjPropostas.Fields("tpDuracaoProp").Value
            valor1 = tProjPropostas.Fields("valor1").Value
            qtde1 = tProjPropostas.Fields("qtde1").Value
            idTpValor1 = tProjPropostas.Fields("idTpValor1").Value
            valor2 = tProjPropostas.Fields("valor2").Value
            qtde2 = tProjPropostas.Fields("qtde2").Value
            idTpValor2 = tProjPropostas.Fields("idTpValor2").Value
            valor3 = tProjPropostas.Fields("valor3").Value
            qtde3 = tProjPropostas.Fields("qtde3").Value
            idTpValor3 = tProjPropostas.Fields("idTpValor3").Value
            idMoeda = tProjPropostas.Fields("idMoeda").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjPropostas.Close()
                tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjPropostas - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNomeExt(descricExt As String, atualiz As Short) As Integer
        '* *******************************************
        '* * Localiza um registro específico,        *
        '* * baseado na Chave da tabela              *
        '* * ProjPropostas.                          *
        '* *                                         *
        '* * descricExt = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe    *
        '* *******************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNomeExt
        vSelect = "SELECT * FROM ProjPropostas WHERE nomeProposta = '" & descricExt & "'"
        tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNomeExt = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjPropostas.EOF Then
            localizaNomeExt = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProposta = tProjPropostas.Fields("idProposta").Value
            versao = tProjPropostas.Fields("versao").Value
            ano = tProjPropostas.Fields("ano").Value
            tipo = tProjPropostas.Fields("tipo").Value
            idPropostaExt = tProjPropostas.Fields("idPropostaExt").Value
            nomeProposta = tProjPropostas.Fields("nomeProposta").Value
            dtInicio = tProjPropostas.Fields("dtInicio").Value
            dtFim = tProjPropostas.Fields("dtFim").Value
            idProjeto = tProjPropostas.Fields("idProjeto").Value
            idCliente = tProjPropostas.Fields("idCliente").Value
            idColaborador = tProjPropostas.Fields("idColaborador").Value
            idServProp = tProjPropostas.Fields("idServProp").Value
            idMotivosProp = tProjPropostas.Fields("idMotivosProp").Value
            idModeloContrat = tProjPropostas.Fields("idModeloContrat").Value
            idResultado = tProjPropostas.Fields("idResultado").Value
            idAcoesProp = tProjPropostas.Fields("idAcoesProp").Value
            qtdAcoes = tProjPropostas.Fields("qtdAcoes").Value
            idApuracaoHoras = tProjPropostas.Fields("idApuracaoHoras").Value
            duracaoProp = tProjPropostas.Fields("duracaoProp").Value
            tpDuracaoProp = tProjPropostas.Fields("tpDuracaoProp").Value
            valor1 = tProjPropostas.Fields("valor1").Value
            qtde1 = tProjPropostas.Fields("qtde1").Value
            idTpValor1 = tProjPropostas.Fields("idTpValor1").Value
            valor2 = tProjPropostas.Fields("valor2").Value
            qtde2 = tProjPropostas.Fields("qtde2").Value
            idTpValor2 = tProjPropostas.Fields("idTpValor2").Value
            valor3 = tProjPropostas.Fields("valor3").Value
            qtde3 = tProjPropostas.Fields("qtde3").Value
            idTpValor3 = tProjPropostas.Fields("idTpValor3").Value
            idMoeda = tProjPropostas.Fields("idMoeda").Value
        End If

ElocalizaNomeExt:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjPropostas.Close()
                tProjPropostas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNomeExt = Err.Number
                MsgBox("Classe ProjPropostas - localizaNomeExt" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(descric As String) As Integer

        On Error GoTo Einclui
        tProjPropostas.AddNew()
        tProjPropostas.Fields("idProposta").Value = idProposta
        tProjPropostas.Fields("versao").Value = versao
        tProjPropostas.Fields("ano").Value = ano
        tProjPropostas.Fields("tipo").Value = tipo
        tProjPropostas.Fields("idPropostaExt").Value = idPropostaExt
        tProjPropostas.Fields("nomeProposta").Value = nomeProposta
        tProjPropostas.Fields("dtInicio").Value = dtInicio
        tProjPropostas.Fields("dtFim").Value = dtFim
        tProjPropostas.Fields("idProjeto").Value = idProjeto
        'Carrega Descrição da Proposta:
        tProjPropostas.Fields("descricProposta").Value = IIf(descric = "", " ", descric)
        tProjPropostas.Fields("idCliente").Value = idCliente
        tProjPropostas.Fields("idColaborador").Value = idColaborador
        tProjPropostas.Fields("idServProp").Value = idServProp
        tProjPropostas.Fields("idMotivosProp").Value = idMotivosProp
        tProjPropostas.Fields("idModeloContrat").Value = idModeloContrat
        tProjPropostas.Fields("idResultado").Value = idResultado
        tProjPropostas.Fields("idAcoesProp").Value = idAcoesProp
        tProjPropostas.Fields("qtdAcoes").Value = qtdAcoes
        tProjPropostas.Fields("idApuracaoHoras").Value = idApuracaoHoras
        tProjPropostas.Fields("duracaoProp").Value = duracaoProp
        tProjPropostas.Fields("tpDuracaoProp").Value = tpDuracaoProp
        tProjPropostas.Fields("valor1").Value = valor1
        tProjPropostas.Fields("qtde1").Value = qtde1
        tProjPropostas.Fields("idTpValor1").Value = idTpValor1
        tProjPropostas.Fields("valor2").Value = valor2
        tProjPropostas.Fields("qtde2").Value = qtde2
        tProjPropostas.Fields("idTpValor2").Value = idTpValor2
        tProjPropostas.Fields("valor3").Value = valor3
        tProjPropostas.Fields("qtde3").Value = qtde3
        tProjPropostas.Fields("idTpValor3").Value = idTpValor3
        tProjPropostas.Fields("idMoeda").Value = idMoeda
        tProjPropostas.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjPropostas - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(descric As String) As Integer

        On Error GoTo Ealtera
        tProjPropostas.Fields("idProposta").Value = idProposta
        tProjPropostas.Fields("versao").Value = versao
        tProjPropostas.Fields("ano").Value = ano
        tProjPropostas.Fields("tipo").Value = tipo
        tProjPropostas.Fields("idPropostaExt").Value = idPropostaExt
        tProjPropostas.Fields("nomeProposta").Value = nomeProposta
        tProjPropostas.Fields("dtInicio").Value = dtInicio
        tProjPropostas.Fields("dtFim").Value = dtFim
        tProjPropostas.Fields("idProjeto").Value = idProjeto
        'Carrega Descrição da Proposta:
        tProjPropostas.Fields("descricProposta").Value = IIf(descric = "", " ", descric)
        tProjPropostas.Fields("idCliente").Value = idCliente
        tProjPropostas.Fields("idColaborador").Value = idColaborador
        tProjPropostas.Fields("idServProp").Value = idServProp
        tProjPropostas.Fields("idMotivosProp").Value = idMotivosProp
        tProjPropostas.Fields("idModeloContrat").Value = idModeloContrat
        tProjPropostas.Fields("idResultado").Value = idResultado
        tProjPropostas.Fields("idAcoesProp").Value = idAcoesProp
        tProjPropostas.Fields("qtdAcoes").Value = qtdAcoes
        tProjPropostas.Fields("idApuracaoHoras").Value = idApuracaoHoras
        tProjPropostas.Fields("duracaoProp").Value = duracaoProp
        tProjPropostas.Fields("tpDuracaoProp").Value = tpDuracaoProp
        tProjPropostas.Fields("valor1").Value = valor1
        tProjPropostas.Fields("qtde1").Value = qtde1
        tProjPropostas.Fields("idTpValor1").Value = idTpValor1
        tProjPropostas.Fields("valor2").Value = valor2
        tProjPropostas.Fields("qtde2").Value = qtde2
        tProjPropostas.Fields("idTpValor2").Value = idTpValor2
        tProjPropostas.Fields("valor3").Value = valor3
        tProjPropostas.Fields("qtde3").Value = qtde3
        tProjPropostas.Fields("idTpValor3").Value = idTpValor3
        tProjPropostas.Fields("idMoeda").Value = idMoeda
        tProjPropostas.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjPropostas - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function descricProposta() As String

        descricProposta = tProjPropostas.Fields("descricProposta").Value

    End Function

    Public Function elimina(proposta As Single, versao As Short, tipo As Short, ano As Short, propostaExt As String) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjPropostas.                       *
        '* *                                      *
        '* * proposta    = identif. para pesquisa *
        '* * versao      = identif. para pesquisa *
        '* * tipo        = identif. para pesquisa *
        '* * ano         = identif. para pesquisa *
        '* * propostaExt = identif. para pesquisa *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim vProposta As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de segMensagens")
            elimina = -1
            Exit Function
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Monta a proposta no formato de exibição:
        vProposta = tipo & String.Format(proposta, "0###") & "-" & String.Format(versao, "0#") & "/" & ano

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em ProjAlocHoras:
        vSelec = "SELECT * FROM ProjAlocHoras WHERE idProposta = '" & vProposta & "'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há Proposta vinculada à alocação de horas de projeto." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM ProjPropostas WHERE "
        vSelec = vSelec & " idProposta = " & proposta
        vSelec = vSelec & " AND versao = " & versao
        vSelec = vSelec & " AND tipo = '" & tipo & "'"
        vSelec = vSelec & " AND idPropostaExt = '" & propostaExt & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjPropostas - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function eliminaExt(propostaExt As String) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjPropostas.                       *
        '* *                                      *
        '* * propostaExt = identif. para pesquisa *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de segMensagens")
            eliminaExt = -1
            Exit Function
        End If

        On Error GoTo EeliminaExt
        eliminaExt = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em ProjAlocHoras:
        vSelec = "SELECT * FROM ProjAlocHoras WHERE idPropostaExt = '" & propostaExt & "'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há Proposta vinculada à alocação de horas de projeto." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            eliminaExt = RC
            Exit Function
        End If

        vSelec = "DELETE FROM ProjPropostas WHERE "
        vSelec = vSelec & " idPropostaExt = '" & propostaExt & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

EeliminaExt:
        If Err.Number Then
            eliminaExt = Err.Number
            MsgBox("Classe ProjPropostas - eliminaExt" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
