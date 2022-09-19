Option Explicit On
Imports System.Windows.Forms

Public Class ProjProjetos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850719"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320719"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0719"
#End Region


#Region "Variaveis de ambiente"
    Dim ApontHoras As ApontHoras

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

    Private mvarlocalContrProj As Short
    Public Property localContrProj() As Short
        Get
            Return mvarlocalContrProj
        End Get
        Set(value As Short)
            mvarlocalContrProj = value
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

    Private mvaridColabComerc As Single
    Public Property idColabComerc() As Single
        Get
            Return mvaridColabComerc
        End Get
        Set(value As Single)
            mvaridColabComerc = value
        End Set
    End Property

    Private mvaridColabTecnico As Single
    Public Property idColabTecnico() As Single
        Get
            Return mvaridColabTecnico
        End Get
        Set(value As Single)
            mvaridColabTecnico = value
        End Set
    End Property

    Private mvarnaturProjeto As Short
    Public Property naturProjeto() As Short
        Get
            Return mvarnaturProjeto
        End Get
        Set(value As Short)
            mvarnaturProjeto = value
        End Set
    End Property

    Private mvaridColabProjeto As Single
    Public Property idColabProjeto() As Single
        Get
            Return mvaridColabProjeto
        End Get
        Set(value As Single)
            mvaridColabProjeto = value
        End Set
    End Property

    Private mvardtInicProj As Date
    Public Property dtInicProj() As Date
        Get
            Return mvardtInicProj
        End Get
        Set(value As Date)
            mvardtInicProj = value
        End Set
    End Property

    Private mvardtFinalProj As Date
    Public Property dtFinalProj() As Date
        Get
            Return mvardtFinalProj
        End Get
        Set(value As Date)
            mvardtFinalProj = value
        End Set
    End Property

    Private mvaridModalidade As Short
    Public Property idModalidade() As Short
        Get
            Return mvaridModalidade
        End Get
        Set(value As Short)
            mvaridModalidade = value
        End Set
    End Property

    Private mvardiaInicFat As Short
    Public Property diaInicFat() As Short
        Get
            Return mvardiaInicFat
        End Get
        Set(value As Short)
            mvardiaInicFat = value
        End Set
    End Property

    Private mvartpDiaInicFat As Short
    Public Property tpDiaInicFat() As Short
        Get
            Return mvartpDiaInicFat
        End Get
        Set(value As Short)
            mvartpDiaInicFat = value
        End Set
    End Property

    Private mvardiaFimFat As Short
    Public Property diaFimFat() As Short
        Get
            Return mvardiaFimFat
        End Get
        Set(value As Short)
            mvardiaFimFat = value
        End Set
    End Property

    Private mvartpDiaFimFat As Short
    Public Property tpDiaFimFat() As Short
        Get
            Return mvartpDiaFimFat
        End Get
        Set(value As Short)
            mvartpDiaFimFat = value
        End Set
    End Property

    Private mvareMailResp As String
    Public Property eMailResp() As String
        Get
            Return mvareMailResp
        End Get
        Set(value As String)
            mvareMailResp = value
        End Set
    End Property

    Private mvaridAgrupaVinculo As Short
    Public Property idAgrupaVinculo() As Short
        Get
            Return mvaridAgrupaVinculo
        End Get
        Set(value As Short)
            mvaridAgrupaVinculo = value
        End Set
    End Property

    Private mvaridAgrupProjeto As Double
    Public Property idAgrupProjeto() As Double
        Get
            Return mvaridAgrupProjeto
        End Get
        Set(value As Double)
            mvaridAgrupProjeto = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tProjProjetos As ADODB.Recordset
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
        tProjProjetos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjProjetos.Open("ProjProjetos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjProjetos ORDER BY nomeProjeto"
                tProjProjetos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjProjetos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjProjetos.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjProjetos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjProjetos.Open("ProjProjetos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjProjetos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjProjetos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjProjetos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjProjetos.EOF Then
            idProjeto = tProjProjetos.Fields("idProjeto").Value
            nomeProjeto = tProjProjetos.Fields("nomeProjeto").Value
            localContrProj = tProjProjetos.Fields("localContrProj").Value
            idOrganiz = tProjProjetos.Fields("idOrganiz").Value
            idSetorOrg = tProjProjetos.Fields("idSetorOrg").Value
            idEmpresa = tProjProjetos.Fields("idEmpresa").Value
            idDepto = tProjProjetos.Fields("idDepto").Value
            naturProjeto = tProjProjetos.Fields("naturProjeto").Value
            dtInicProj = tProjProjetos.Fields("dtInicProj").Value
            dtFinalProj = tProjProjetos.Fields("dtFinalProj").Value
            idColabComerc = tProjProjetos.Fields("idColabComerc").Value
            idColabTecnico = tProjProjetos.Fields("idColabTecnico").Value
            idColabProjeto = tProjProjetos.Fields("idColabProjeto").Value
            idModalidade = tProjProjetos.Fields("idModalidade").Value
            diaInicFat = tProjProjetos.Fields("diaInicFat").Value
            tpDiaInicFat = tProjProjetos.Fields("tpDiaInicFat").Value
            diaFimFat = tProjProjetos.Fields("diaFimFat").Value
            tpDiaFimFat = tProjProjetos.Fields("tpDiaFimFat").Value
            eMailResp = tProjProjetos.Fields("eMailResp").Value
            idAgrupaVinculo = tProjProjetos.Fields("idAgrupaVinculo").Value
            idAgrupProjeto = tProjProjetos.Fields("idAgrupProjeto").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjProjetos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(projeto As Double, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjProjetos.                        *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjProjetos WHERE idProjeto = " & projeto
        tProjProjetos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjProjetos.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjProjetos.Fields("idProjeto").Value
            nomeProjeto = tProjProjetos.Fields("nomeProjeto").Value
            localContrProj = tProjProjetos.Fields("localContrProj").Value
            idOrganiz = tProjProjetos.Fields("idOrganiz").Value
            idSetorOrg = tProjProjetos.Fields("idSetorOrg").Value
            idEmpresa = tProjProjetos.Fields("idEmpresa").Value
            idDepto = tProjProjetos.Fields("idDepto").Value
            naturProjeto = tProjProjetos.Fields("naturProjeto").Value
            dtInicProj = tProjProjetos.Fields("dtInicProj").Value
            dtFinalProj = tProjProjetos.Fields("dtFinalProj").Value
            idColabComerc = tProjProjetos.Fields("idColabComerc").Value
            idColabTecnico = tProjProjetos.Fields("idColabTecnico").Value
            idColabProjeto = tProjProjetos.Fields("idColabProjeto").Value
            idModalidade = tProjProjetos.Fields("idModalidade").Value
            diaInicFat = tProjProjetos.Fields("diaInicFat").Value
            tpDiaInicFat = tProjProjetos.Fields("tpDiaInicFat").Value
            diaFimFat = tProjProjetos.Fields("diaFimFat").Value
            tpDiaFimFat = tProjProjetos.Fields("tpDiaFimFat").Value
            eMailResp = tProjProjetos.Fields("eMailResp").Value
            idAgrupaVinculo = tProjProjetos.Fields("idAgrupaVinculo").Value
            idAgrupProjeto = tProjProjetos.Fields("idAgrupProjeto").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjProjetos.Close()
                tProjProjetos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjProjetos - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * GerEmpresas.                         *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjProjetos WHERE nomeProjeto = '" & descric & "'"
        tProjProjetos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjProjetos.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjProjetos.Fields("idProjeto").Value
            nomeProjeto = tProjProjetos.Fields("nomeProjeto").Value
            localContrProj = tProjProjetos.Fields("localContrProj").Value
            idOrganiz = tProjProjetos.Fields("idOrganiz").Value
            idSetorOrg = tProjProjetos.Fields("idSetorOrg").Value
            idEmpresa = tProjProjetos.Fields("idEmpresa").Value
            idDepto = tProjProjetos.Fields("idDepto").Value
            naturProjeto = tProjProjetos.Fields("naturProjeto").Value
            dtInicProj = tProjProjetos.Fields("dtInicProj").Value
            dtFinalProj = tProjProjetos.Fields("dtFinalProj").Value
            idColabComerc = tProjProjetos.Fields("idColabComerc").Value
            idColabTecnico = tProjProjetos.Fields("idColabTecnico").Value
            idColabProjeto = tProjProjetos.Fields("idColabProjeto").Value
            idModalidade = tProjProjetos.Fields("idModalidade").Value
            diaInicFat = tProjProjetos.Fields("diaInicFat").Value
            tpDiaInicFat = tProjProjetos.Fields("tpDiaInicFat").Value
            diaFimFat = tProjProjetos.Fields("diaFimFat").Value
            tpDiaFimFat = tProjProjetos.Fields("tpDiaFimFat").Value
            eMailResp = tProjProjetos.Fields("eMailResp").Value
            idAgrupaVinculo = tProjProjetos.Fields("idAgrupaVinculo").Value
            idAgrupProjeto = tProjProjetos.Fields("idAgrupProjeto").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjProjetos.Close()
                tProjProjetos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjProjetos - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tProjProjetos.AddNew()
        tProjProjetos(0).Value = idProjeto
        tProjProjetos(1).Value = nomeProjeto
        tProjProjetos(2).Value = localContrProj
        tProjProjetos(3).Value = idOrganiz
        tProjProjetos(4).Value = idSetorOrg
        tProjProjetos(5).Value = idEmpresa
        tProjProjetos(6).Value = idDepto
        tProjProjetos(7).Value = naturProjeto
        tProjProjetos(8).Value = dtInicProj
        tProjProjetos(9).Value = dtFinalProj
        tProjProjetos(10).Value = idColabComerc
        tProjProjetos(11).Value = idColabTecnico
        tProjProjetos(12).Value = idColabProjeto
        tProjProjetos(13).Value = idModalidade
        tProjProjetos.Fields("diaInicFat").Value = diaInicFat
        tProjProjetos.Fields("tpDiaInicFat").Value = tpDiaInicFat
        tProjProjetos.Fields("diaFimFat").Value = diaFimFat
        tProjProjetos.Fields("tpDiaFimFat").Value = tpDiaFimFat
        tProjProjetos.Fields("eMailResp").Value = eMailResp
        tProjProjetos.Fields("idAgrupaVinculo").Value = idAgrupaVinculo
        tProjProjetos.Fields("idAgrupProjeto").Value = idAgrupProjeto
        tProjProjetos.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjProjetos - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tProjProjetos(1).Value = nomeProjeto
        tProjProjetos(2).Value = localContrProj
        tProjProjetos(3).Value = idOrganiz
        tProjProjetos(4).Value = idSetorOrg
        tProjProjetos(5).Value = idEmpresa
        tProjProjetos(6).Value = idDepto
        tProjProjetos(7).Value = naturProjeto
        tProjProjetos(8).Value = dtInicProj
        tProjProjetos(9).Value = dtFinalProj
        tProjProjetos(10).Value = idColabComerc
        tProjProjetos(11).Value = idColabTecnico
        tProjProjetos(12).Value = idColabProjeto
        tProjProjetos(13).Value = idModalidade
        tProjProjetos.Fields("diaInicFat").Value = diaInicFat
        tProjProjetos.Fields("tpDiaInicFat").Value = tpDiaInicFat
        tProjProjetos.Fields("diaFimFat").Value = diaFimFat
        tProjProjetos.Fields("tpDiaFimFat").Value = tpDiaFimFat
        tProjProjetos.Fields("eMailResp").Value = eMailResp
        tProjProjetos.Fields("idAgrupaVinculo").Value = idAgrupaVinculo
        tProjProjetos.Fields("idAgrupProjeto").Value = idAgrupProjeto
        tProjProjetos.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjProjetos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function elimina(projeto As Double) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjProjetos.                        *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens
        Dim GerParametros As GeraisD.GerParametros

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens

        'Define uma instância para analisar o último período contábil fechado:
        GerParametros = New GeraisD.GerParametros

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        RC = VerificaTermino(0, projeto)
        If RC <> 0 Then
            MsgBox("Projeto " & projeto & " não eliminado.")
            Exit Function
        End If

        vSelec = "DELETE FROM ProjProjetos WHERE "
        vSelec = vSelec & " idProjeto = " & projeto
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjProjetos - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function VerificaTermino(vTipo As Short, vProjeto As Double, Optional vDtFim As Date = Nothing) As Integer
        '* ***************************************************
        '* Verifica se há pendências a serem resolvidas      *
        '* antes de efetivar o término/eliminação do Projeto *
        '* ***************************************************
        '* vtipo pode ser:                                   *
        '* 0 - não verifica dtFim; utilizado para eliminar   *
        '* 1 - verifica dtFim; utilizado para encerrar       *
        '* 2 - verifica dtFim - projeto aberto mas inativo   *
        '* ***************************************************
        '* VerificaTermino pode variar de:                   *
        '*  < 0 = ocorreu erro                               *
        '*  = 0 = não encontrou nenhuma alocação pendente    *
        '*  1 a 20 = Alocações pendentes                     *
        '* ***************************************************

        Dim RC As Integer
        Dim vDesaloca(0 To 24) As Boolean
        Dim vSelec As String
        Dim vSelec1 As String
        Dim vSelec2 As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim ParCadastro As Object
        Dim segMensagens As New SegurancaD.segMensagens

        On Error GoTo EVerificaTermino

        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de segMensagens")
            VerificaTermino = -1
            Exit Function
        End If

        'Define instância de ParCadastro
        ParCadastro = New Object 'Dependencia' 'New ParceiriasD.ParCadastro
        RC = ParCadastro.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de ParCadastro")
            VerificaTermino = -1
            Exit Function
        End If

        ApontHoras = New ApontHoras
        RC = ApontHoras.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de ApontHoras (Audihoras)")
            VerificaTermino = -1
            Exit Function
        End If

        VerificaTermino = 0     'Se não houver erro
        For RC = 1 To 21
            vDesaloca(RC) = False
        Next RC
        vSelec1 = "Há alocações em aberto para este Projeto:" & Chr(10)
        vSelec2 = "Há alocações em aberto Impeditivas para este Projeto:" & Chr(10)

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Pesquisa em ProjAlocAtiv:
        vSelec = "SELECT * FROM ProjAlocAtiv WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 1
                vDesaloca(1) = True
                vSelec1 = vSelec1 & "Atividades do Projeto" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjAlocAtiv")
                VerificaTermino = -1
                Exit Function
            End If
        End If
        'Pesquisa em ProjAlocSAtiv:
        vSelec = "SELECT * FROM ProjAlocSAtiv WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 2
                vDesaloca(2) = True
                vSelec1 = vSelec1 & "Sub-Atividades dos Projetos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjAlocSAtiv")
                VerificaTermino = -2
                Exit Function
            End If
        End If
        'Pesquisa em ProjAlocDetAtiv:
        vSelec = "SELECT * FROM ProjAlocDetAtiv WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 3
                vDesaloca(3) = True
                vSelec1 = vSelec1 & "Detalhes das Atividades dos Projetos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjAlocDetAtiv")
                VerificaTermino = -3
                Exit Function
            End If
        End If
        'Pesquisa em ProjDoctosProj:
        vSelec = "SELECT * FROM ProjDoctosProj WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 4
                vDesaloca(4) = True
                vSelec1 = vSelec1 & "Documentos do Projeto" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjDoctosProj")
                VerificaTermino = -4
                Exit Function
            End If
        End If
        'Pesquisa em ProjFatosRelev:
        vSelec = "SELECT * FROM ProjFatosRelev WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtSolucaoFato, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtSolucaoFato, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 5
                vDesaloca(5) = True
                vSelec1 = vSelec1 & "Ocorrências do Projeto" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjFatosRelev")
                VerificaTermino = -5
                Exit Function
            End If
        End If
        'Pesquisa em ProjFluxo:
        vSelec = "SELECT * FROM ProjFluxo WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtStatus, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 6
                vDesaloca(6) = True
                vSelec1 = vSelec1 & "Status do Projeto" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjFluxo")
                VerificaTermino = -6
                Exit Function
            End If
        End If

        'Pesquisa em BibAlocProj:
        vSelec = "SELECT * FROM BibAlocProj WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 7
                vDesaloca(7) = True
                vSelec1 = vSelec1 & "Alocações em Publicações" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de BibAlocProj")
                VerificaTermino = -7
                Exit Function
            End If
        End If

        'Pesquisa em BibHAlocProj:
        vSelec = "SELECT * FROM BibHAlocProj WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 7
                vDesaloca(7) = True
                vSelec1 = vSelec1 & "Alocações em Publicações" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de BibAlocProj")
                VerificaTermino = -7
                Exit Function
            End If
        End If
        'Pesquisa em CandVagas:
        vSelec = "SELECT * FROM CandVagas WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 8
                vDesaloca(8) = True
                vSelec1 = vSelec1 & "Vagas Abertas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de CandVagas")
                VerificaTermino = -8
                Exit Function
            End If
        End If
        'Pesquisa em OcorAlocProj:
        vSelec = "SELECT * FROM OcorAlocProjetos WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 9
                vDesaloca(9) = True
                vSelec1 = vSelec1 & "Alocações de Colaboradores nos Projetos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAlocProj")
                VerificaTermino = -9
                Exit Function
            End If
        End If
        'Pesquisa em OcorHAlocProj:
        vSelec = "SELECT * FROM OcorHAlocProjetos WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 9
                vDesaloca(9) = True
                vSelec1 = vSelec1 & "Alocações de Colaboradores nos Projetos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAlocProj")
                VerificaTermino = -9
                Exit Function
            End If
        End If
        'Pesquisa em OcorApontHoras:
        vSelec = "SELECT * FROM OcorApontHoras WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, data, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 10
                vDesaloca(10) = True
                vSelec2 = vSelec2 & "Apontamento de Horas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorApontHoras")
                VerificaTermino = -10
                Exit Function
            End If
        End If
        'Pesquisa em OcorHApontHoras:
        vSelec = "SELECT * FROM OcorHApontHoras WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, data, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 10
                vDesaloca(10) = True
                vSelec2 = vSelec2 & "Apontamento de Horas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorApontHoras")
                VerificaTermino = -10
                Exit Function
            End If
        End If
        'Pesquisa em OcorAvaliacColab:
        vSelec = "SELECT * FROM OcorAvaliacColab WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtAvaliacao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 11
                vDesaloca(11) = True
                vSelec1 = vSelec1 & "Avaliação dos Colaboradores" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAvaliacColab")
                VerificaTermino = -11
                Exit Function
            End If
        End If
        'Pesquisa em OcorHAvaliacColab:
        vSelec = "SELECT * FROM OcorHAvaliacColab WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtAvaliacao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 11
                vDesaloca(11) = True
                vSelec1 = vSelec1 & "Avaliação dos Colaboradores" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAvaliacColab")
                VerificaTermino = -11
                Exit Function
            End If
        End If
        'Pesquisa em DirDoctosDir:
        vSelec = "SELECT * FROM DirDoctosDir WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 12
                vDesaloca(12) = True
                vSelec1 = vSelec1 & "Alocações em Diretrizes Administrativas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de DirDoctosDir")
                VerificaTermino = -12
                Exit Function
            End If
        End If
        'Pesquisa em EstDeposito:
        vSelec = "SELECT * FROM EstAlocProjDeposito WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 13
                vDesaloca(13) = True
                vSelec1 = vSelec1 & "Alocações em Depósitos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de EstDeposito")
                VerificaTermino = -13
                Exit Function
            End If
        End If
        'Pesquisa em InvAlocProj:
        vSelec = "SELECT * FROM InvAlocProj WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 14
                vDesaloca(14) = True
                vSelec1 = vSelec1 & "Alocações em Ativos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de InvAlocProj")
                VerificaTermino = -14
                Exit Function
            End If
        End If
        'Pesquisa em InvHAlocProj:
        vSelec = "SELECT * FROM InvHAlocProj WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 14
                vDesaloca(14) = True
                vSelec1 = vSelec1 & "Alocações em Ativos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de InvAlocProj")
                VerificaTermino = -14
                Exit Function
            End If
        End If
        'Pesquisa em ParCadastro:
        vSelec = "SELECT * FROM ParCadastro WHERE identificacao = " & vProjeto
        vSelec = vSelec & " AND tpAtuacao = 1"
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 15
                vDesaloca(15) = True
                vSelec1 = vSelec1 & "Cadastro de Parcerias" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ParCadastro")
                VerificaTermino = -15
                Exit Function
            End If
        End If
        'Pesquisa em ParOcorrenc:
        vSelec = "SELECT ParCadastro.idParceria, tpAtuacao, identificacao, dtFim, ParOcorrenc.dtSolucao"
        vSelec = vSelec & " INTO repParProj"
        vSelec = vSelec & " FROM ParCadastro INNER JOIN ParOcorrenc"
        vSelec = vSelec & " ON ParCadastro.idParceria = ParOcorrenc.idParceria"
        vSelec = vSelec & " WHERE identificacao = " & vProjeto
        vSelec = vSelec & " AND tpAtuacao = 1"
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtSolucao, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtSolucao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        ParCadastro.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        vSelec = "SELECT * FROM repParProj"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 16
                vDesaloca(16) = True
                vSelec1 = vSelec1 & "Ocorrências de Parceria com Projeto Específico" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ParOcorrenc")
                VerificaTermino = -16
                Exit Function
            End If
        End If
        ParCadastro.Db.Execute("DROP TABLE repParProj", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        'Pesquisa em RDNotasExplicat:
        vSelec = "SELECT * FROM RDNotasExplicat WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), dtReferencia, 112) > '" & String.Format(vDtFim, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 17
                vDesaloca(17) = True
                vSelec1 = vSelec1 & "ReceitasDespesas - Notas Explicativas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de RDNotasExplicat")
                VerificaTermino = -17
                Exit Function
            End If
        End If
        'Pesquisa em RDRecDesp:
        vSelec = "SELECT * FROM RDRecDesp WHERE idProjeto = " & Val(vProjeto)
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, competencia, 112) > '" & String.Format(vDtFim, "yyyymm") & "01'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
            'vSelec = vSelec & " OR Convert(datetime, dtEfetivacao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 18
                vDesaloca(18) = True
                vSelec2 = vSelec2 & "Receitas e Despesas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de RDRecDesp")
                VerificaTermino = -18
                Exit Function
            End If
        End If
        'Pesquisa em RDHRecDesp:
        vSelec = "SELECT * FROM RDHRecDesp WHERE idProjeto = " & Val(vProjeto)
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, competencia, 112) > '" & String.Format(vDtFim, "yyyymm") & "01'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
            vSelec = vSelec & " OR Convert(datetime, dtEfetivacao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 18
                vDesaloca(18) = True
                vSelec2 = vSelec2 & "Receitas e Despesas" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de RDRecDesp")
                VerificaTermino = -18
                Exit Function
            End If
        End If
        'Pesquisa em RDFaturamento:
        vSelec = "SELECT * FROM RDFaturamento WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFaturamento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 19
                vDesaloca(19) = True
                vSelec2 = vSelec2 & "Relação de Faturamento" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de RDFaturamento")
                VerificaTermino = -19
                Exit Function
            End If
        End If
        'Pesquisa em RDPagamentos:
        vSelec = "SELECT * FROM RDPagamentos WHERE idProjeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtPagamento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 20
                vDesaloca(20) = True
                vSelec2 = vSelec2 & "Relação de Pagamentos" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de RDPagamentos")
                VerificaTermino = -20
                Exit Function
            End If
        End If
        '* ***********************
        '* Pesquisa no AudiHoras *
        '* ***********************
        'Pesquisa em ApontHoras:
        vSelec = "SELECT * FROM ApontHoras WHERE fk_Projeto = " & vProjeto
        If vTipo >= 1 Then
            vSelec = vSelec & " AND Convert(datetime, data, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistrosAudiHoras(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 21
                vDesaloca(21) = True
                vSelec2 = vSelec2 & "Relação de Projetos no Audihoras" & Chr(10)
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ApontProjetos - Audihoras")
                VerificaTermino = -21
                Exit Function
            End If
        End If

        '* ***************************************************************************************
        '* Se há Alocações Pendentes Impeditivas, não permite continuar com esta data de término *
        '* ***************************************************************************************
        If vTipo = 1 Then
            If VerificaTermino = 10 Or (VerificaTermino >= 18 And VerificaTermino <= 21) Then
                vSelec2 = vSelec2 & "Há data posterior a " & vDtFim & "." & Chr(10)
                vSelec2 = vSelec2 & "Ajuste a data de encerramento do projeto antes de encerrá-lo." & Chr(10)
                MsgBox(vSelec2)
                Exit Function
            End If
        Else
            If VerificaTermino <> 0 And vTipo <> 2 Then
                MsgBox("Há relacionamento com este projeto. Eliminação não autorizada")
                Exit Function
            End If
        End If
        '* ************************************************************************************
        '* Se há Alocações Pendentes, Pergunta se as fechará com a data de término do projeto *
        '* ************************************************************************************
        If vTipo <> 2 Then
            If VerificaTermino > 0 Then
                vSelec1 = vSelec1 & "Encerra as alocações com a data " & vDtFim & " ?" & Chr(10)
                RC = MsgBox(vSelec1, vbOKCancel, "Confirmação!")
                If RC = 2 Then     'Cancela
                    VerificaTermino = -22
                    Exit Function
                Else
                    If vDesaloca(1) = True Then
                        vSelec = "UPDATE ProjAlocAtiv SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'Altera ApontAlocAtiv (Audihoras):
                        vSelec = "UPDATE ApontAlocAtiv SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE fk_Projeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        ApontHoras.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(2) = True Then
                        vSelec = "UPDATE ProjAlocSAtiv SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'Altera ApontAlocSAtiv (Audihoras):
                        vSelec = "UPDATE ApontAlocSAtiv SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE fk_Projeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        ApontHoras.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(3) = True Then
                        vSelec = "UPDATE ProjAlocDetAtiv SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(4) = True Then
                        vSelec = "UPDATE ProjDoctosProj SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(5) = True Then
                        vSelec = "UPDATE ProjFatosRelev SET dtSolucaoFato = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtSolucaoFato, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtSolucaoFato, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(6) = True Then
                        vSelec = "UPDATE ProjFluxo SET dtStatus = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND Convert(datetime, dtStatus, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(7) = True Then
                        vSelec = "UPDATE BibAlocProj SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(8) = True Then
                        vSelec = "UPDATE CandVagas SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(9) = True Then
                        vSelec = "UPDATE OcorAlocProjetos SET dtFinalProjeto = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'Altera ApontProjetos (Audihoras):
                        vSelec = "UPDATE ApontProjetos SET dtFinalProj = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & "WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFinalProj, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFinalProj, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        ApontHoras.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(11) = True Then
                        vSelec = "UPDATE OcorAvaliacColab SET dtAvaliacao = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND Convert(datetime, dtAvaliacao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(12) = True Then
                        vSelec = "UPDATE DirDoctosDir SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(13) = True Then
                        vSelec = "UPDATE EstDeposito SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(14) = True Then
                        vSelec = "UPDATE InvAlocProj SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(15) = True Then
                        vSelec = "UPDATE ParCadastro SET dtFim = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                        vSelec = vSelec & " WHERE identificacao = " & vProjeto
                        vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
                        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                        vSelec = vSelec & " AND tpAtuacao = 1"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If vDesaloca(16) = True Then
                        vSelec = "SELECT * FROM ParCadastro WHERE identificacao = " & vProjeto
                        vSelec = vSelec & " AND tpAtuacao = 1"
                        RC = ParCadastro.dbConecta(0, 1, vSelec)
                        If RC <> 0 Then
                            MsgBox(" Erro na abertura de ParCadastro")
                            VerificaTermino = -23
                            Exit Function
                        End If
                        RC = ParCadastro.leSeq(1)
                        Do While RC = 0
                            vSelec = "UPDATE ParOcorrenc SET dtSolucao = Convert(datetime, '" & String.Format(vDtFim, "yyyymmdd") & "', 112) "
                            vSelec = vSelec & "WHERE idParceria = " & ParCadastro.idParceria
                            vSelec = vSelec & " AND (Convert(datetime, dtSolucao, 112) = '18991230'"
                            vSelec = vSelec & " OR Convert(datetime, dtSolucao, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
                            segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            RC = ParCadastro.leSeq(0)
                        Loop
                    End If
                    If vDesaloca(17) = True Then
                        vSelec = "UPDATE RDNotasExplicat SET dtReferencia = Convert(datetime, '" & String.Format(vDtFim, "yyyymm") & "01', 112) "
                        vSelec = vSelec & " WHERE idProjeto = " & vProjeto
                        vSelec = vSelec & " AND Convert(nvarchar(6), dtReferencia, 112) > '" & String.Format(vDtFim, "yyyymm") & "'"
                        segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    VerificaTermino = 0
                End If
            End If
        End If

EVerificaTermino:
        If Err.Number Then
            Select Case Err.Number
            'Tabela já existe
                Case -2147217900
                    segMensagens.Db.Execute("DROP TABLE repParProj", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    segMensagens.Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    Resume Next
                Case Else
                    MsgBox(Err.Number & " - " & Err.Description)
                    'Altera a estilizacao do cursor para o padrão
                    Cursor.Current = Cursors.Default
            End Select
        End If

    End Function
End Class
