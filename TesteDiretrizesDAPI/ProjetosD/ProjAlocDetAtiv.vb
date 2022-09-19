Option Explicit On
Public Class ProjAlocDetAtiv
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850706"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320706"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0706"
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

    Private mvaridDetAtividade As Short
    Public Property idDetAtividade() As Short
        Get
            Return mvaridDetAtividade
        End Get
        Set(value As Short)
            mvaridDetAtividade = value
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

#End Region

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tProjAlocDetAtiv As ADODB.Recordset
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
        tProjAlocDetAtiv = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjAlocDetAtiv.Open("ProjAlocDetAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjAlocDetAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocDetAtiv.Close()

                If tipo = 1 Then
                    tProjAlocDetAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjAlocDetAtiv.Open("ProjAlocDetAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjAlocDetAtiv - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjAlocDetAtiv.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjAlocDetAtiv.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjAlocDetAtiv.EOF Then
            idProjeto = tProjAlocDetAtiv.Fields("idProjeto").Value
            idAtividade = tProjAlocDetAtiv.Fields("idAtividade").Value
            idSAtividade = tProjAlocDetAtiv.Fields("idSAtividade").Value
            idDetAtividade = tProjAlocDetAtiv.Fields("idDetAtividade").Value
            dtInicio = tProjAlocDetAtiv.Fields("dtInicio").Value
            dtFim = tProjAlocDetAtiv.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjAlocDetAtiv - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(projeto As Double, atividade As Short, sAtividade As Short, dAtividade As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjFluxo.                           *
        '* *                                      *
        '* * projeto   = identif. para pesquisa + *
        '* * atividade = identif. para pesquisa + *
        '* * satividade= identif. para pesquisa + *
        '* * datividade= identif. para pesquisa   *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjAlocDetAtiv WHERE idProjeto = " & projeto
        vSelect = vSelect & " AND idAtividade = " & atividade
        vSelect = vSelect & " AND idSAtividade = " & sAtividade
        vSelect = vSelect & " AND idDetAtividade = " & dAtividade
        tProjAlocDetAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjAlocDetAtiv.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjAlocDetAtiv.Fields("idProjeto").Value
            idAtividade = tProjAlocDetAtiv.Fields("idAtividade").Value
            idSAtividade = tProjAlocDetAtiv.Fields("idSAtividade").Value
            idDetAtividade = tProjAlocDetAtiv.Fields("idDetAtividade").Value
            dtInicio = tProjAlocDetAtiv.Fields("dtInicio").Value
            dtFim = tProjAlocDetAtiv.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocDetAtiv.Close()
                tProjAlocDetAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjAlocDetAtiv - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tProjAlocDetAtiv.AddNew()
        tProjAlocDetAtiv(0).Value = idProjeto
        tProjAlocDetAtiv(1).Value = idAtividade
        tProjAlocDetAtiv(2).Value = idSAtividade
        tProjAlocDetAtiv(3).Value = idDetAtividade
        tProjAlocDetAtiv(4).Value = dtInicio
        tProjAlocDetAtiv(5).Value = dtFim
        tProjAlocDetAtiv.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjAlocDetAtiv - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tProjAlocDetAtiv(4).Value = dtInicio
        tProjAlocDetAtiv(5).Value = dtFim
        tProjAlocDetAtiv.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjAlocDetAtiv - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(projeto As Double, ativ As Short, sAtiv As Short, dAtiv As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjAlocDetAtiv.                     *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* * ativ       = identif. para pesquisa  *
        '* * sAtiv      = identif. para pesquisa  *
        '* * dAtiv      = identif. para pesquisa  *
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
            elimina = -1
            Exit Function
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em OcorApontHoras:
        vSelec = "SELECT * FROM OcorApontHoras WHERE idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        vSelec = vSelec & " AND idDetAtividade = " & dAtiv
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If
        'Pesquisa em OcorHApontHoras:
        vSelec = "SELECT * FROM OcorHApontHoras WHERE idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        vSelec = vSelec & " AND idDetAtividade = " & dAtiv
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM ProjAlocDetAtiv WHERE "
        vSelec = vSelec & " idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        vSelec = vSelec & " AND idDetAtividade = " & dAtiv
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjAlocDetAtiv - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
