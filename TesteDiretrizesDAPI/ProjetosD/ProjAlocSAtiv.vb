Option Explicit On
Imports System.Windows.Forms

Public Class ProjAlocSAtiv
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850709"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320709"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0709"
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

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tProjAlocSAtiv As ADODB.Recordset
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
        tProjAlocSAtiv = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjAlocSAtiv.Open("ProjAlocSAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjAlocSAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocSAtiv.Close()

                If tipo = 1 Then
                    tProjAlocSAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjAlocSAtiv.Open("ProjAlocSAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjAlocSAtiv - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjAlocSAtiv.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjAlocSAtiv.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjAlocSAtiv.EOF Then
            idProjeto = tProjAlocSAtiv.Fields("idProjeto").Value
            idAtividade = tProjAlocSAtiv.Fields("idAtividade").Value
            idSAtividade = tProjAlocSAtiv.Fields("idSAtividade").Value
            dtInicio = tProjAlocSAtiv.Fields("dtInicio").Value
            dtFim = tProjAlocSAtiv.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjAlocSAtiv - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(projeto As Double, atividade As Short, sAtividade As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjAlocSAtiv.                       *
        '* *                                      *
        '* * projeto   = identif. para pesquisa + *
        '* * atividade = identif. para pesquisa + *
        '* * satividade= identif. para pesquisa + *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjAlocSAtiv WHERE idProjeto = " & projeto
        vSelect = vSelect & " AND idAtividade = " & atividade
        vSelect = vSelect & " AND idSAtividade = " & sAtividade
        tProjAlocSAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjAlocSAtiv.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjAlocSAtiv.Fields("idProjeto").Value
            idAtividade = tProjAlocSAtiv.Fields("idAtividade").Value
            idSAtividade = tProjAlocSAtiv.Fields("idSAtividade").Value
            dtInicio = tProjAlocSAtiv.Fields("dtInicio").Value
            dtFim = tProjAlocSAtiv.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocSAtiv.Close()
                tProjAlocSAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjAlocSAtiv - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tProjAlocSAtiv.AddNew()
        tProjAlocSAtiv(0).Value = idProjeto
        tProjAlocSAtiv(1).Value = idAtividade
        tProjAlocSAtiv(2).Value = idSAtividade
        tProjAlocSAtiv(3).Value = dtInicio
        tProjAlocSAtiv(4).Value = dtFim
        tProjAlocSAtiv.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjAlocSAtiv - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tProjAlocSAtiv(3).Value = dtInicio
        tProjAlocSAtiv(4).Value = dtFim
        tProjAlocSAtiv.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjAlocSAtiv - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(projeto As Double, ativ As Short, sAtiv As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjAlocSAtiv.                       *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* * ativ       = identif. para pesquisa  *
        '* * sAtiv      = identif. para pesquisa  *
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
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM ProjAlocSAtiv WHERE "
        vSelec = vSelec & " idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjAlocSAtiv - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function VerificaTermino(vTipo As Short, vProjeto As Double, vAtividade As Short, vSAtividade As Short, Optional vDtFim As Date = Nothing) As Integer
        '* ***************************************************
        '* Verifica se há pendências a serem resolvidas      *
        '* antes de efetivar o término/eliminação da         *
        '* Atividade alocada no Projeto                      *
        '* ***************************************************
        '* vtipo pode ser:                                   *
        '* 0 - não verifica dtFim; utilizado para eliminar   *
        '* 1 - verifica dtFim; utilizado para encerrar       *
        '* ***************************************************
        '* VerificaTermino pode variar de:                   *
        '*  < 0 = ocorreu erro                               *
        '*  = 0 = não encontrou nenhuma alocação pendente    *
        '*  1 a 20 = Alocações pendentes                     *
        '* ***************************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim vSelec1 As String
        Dim vSelec2 As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As New SegurancaD.segMensagens

        On Error GoTo EVerificaTermino

        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro na abertura de segMensagens")
            VerificaTermino = -1
            Exit Function
        End If

        VerificaTermino = 0     'Se não houver erro
        vSelec1 = "Há alocações da Sub-Atividade em aberto para este Projeto:" & Chr(10)
        vSelec2 = "Há alocações em aberto Impeditivas para este Projeto:" & Chr(10)

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Pesquisa em ProjAlocDetAtiv:
        vSelec = "SELECT * FROM ProjAlocDetAtiv WHERE idProjeto = " & vProjeto
        vSelec = vSelec & " AND idAtividade = " & vAtividade
        vSelec = vSelec & " AND idSAtividade = " & vSAtividade
        If vTipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 2
                vSelec1 = vSelec1 & "Detalhes das Atividades dos Projetos" & Chr(10)
                Exit Function
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de ProjAlocDetAtiv")
                VerificaTermino = -2
                Exit Function
            End If
        End If
        'Pesquisa em OcorAlocProj:
        vSelec = "SELECT * FROM OcorAlocProjetos WHERE idProjeto = " & vProjeto
        vSelec = vSelec & " AND idAtividade = " & vAtividade
        vSelec = vSelec & " AND idSAtividade = " & vSAtividade
        If vTipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 3
                vSelec1 = vSelec1 & "Alocações de Colaboradores nos Projetos" & Chr(10)
                Exit Function
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAlocProj")
                VerificaTermino = -3
                Exit Function
            End If
        End If
        'Pesquisa em OcorHAlocProj:
        vSelec = "SELECT * FROM OcorHAlocProjetos WHERE idProjeto = " & vProjeto
        vSelec = vSelec & " AND idAtividade = " & vAtividade
        vSelec = vSelec & " AND idSAtividade = " & vSAtividade
        If vTipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 4
                vSelec1 = vSelec1 & "Alocações de Colaboradores Histórico nos Projetos" & Chr(10)
                Exit Function
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorAlocProj")
                VerificaTermino = -4
                Exit Function
            End If
        End If
        'Pesquisa em OcorApontHoras:
        vSelec = "SELECT * FROM OcorApontHoras WHERE idProjeto = " & vProjeto
        vSelec = vSelec & " AND idAtividade = " & vAtividade
        vSelec = vSelec & " AND idSAtividade = " & vSAtividade
        If vTipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, data, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 5
                vSelec2 = vSelec2 & "Apontamento de Horas" & Chr(10)
                Exit Function
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorApontHoras")
                VerificaTermino = -5
                Exit Function
            End If
        End If
        'Pesquisa em OcorHApontHoras:
        vSelec = "SELECT * FROM OcorHApontHoras WHERE idProjeto = " & vProjeto
        vSelec = vSelec & " AND idAtividade = " & vAtividade
        vSelec = vSelec & " AND idSAtividade = " & vSAtividade
        If vTipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, data, 112) > '" & String.Format(vDtFim, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            'Encontrou:
            If RC = 1 Then
                VerificaTermino = 6
                vSelec2 = vSelec2 & "Apontamento de Horas Histórico" & Chr(10)
                Exit Function
            End If
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                MsgBox("Erro na Pesquisa de OcorApontHoras")
                VerificaTermino = -6
                Exit Function
            End If
        End If

EVerificaTermino:
        If Err.Number Then
            MsgBox(Err.Number & " - " & Err.Description)
            'old Screen.MousePointer = vbDefault
            Cursor.Current = Cursors.Default
        End If

    End Function

End Class
