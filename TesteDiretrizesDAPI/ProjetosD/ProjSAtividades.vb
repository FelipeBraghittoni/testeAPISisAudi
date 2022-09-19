Option Explicit On
Public Class ProjSAtividades

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850722"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320722"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0722"
#End Region

#Region "Variaveis de ambiente"

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
    Public tProjSAtividades As ADODB.Recordset
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
        tProjSAtividades = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjSAtividades.Open("ProjSAtividades", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjSAtividades ORDER BY nomeSAtividade"
                tProjSAtividades.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjSAtividades.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjSAtividades.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjSAtividades.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjSAtividades.Open("ProjSAtividades", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjSAtividades - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjSAtividades.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjSAtividades.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjSAtividades.EOF Then
            idSAtividade = tProjSAtividades.Fields("idSAtividade").Value
            nomeSAtividade = tProjSAtividades.Fields("nomeSAtividade").Value
            dtInicio = tProjSAtividades.Fields("dtInicio").Value
            dtFim = tProjSAtividades.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjSAtividades - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(sAtividade As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjSAtividades.                     *
        '* *                                      *
        '* * sAtividade  = identif. para pesquisa *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjSAtividades WHERE idSAtividade = " & sAtividade
        tProjSAtividades.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjSAtividades.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idSAtividade = tProjSAtividades.Fields("idSAtividade").Value
            nomeSAtividade = tProjSAtividades.Fields("nomeSAtividade").Value
            dtInicio = tProjSAtividades.Fields("dtInicio").Value
            dtFim = tProjSAtividades.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjSAtividades.Close()
                tProjSAtividades.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjSAtividades - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjSAtividades.                     *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjSAtividades WHERE nomeSAtividade = '" & descric & "'"
        tProjSAtividades.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)


        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjSAtividades.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idSAtividade = tProjSAtividades.Fields("idSAtividade").Value
            nomeSAtividade = tProjSAtividades.Fields("nomeSAtividade").Value
            dtInicio = tProjSAtividades.Fields("dtInicio").Value
            dtFim = tProjSAtividades.Fields("dtFim").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjSAtividades.Close()
                tProjSAtividades.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjSAtividades - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tProjSAtividades.AddNew()
        tProjSAtividades(0).Value = cCodigo
        tProjSAtividades(1).Value = cDescricao
        tProjSAtividades.Fields("dtInicio").Value = dtInicio
        tProjSAtividades.Fields("dtFim").Value = dtFim
        tProjSAtividades.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjSAtividades - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tProjSAtividades(1).Value = cDescricao
        tProjSAtividades.Fields("dtInicio").Value = dtInicio
        tProjSAtividades.Fields("dtFim").Value = dtFim
        tProjSAtividades.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjSAtividades - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjSAtividades.                     *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens

        If RC <> 0 Then
            MsgBox("Erro na abertura de segMensagens")
            elimina = -1
            Exit Function
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        '* **********************************************
        '* Verifica se não fere integridade relacional, *
        '* quando controlada por código:                *
        '* **********************************************
        'Pesquisa em ProjAlocSAtiv:
        vSelec = "SELECT * FROM ProjAlocSAtiv WHERE idSAtividade = " & codigo
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há sub-atividades alocadas em Projetos." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If
        'Pesquisa em OcorApontHoras:
        vSelec = "SELECT * FROM OcorApontHoras WHERE idSAtividade = " & codigo
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há Apontamento de Horas nesta Sub-Atividade." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If
        'Pesquisa em OcorSaldoHoras:
        vSelec = "SELECT * FROM OcorSaldoHoras WHERE idSAtividade = " & codigo
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há Saldo de Horas nesta Sub-Atividade." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM ProjSAtividades WHERE "
        vSelec = vSelec & " idSAtividade = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjSAtividades - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
