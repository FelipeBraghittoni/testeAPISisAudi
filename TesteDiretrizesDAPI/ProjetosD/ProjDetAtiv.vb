Option Explicit On
Public Class ProjDetAtiv
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850711"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320711"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0711"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridDetAtividade As Short
    Public Property idDetAtividade() As Short
        Get
            Return mvaridDetAtividade
        End Get
        Set(value As Short)
            mvaridDetAtividade = value
        End Set
    End Property

    Private mvarnomeDetAtiv As String
    Public Property nomeDetAtiv() As String
        Get
            Return mvarnomeDetAtiv
        End Get
        Set(value As String)
            mvarnomeDetAtiv = value
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
    Public tProjDetAtiv As ADODB.Recordset
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
        tProjDetAtiv = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjDetAtiv.Open("ProjDetAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjDetAtiv ORDER BY nomeDetAtiv"
                tProjDetAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjDetAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjDetAtiv.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjDetAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjDetAtiv.Open("ProjDetAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic,)
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjDetAtiv - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjDetAtiv.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjDetAtiv.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjDetAtiv.EOF Then
            idDetAtividade = tProjDetAtiv.Fields("idDetAtividade").Value
            nomeDetAtiv = tProjDetAtiv.Fields("nomeDetAtiv").Value
            dtInicio = tProjDetAtiv.Fields("dtInicio").Value
            dtFim = tProjDetAtiv.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjDetAtiv - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjDetAtiv.                         *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjDetAtiv WHERE nomeDetAtiv = '" & descric & "'"
        tProjDetAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjDetAtiv.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idDetAtividade = tProjDetAtiv.Fields("idDetAtividade").Value
            nomeDetAtiv = tProjDetAtiv.Fields("nomeDetAtiv").Value
            dtInicio = tProjDetAtiv.Fields("dtInicio").Value
            dtFim = tProjDetAtiv.Fields("dtFim").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjDetAtiv.Close()
                tProjDetAtiv.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjDetAtiv - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tProjDetAtiv.AddNew()
        tProjDetAtiv(0).Value = cCodigo
        tProjDetAtiv(1).Value = cDescricao
        tProjDetAtiv.Fields("dtInicio").Value = dtInicio
        tProjDetAtiv.Fields("dtFim").Value = dtFim
        tProjDetAtiv.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjDetAtiv - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tProjDetAtiv(1).Value = cDescricao
        tProjDetAtiv.Fields("dtInicio").Value = dtInicio
        tProjDetAtiv.Fields("dtFim").Value = dtFim
        tProjDetAtiv.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjDetAtiv - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjDetAtiv.                         *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
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

        '* **********************************************
        '* Verifica se não fere integridade relacional, *
        '* quando controlada por código:                *
        '* **********************************************
        'Pesquisa em ProjAlocDetAtiv:
        vSelec = "SELECT * FROM ProjAlocDetAtiv WHERE idDetAtividade = " & codigo
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há detalhes de atividade alocadas em Projetos." & Err.Number & "Eliminação negada.")
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            Exit Function
        End If

        vSelec = "DELETE FROM ProjDetAtiv WHERE "
        vSelec = vSelec & " idDetAtividade = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjDetAtiv - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
