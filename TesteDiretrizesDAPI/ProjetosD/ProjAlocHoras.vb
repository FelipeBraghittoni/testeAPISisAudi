Option Explicit On
Public Class ProjAlocHoras
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850708"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320708"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0708"
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

    Private mvaridProposta As String
    Public Property idProposta() As String
        Get
            Return mvaridProposta
        End Get
        Set(value As String)
            mvaridProposta = value
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

    Private mvarhoras As Single
    Public Property horas() As Single
        Get
            Return mvarhoras
        End Get
        Set(value As Single)
            mvarhoras = value
        End Set
    End Property

    Private mvaridUsuario As String
    Public Property idUsuario() As String
        Get
            Return mvaridUsuario
        End Get
        Set(value As String)
            mvaridUsuario = value
        End Set
    End Property

    Private mvardtHsAlteracao As Date
    Public Property dtHsAlteracao() As Date
        Get
            Return mvardtHsAlteracao
        End Get
        Set(value As Date)
            mvardtHsAlteracao = value
        End Set
    End Property

    Private mvardcHoras As String
    Public Property dcHoras() As String
        Get
            Return mvardcHoras
        End Get
        Set(value As String)
            mvardcHoras = value
        End Set
    End Property

    Private mvarorigem As Short
    Public Property origem() As Short
        Get
            Return mvarorigem
        End Get
        Set(value As Short)
            mvarorigem = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tProjAlocHoras As ADODB.Recordset
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
        tProjAlocHoras = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjAlocHoras.Open("ProjAlocHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocHoras.Close()

                If tipo = 1 Then
                    tProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjAlocHoras.Open("ProjAlocHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjAlocHoras - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjAlocHoras.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjAlocHoras.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjAlocHoras.EOF Then
            idProjeto = tProjAlocHoras.Fields("idProjeto").Value
            idAtividade = tProjAlocHoras.Fields("idAtividade").Value
            idSAtividade = tProjAlocHoras.Fields("idSAtividade").Value
            idProposta = tProjAlocHoras.Fields("idProposta").Value
            dtInicio = tProjAlocHoras.Fields("dtInicio").Value
            dtFim = tProjAlocHoras.Fields("dtFim").Value
            horas = tProjAlocHoras.Fields("horas").Value
            idUsuario = tProjAlocHoras.Fields("idUsuario").Value
            dtHsAlteracao = tProjAlocHoras.Fields("dtHsAlteracao").Value
            'Especifica se é inclusão (C) ou exclusão (D) de horas:
            dcHoras = tProjAlocHoras.Fields("dcHoras").Value
            'Origem da alocação (1 - manual) (2 - sistema):
            origem = tProjAlocHoras.Fields("origem").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjAlocHoras - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localizaExt(projeto As Double, ativ As Short, sAtiv As Short, proposta As String, inicio As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjAlocHoras.                       *
        '* *                                      *
        '* * proposta = identif. para pesquisa    *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelec As String

        On Error GoTo ElocalizaExt
        vSelec = "SELECT * FROM ProjAlocHoras WHERE idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        If proposta = " " Then
            vSelec = vSelec & " AND idProposta = ' '"
        Else
            vSelec = vSelec & " AND idProposta = '" & Trim(proposta) & "'"
        End If
        vSelec = vSelec & " AND dtInicio = Convert(datetime, '" & String.Format(inicio, "yyyymmdd") & "', 112)"
        tProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaExt = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjAlocHoras.EOF Then
            localizaExt = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjAlocHoras.Fields("idProjeto").Value
            idAtividade = tProjAlocHoras.Fields("idAtividade").Value
            idSAtividade = tProjAlocHoras.Fields("idSAtividade").Value
            idProposta = tProjAlocHoras.Fields("idProposta").Value
            dtInicio = tProjAlocHoras.Fields("dtInicio").Value
            dtFim = tProjAlocHoras.Fields("dtFim").Value
            horas = tProjAlocHoras.Fields("horas").Value
            idUsuario = tProjAlocHoras.Fields("idUsuario").Value
            dtHsAlteracao = tProjAlocHoras.Fields("dtHsAlteracao").Value
            'Especifica se é inclusão (C) ou exclusão (D) de horas:
            dcHoras = tProjAlocHoras.Fields("dcHoras").Value
            'Origem da alocação (1 - manual) (2 - sistema):
            origem = tProjAlocHoras.Fields("origem").Value
        End If

ElocalizaExt:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjAlocHoras.Close()
                tProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaExt = Err.Number
                MsgBox("Classe ProjAlocHoras - localizaExt" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(descric As String) As Integer

        On Error GoTo Einclui
        tProjAlocHoras.AddNew()
        tProjAlocHoras.Fields("idProjeto").Value = idProjeto
        tProjAlocHoras.Fields("idAtividade").Value = idAtividade
        tProjAlocHoras.Fields("idSAtividade").Value = idSAtividade
        tProjAlocHoras.Fields("idProposta").Value = idProposta
        tProjAlocHoras.Fields("dtInicio").Value = dtInicio
        tProjAlocHoras.Fields("dtFim").Value = dtFim
        tProjAlocHoras.Fields("horas").Value = horas
        tProjAlocHoras.Fields("idUsuario").Value = idUsuario
        tProjAlocHoras.Fields("dtHsAlteracao").Value = dtHsAlteracao
        'Carrega Descrição da Alocação das Horas:
        tProjAlocHoras.Fields("descricAlocHoras").Value = IIf(descric = "", " ", descric)
        'Especifica se é inclusão (C) ou exclusão (D) de horas:
        tProjAlocHoras.Fields("dcHoras").Value = dcHoras
        'Origem da alocação (1 - manual) (2 - sistema):
        tProjAlocHoras.Fields("origem").Value = origem
        tProjAlocHoras.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjAlocHoras - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(descric As String) As Integer

        On Error GoTo Ealtera
        tProjAlocHoras.Fields("idProjeto").Value = idProjeto
        tProjAlocHoras.Fields("idAtividade").Value = idAtividade
        tProjAlocHoras.Fields("idSAtividade").Value = idSAtividade
        tProjAlocHoras.Fields("idProposta").Value = idProposta
        tProjAlocHoras.Fields("dtInicio").Value = dtInicio
        tProjAlocHoras.Fields("dtFim").Value = dtFim
        tProjAlocHoras.Fields("horas").Value = horas
        tProjAlocHoras.Fields("idUsuario").Value = idUsuario
        tProjAlocHoras.Fields("dtHsAlteracao").Value = dtHsAlteracao
        'Carrega Descrição da Alocação das Horas:
        tProjAlocHoras.Fields("descricAlocHoras").Value = IIf(descric = "", " ", descric)
        'Especifica se é inclusão (C) ou exclusão (D) de horas:
        tProjAlocHoras.Fields("dcHoras").Value = dcHoras
        'Origem da alocação (1 - manual) (2 - sistema):
        tProjAlocHoras.Fields("origem").Value = origem
        tProjAlocHoras.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjAlocHoras - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function descricAlocHoras() As String

        descricAlocHoras = tProjAlocHoras.Fields("descricAlocHoras").Value

    End Function

    Public Function elimina(projeto As Double, ativ As Short, sAtiv As Short, proposta As String, inicio As Date) As Integer
        '* *************************************
        '* * Elimina um registro da tabela     *
        '* * ProjAlocHoras                     *
        '* *                                   *
        '* * projeto  = identif. para pesquisa *
        '* * ativ     = identif. para pesquisa *
        '* * sAtiv    = identif. para pesquisa *
        '* * proposta = identif. para pesquisa *
        '* * inicio   = identif. para pesquisa *
        '* *************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        vSelec = "DELETE FROM ProjAlocHoras WHERE idProjeto = " & projeto
        vSelec = vSelec & " AND idAtividade = " & ativ
        vSelec = vSelec & " AND idSAtividade = " & sAtiv
        If proposta = "               " Then
            vSelec = vSelec & " AND idProposta = ' '"
        Else
            vSelec = vSelec & " AND idProposta = '" & Trim(proposta) & "'"
        End If
        vSelec = vSelec & " AND dtInicio = Convert(datetime, '" & String.Format(inicio, "yyyymmdd") & "', 112)"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjAlocHoras - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CalcHorasAloc(vProj As Double, vAtiv As Short, vSAtiv As Short, vPeriodo As Date) As Integer

        Dim vSelec As String
        Dim RC As Integer
        Dim vHsAlocC As Long
        Dim vHsAlocD As Long

        vSelec = "SELECT * FROM ProjAlocHoras WHERE idProjeto = " & vProj
        vSelec = vSelec & " AND idAtividade = " & vAtiv
        vSelec = vSelec & " AND idSAtividade = " & vSAtiv
        vSelec = vSelec & " AND Convert(nvarchar(6), dtInicio, 112) <= '" & String.Format(vPeriodo, "yyyymm") & " '"
        vSelec = vSelec & " AND (Convert(nvarchar(6), dtFim, 112) >= '" & String.Format(vPeriodo, "yyyymm") & "'"
        vSelec = vSelec & " OR Convert(datetime, dtFim, 112) = '18991230')"
        vSelec = vSelec & " ORDER BY dtInicio"
        RC = dbConecta(0, 1, vSelec)
        RC = leSeq(1)
        Do While RC = 0
            If dcHoras = "D" Then
                vHsAlocD = vHsAlocD + horas
            Else
                vHsAlocC = vHsAlocC + horas
            End If
            'Lê o próximo registro:
            RC = leSeq(0)
        Loop
        If RC = 1016 Then
            CalcHorasAloc = vHsAlocC - vHsAlocD
        End If

    End Function

    Public Function AtualizaHsAloc(vProj As Double, vAtiv As Short, vSAtiv As Short, vPeriodo As Date, vHoras As Integer) As Integer

        Dim vSelec As String
        Dim RC As Integer
        Dim I As Short
        Dim J As Short
        Dim vDiaF As Short
        Dim vData As Date
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segUsuario As SegurancaD.segUsuario

        On Error GoTo EAtualizaHsAloc

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais
        'Para buscar o usuário que processou o sistema:
        segUsuario = New SegurancaD.segUsuario

        '* ****************************************************
        '* Verifica se o registro em ProjAlocHoras ja existe: *
        '* ****************************************************
        vSelec = "SELECT * FROM ProjAlocHoras WHERE idProjeto = " & vProj
        vSelec = vSelec & " AND idAtividade = " & vAtiv
        vSelec = vSelec & " AND idSAtividade = " & vSAtiv
        vSelec = vSelec & " AND Convert(datetime, dtInicio, 112) = '" & String.Format(vPeriodo, "yyyymmdd") & "'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC = 0 Then      'Não existe
            '* *************************************
            '* Inclui o registro em ProjAlocHoras: *
            '* *************************************
            vSelec = "INSERT INTO ProjAlocHoras (idProjeto, idAtividade, idSAtividade, idProposta, dtInicio, "
            vSelec = vSelec & "dtFim, horas, idUsuario, dtHsAlteracao, descricAlocHoras, dcHoras, origem)"
            vSelec = vSelec & " VALUES (" & vProj & ", "
            vSelec = vSelec & vAtiv & ", "
            vSelec = vSelec & vSAtiv & ", "
            'Como não existe, grava com Proposta = " ":
            vSelec = vSelec & "' ', "
            'Primeiro dia do mês:
            vSelec = vSelec & " Convert(datetime, '" & String.Format(vPeriodo, "yyyymmdd") & "', 112), "
            'Último dia do mês:
            vSelec = vSelec & " Convert(datetime, '" & String.Format(DateAdd("d", -1, DateAdd("M", 1, vPeriodo)), "yyyymmdd") & "', 112), "
            vSelec = vSelec & vHoras & ", '"
            vSelec = vSelec & Left(segUsuario.leDadosParaLog, 25) & "', "
            vSelec = vSelec & "Convert(DateTime, '" & String.Format(PesqReg.dataServ(), "yyyy-mm-dd hh:mm:ss") & "', 120), '"
            vSelec = vSelec & "Horas consumidas em " & String.Format(vPeriodo, "mm/yyyy") & "', '"
            vSelec = vSelec & "D', "
            vSelec = vSelec & "2)"
            Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            'Já existe um registro:
            If RC = 1 Then
                'Verifica se foi gerado pelo sistema:
                vSelec = "SELECT * FROM ProjAlocHoras WHERE idProjeto = " & vProj
                vSelec = vSelec & " AND idAtividade = " & vAtiv
                vSelec = vSelec & " AND idSAtividade = " & vSAtiv
                vSelec = vSelec & " AND idProposta = ' '"
                vSelec = vSelec & " AND Convert(datetime, dtInicio, 112) = '" & String.Format(vPeriodo, "yyyymmdd") & "'"
                vSelec = vSelec & " AND origem = 2"
                RC = PesqReg.pesqRegistros(vSelec)
                'Não foi gerado pelo sistema:
                If RC = 0 Then
                    'Pesquisa até encontrar um dia sem horas alocadas:
                    For I = 1 To Day(DateAdd("d", -1, DateAdd("M", 1, vPeriodo)))
                        vSelec = "SELECT * FROM ProjAlocHoras WHERE idProjeto = " & vProj
                        vSelec = vSelec & " AND idAtividade = " & vAtiv
                        vSelec = vSelec & " AND idSAtividade = " & vSAtiv
                        vSelec = vSelec & " AND idProposta = ' '"
                        vSelec = vSelec & " AND Convert(datetime, dtInicio, 112) = '" & String.Format(vPeriodo, "yyyymm") & String.Format(I, "0#") & "'"
                        RC = PesqReg.pesqRegistros(vSelec)
                        'Não encontrou um registro com esta chave:
                        If RC = 0 Then
                            '* *************************************
                            '* Inclui o registro em ProjAlocHoras: *
                            '* *************************************
                            vSelec = "INSERT INTO ProjAlocHoras (idProjeto, idAtividade, idSAtividade, idProposta, dtInicio, "
                            vSelec = vSelec & "dtFim, horas, idUsuario, dtHsAlteracao, descricAlocHoras, dcHoras, origem)"
                            vSelec = vSelec & " VALUES (" & vProj & ", "
                            vSelec = vSelec & vAtiv & ", "
                            vSelec = vSelec & vSAtiv & ", "
                            vSelec = vSelec & "' ', "
                            'Primeiro dia do mês:
                            vSelec = vSelec & " Convert(datetime, '" & String.Format(vPeriodo, "yyyymm") & String.Format(I, "0#") & "', 112), "
                            'Último dia do mês:
                            vSelec = vSelec & " Convert(datetime, '" & String.Format(DateAdd("d", -1, DateAdd("M", 1, DateAdd("d", I, vPeriodo))), "yyyymmdd") & "', 112), "
                            vSelec = vSelec & vHoras & ", '"
                            vSelec = vSelec & Left(segUsuario.leDadosParaLog, 25) & "', "
                            vSelec = vSelec & "Convert(DateTime, '" & String.Format(PesqReg.dataServ(), "yyyy-mm-dd hh:mm:ss") & "', 120), '"
                            vSelec = vSelec & "Horas consumidas em " & String.Format(vPeriodo, "mm/yyyy") & "', '"
                            vSelec = vSelec & "D', "
                            vSelec = vSelec & "2)"
                            Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    Next I
                Else
                    'Foi gerado pelo sistema; atualiza Horas:
                    If RC = 1 Then
                        vSelec = "UPDATE ProjAlocHoras SET horas = " & vHoras
                        vSelec = vSelec & " WHERE idProjeto = " & vProj
                        vSelec = vSelec & " AND idAtividade = " & vAtiv
                        vSelec = vSelec & " AND idSAtividade = " & vSAtiv
                        vSelec = vSelec & " AND idProposta = ' '"
                        vSelec = vSelec & " AND Convert(datetime, dtInicio, 112) = '" & String.Format(vPeriodo, "yyyymmdd") & "'"
                        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            Else
                vSelec = "Erro na função AtualizaHsAloc:" & Chr(10)
                vSelec = vSelec & "Projeto: " & vProj & Chr(10)
                vSelec = vSelec & "Horas: " & vHoras & Chr(10)
                MsgBox(vSelec)
            End If
        End If

EAtualizaHsAloc:
        If Err.Number Then
            vSelec = "Erro na função AtualizaHsAloc:" & Chr(10)
            vSelec = vSelec & "Projeto: " & vProj & Chr(10)
            vSelec = vSelec & "Horas: " & vHoras & Chr(10)
            MsgBox(vSelec & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
