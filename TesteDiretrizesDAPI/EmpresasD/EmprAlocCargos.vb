Option Explicit On

Public Class EmprAlocCargos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850301"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320301"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0301"
#End Region

#Region "Variaveis de Ambiente"
    Private mvaridEmpresa As Short
    Public Property idEmpresa() As Short
        Get
            Return mvaridEmpresa
        End Get
        Set(value As Short)
            mvaridEmpresa = value
        End Set
    End Property


    Private mvaridCargo As Short
    Public Property idCargo() As Short
        Get
            Return mvaridCargo
        End Get
        Set(value As Short)
            mvaridCargo = value
        End Set
    End Property


    Private mvardtInicio As Date
    Public Property dtInicio() As Date
        Get
            Return mvardtInicio
        End Get
        Set(ByVal value As Date)
            mvardtInicio = value
        End Set
    End Property

    Private mvardtFim As Date
    Public Property dtFim() As Date
        Get
            Return mvardtFim
        End Get
        Set(ByVal value As Date)
            mvardtFim = value
        End Set
    End Property

#End Region

#Region "Conex�o com banco"
    Public Db As ADODB.Connection
    Public tEmprAlocCargos As ADODB.Recordset
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
        tEmprAlocCargos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tEmprAlocCargos.Open("EmprAlocCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprAlocCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocCargos.Close
                If tipo = 1 Then
                    tEmprAlocCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprAlocCargos.Open("EmprAlocCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                'dbConecta = Err
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprAlocCargos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprAlocCargos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprAlocCargos.MoveNext    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprAlocCargos.EOF Then
            idEmpresa = tEmprAlocCargos.Fields("idEmpresa").Value
            idCargo = tEmprAlocCargos.Fields("idCargo").Value
            dtInicio = tEmprAlocCargos.Fields("dtInicio").Value
            dtFim = tEmprAlocCargos.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprAlocCargos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, cargo As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprAlocCargos.                          *
        '* *                                      *
        '* * empresa + cargo = campos chave       *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprAlocCargos WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idCargo = " & cargo
        tEmprAlocCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprAlocCargos.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprAlocCargos.Fields("idEmpresa").Value
            idCargo = tEmprAlocCargos.Fields("idCargo").Value
            dtInicio = tEmprAlocCargos.Fields("dtInicio").Value
            dtFim = tEmprAlocCargos.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocCargos.Close
                tEmprAlocCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprAlocCargos - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tEmprAlocCargos.AddNew
        tEmprAlocCargos(0).Value = idEmpresa
        tEmprAlocCargos(1).Value = idCargo
        tEmprAlocCargos(2).Value = dtInicio
        tEmprAlocCargos(3).Value = dtFim
        tEmprAlocCargos.Update
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprAlocCargos - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tEmprAlocCargos(2).Value = dtInicio
        tEmprAlocCargos(3).Value = dtFim
        tEmprAlocCargos.Update
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprAlocCargos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(empresa As Short, cargo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprAlocCargos.                      *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * cargo      = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim I As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim PesqReg1 As GeraisD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro no acesso à tabela segMensagens." & Err.Number & "Processamento interrompido.")
            elimina = 10
            Exit Function
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais
        PesqReg1 = New GeraisD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em OcorAumProm:
        vSelec = "SELECT * FROM OcorAumProm WHERE idCargo = " & cargo
        RC = PesqReg.pesqRegistros(vSelec)
        'RC = 1 - Encontrou registros
        'RC = 2 - Ocorreu algum erro na abertura da tabela
        If RC = 1 Or RC = 2 Then
            If RC = 2 Then
                segMensagens.exibeMsg(, 1069)
                elimina = 2
                Exit Function
            End If
            'Para cada registro encontrado,
            'verifica se o cargo é da mesma Empresa:
            'Lê tabela de OcorAumProm e busca o valor de idColaborador:
            RC = PesqReg.leSeq(1)
            Do While RC = 0
                'leSeq() tem um significado um pouco diferentes do que
                'nos métodos das outras classes.
                'Retorna apenas o valor de idEmpresa.
                'Busca o valor da Empresa, a partir de idColaborador,
                'na tabela OcorDeptoColab:
                vSelec = "SELECT * FROM OcorDeptoColab WHERE idColaborador = " & PesqReg.idColaborador
                I = PesqReg1.pesqRegistros1(vSelec)
                'I = 1 - Encontrou registros
                'I = 2 - Ocorreu algum erro na abertura da tabela
                If I = 2 Then
                    segMensagens.exibeMsg(, 1069)
                    elimina = 2
                    Exit Function
                End If
                I = PesqReg1.leSeq1(1)
                'Testa retorno da leitura da tabela. Registro tem que existir:
                If I <> 0 Then
                    If RC <> 1016 Then
                        elimina = RC
                        Exit Function
                    End If
                Else
                    'Encontrou o registro; busca código da Empresa:
                    If PesqReg1.idEmpresa = empresa Then
                        'Violação de integridade referencial:
                        segMensagens.exibeMsg(, 1067)
                        elimina = 1067
                        Exit Function
                    End If
                End If
                RC = PesqReg.leSeq(0)
            Loop
            'Testa retorno da leitura da tabela:
            If RC <> 1016 Then
                elimina = RC
                Exit Function
            End If
        End If

        'Só permite a eliminação de cargos ativos:
        vSelec = "SELECT * FROM EmprAlocCargos WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idCargo = " & cargo
        vSelec = vSelec & " AND Convert(datetime, dtFim, 112) <> '18991230'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1349)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM EmprAlocCargos WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idCargo = " & cargo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprAlocCargos - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function




End Class
