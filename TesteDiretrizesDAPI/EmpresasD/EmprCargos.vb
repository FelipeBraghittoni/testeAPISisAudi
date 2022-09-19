Option Explicit On
Public Class EmprCargos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850306"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320306"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0306"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridCargo As Short
    Public Property idCargo() As Short
        Get
            Return mvaridCargo
        End Get
        Set(value As Short)
            mvaridCargo = value
        End Set
    End Property

    Private mvarnomeCargo As String
    Public Property nomeCargo() As Short
        Get
            Return mvarnomeCargo
        End Get
        Set(value As Short)
            mvarnomeCargo = value
        End Set
    End Property

    Private mvarconfiancaCargo As String
    Public Property confiancaCargo() As Short
        Get
            Return mvarconfiancaCargo
        End Get
        Set(value As Short)
            mvarconfiancaCargo = value
        End Set
    End Property


    Private mvarpermanMinCargo As Short
    Public Property permanMinCargo() As Short
        Get
            Return mvarpermanMinCargo
        End Get
        Set(value As Short)
            mvarpermanMinCargo = value
        End Set
    End Property


    Private mvarpermanMaxCargo As Short
    Public Property permanMaxCargo() As Short
        Get
            Return mvarpermanMaxCargo
        End Get
        Set(value As Short)
            mvarpermanMaxCargo = value
        End Set
    End Property


    Private mvarsalarioCargoMin As Double
    Public Property salarioCargoMin() As Double
        Get
            Return mvarsalarioCargoMin
        End Get
        Set(value As Double)
            mvarsalarioCargoMin = value
        End Set
    End Property


    Private mvarsalarioCargo1 As Double
    Public Property salarioCargo1() As Double
        Get
            Return mvarsalarioCargo1
        End Get
        Set(value As Double)
            mvarsalarioCargo1 = value
        End Set
    End Property


    Private mvarsalarioCargoMed As Double
    Public Property salarioCargoMed() As Double
        Get
            Return mvarsalarioCargoMed
        End Get
        Set(value As Double)
            mvarsalarioCargoMed = value
        End Set
    End Property


    Private mvarsalarioCargo3 As Double
    Public Property salarioCargo3() As Double
        Get
            Return mvarsalarioCargo3
        End Get
        Set(value As Double)
            mvarsalarioCargo3 = value
        End Set
    End Property


    Private mvarsalarioCargoMax As Double
    Public Property salarioCargoMax() As Double
        Get
            Return mvarsalarioCargoMax
        End Get
        Set(value As Double)
            mvarsalarioCargoMax = value
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
    Public tEmprCargos As ADODB.Recordset

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
        tEmprCargos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tEmprCargos.Open("EmprCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM EmprCargos ORDER BY nomeCargo"
                tEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprCargos.Close()

                If tipo = 1 Or tipo = 2 Then
                    tEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprCargos.Open("EmprCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                'dbConecta = Err
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprCargos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprCargos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprCargos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprCargos.EOF Then
            idCargo = tEmprCargos.Fields("idCargo").Value
            nomeCargo = tEmprCargos.Fields("nomeCargo").Value
            confiancaCargo = tEmprCargos.Fields("confiancaCargo").Value
            permanMinCargo = tEmprCargos.Fields("permanMinCargo").Value
            permanMaxCargo = tEmprCargos.Fields("permanMaxCargo").Value
            salarioCargoMin = tEmprCargos.Fields("salarioCargoMin").Value
            salarioCargo1 = tEmprCargos.Fields("salarioCargo1").Value
            salarioCargoMed = tEmprCargos.Fields("salarioCargoMed").Value
            salarioCargo3 = tEmprCargos.Fields("salarioCargo3").Value
            salarioCargoMax = tEmprCargos.Fields("salarioCargoMax").Value
            idMoeda = tEmprCargos.Fields("idMoeda").Value
            dtInicio = tEmprCargos.Fields("dtInicio").Value
            dtFim = tEmprCargos.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprCargos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(cargo As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprCargos.                          *
        '* *                                      *
        '* * cargo      =      campos chave       *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprCargos WHERE idCargo = " & cargo
        tEmprCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprCargos.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idCargo = tEmprCargos.Fields("idCargo").Value
            nomeCargo = tEmprCargos.Fields("nomeCargo").Value
            confiancaCargo = tEmprCargos.Fields("confiancaCargo").Value
            permanMinCargo = tEmprCargos.Fields("permanMinCargo").Value
            permanMaxCargo = tEmprCargos.Fields("permanMaxCargo").Value
            salarioCargoMin = tEmprCargos.Fields("salarioCargoMin").Value
            salarioCargo1 = tEmprCargos.Fields("salarioCargo1").Value
            salarioCargoMed = tEmprCargos.Fields("salarioCargoMed").Value
            salarioCargo3 = tEmprCargos.Fields("salarioCargo3").Value
            salarioCargoMax = tEmprCargos.Fields("salarioCargoMax").Value
            idMoeda = tEmprCargos.Fields("idMoeda").Value
            dtInicio = tEmprCargos.Fields("dtInicio").Value
            dtFim = tEmprCargos.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprCargos.Close()
                tEmprCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprCargos - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprCargos.                          *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM EmprCargos WHERE nomeCargo = '" & descric & "'"
        tEmprCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprCargos.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idCargo = tEmprCargos.Fields("idCargo").Value
            nomeCargo = tEmprCargos.Fields("nomeCargo").Value
            confiancaCargo = tEmprCargos.Fields("confiancaCargo").Value
            permanMinCargo = tEmprCargos.Fields("permanMinCargo").Value
            permanMaxCargo = tEmprCargos.Fields("permanMaxCargo").Value
            salarioCargoMin = tEmprCargos.Fields("salarioCargoMin").Value
            salarioCargo1 = tEmprCargos.Fields("salarioCargo1").Value
            salarioCargoMed = tEmprCargos.Fields("salarioCargoMed").Value
            salarioCargo3 = tEmprCargos.Fields("salarioCargo3").Value
            salarioCargoMax = tEmprCargos.Fields("salarioCargoMax").Value
            idMoeda = tEmprCargos.Fields("idMoeda").Value
            dtInicio = tEmprCargos.Fields("dtInicio").Value
            dtFim = tEmprCargos.Fields("dtFim").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprCargos.Close()
                tEmprCargos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe EmprCargos - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(preReq As String, responsab As String, tarefas As String) As Integer

        On Error GoTo Einclui
        tEmprCargos.AddNew()
        tEmprCargos(0).Value = idCargo
        tEmprCargos(1).Value = nomeCargo
        tEmprCargos(2).Value = IIf(preReq = "", " ", preReq)
        tEmprCargos(3).Value = confiancaCargo
        tEmprCargos(4).Value = IIf(responsab = "", " ", responsab)
        tEmprCargos(5).Value = IIf(tarefas = "", " ", tarefas)
        tEmprCargos(6).Value = permanMinCargo
        tEmprCargos(7).Value = permanMaxCargo
        tEmprCargos(8).Value = salarioCargoMin
        tEmprCargos(9).Value = salarioCargo1
        tEmprCargos(10).Value = salarioCargoMed
        tEmprCargos(11).Value = salarioCargo3
        tEmprCargos(12).Value = salarioCargoMax
        tEmprCargos(13).Value = idMoeda
        tEmprCargos.Fields("dtInicio").Value = dtInicio
        tEmprCargos.Fields("dtFim").Value = dtFim
        tEmprCargos.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprCargos - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function altera(preReq As String, responsab As String, tarefas As String) As Integer

        On Error GoTo Ealtera
        tEmprCargos(1).Value = nomeCargo
        tEmprCargos(2).Value = IIf(preReq = "", " ", preReq)
        tEmprCargos(3).Value = confiancaCargo
        tEmprCargos(4).Value = IIf(responsab = "", " ", responsab)
        tEmprCargos(5).Value = IIf(tarefas = "", " ", tarefas)
        tEmprCargos(6).Value = permanMinCargo
        tEmprCargos(7).Value = permanMaxCargo
        tEmprCargos(8).Value = salarioCargoMin
        tEmprCargos(9).Value = salarioCargo1
        tEmprCargos(10).Value = salarioCargoMed
        tEmprCargos(11).Value = salarioCargo3
        tEmprCargos(12).Value = salarioCargoMax
        tEmprCargos(13).Value = idMoeda
        tEmprCargos.Fields("dtInicio").Value = dtInicio
        tEmprCargos.Fields("dtFim").Value = dtFim
        tEmprCargos.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprCargos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaPreReq() As String

        CarregaPreReq = tEmprCargos.Fields("preRequisitosCargo").Value

    End Function

    Public Function CarregaResponsab() As String

        CarregaResponsab = tEmprCargos.Fields("responsabCargo").Value

    End Function

    Public Function CarregaTarefas() As String

        CarregaTarefas = tEmprCargos.Fields("tarefasCargo").Value

    End Function

    Public Function elimina(cargos As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprCargos.                          *
        '* *                                      *
        '* * cargos     = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        RC = fPesquisa(0, cargos)

        If RC = 0 Then
            vSelec = "DELETE FROM EmprCargos WHERE "
            vSelec = vSelec & " idCargo = " & cargos
            Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            Select Case RC
                Case Is < 3     'ColExtensaoBeneficio
                    msgErro("ColExtensaoBeneficio")
                Case Is < 5     'OcorAumProm
                    msgErro("OcorAumProm")
                Case Is < 7     'OcorHAumProm
                    msgErro("OcorHAumProm")
                Case Is < 9     'EmprAlocCargos
                    msgErro("EmprAlocCargos")
                Case Is < 11    'ProjAlocAtiv
                    msgErro("ProjAlocAtiv")
            End Select
        End If

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprCargos - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function fPesquisa(tipo As Short, cargos As Short, Optional data As Date = Nothing) As Integer
        '* *********************************************
        '* tipo = 0 ==> não pesquisa data de término   *
        '*              utilizado para eliminar linha  *
        '* tipo = 1 ==> pesquisa data de término       *
        '*              utilizado para encerrar depto. *
        '* *********************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro ao abrir segMensagens")
            fPesquisa = -1
            Exit Function
        End If

        fPesquisa = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em ColExtensaoBeneficio:
        vSelec = "SELECT * FROM ColExtensaoBeneficio WHERE idCargoColab = " & cargos
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC
            Exit Function
        End If
        'Pesquisa em OcorAumProm:
        vSelec = "SELECT * FROM OcorAumProm WHERE idCargo = " & cargos
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalAumProm, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalAumProm, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 2
            Exit Function
        End If
        'Pesquisa em OcorHAumProm:
        vSelec = "SELECT * FROM OcorHAumProm WHERE idCargo = " & cargos
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalAumProm, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalAumProm, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 4
            Exit Function
        End If
        'Pesquisa em EmprAlocCargos:
        vSelec = "SELECT * FROM EmprAlocCargos WHERE idCargo = " & cargos
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 6
            Exit Function
        End If
        'Pesquisa em ProjAlocAtiv:
        vSelec = "SELECT * FROM ProjAlocAtiv WHERE idCargo = " & cargos
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 8
            Exit Function
        End If

    End Function

    Private Sub msgErro(tabela As String)

        MsgBox("Há um cargo na tabela " & tabela & " que impede sua eliminação.")

    End Sub

End Class
