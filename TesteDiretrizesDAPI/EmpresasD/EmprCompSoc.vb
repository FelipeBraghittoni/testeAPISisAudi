Option Explicit On
Public Class EmprCompSoc

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850307"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320307"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0307"
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


    Private mvartpSocio As Short
    Public Property tpSocio() As Short
        Get
            Return mvartpSocio
        End Get
        Set(value As Short)
            mvartpSocio = value
        End Set
    End Property


    Private mvaridSocio As Single
    Public Property idSocio() As Single
        Get
            Return mvaridSocio
        End Get
        Set(value As Single)
            mvaridSocio = value
        End Set
    End Property


    Private mvarnomeSocio As String
    Public Property nomeSocio() As String
        Get
            Return mvarnomeSocio
        End Get
        Set(value As String)
            mvarnomeSocio = value
        End Set
    End Property


    Private mvarenderecoSocio As String
    Public Property enderecoSocio() As String
        Get
            Return mvarenderecoSocio
        End Get
        Set(value As String)
            mvarenderecoSocio = value
        End Set
    End Property


    Private mvarcomplemSocio As String
    Public Property complemSocio() As String
        Get
            Return mvarcomplemSocio
        End Get
        Set(value As String)
            mvarcomplemSocio = value
        End Set
    End Property


    Private mvarbairroSocio As String
    Public Property bairroSocio() As String
        Get
            Return mvarbairroSocio
        End Get
        Set(value As String)
            mvarbairroSocio = value
        End Set
    End Property


    Private mvarpaisSocio As Short
    Public Property paisSocio() As Short
        Get
            Return mvarbairroSocio
        End Get
        Set(value As Short)
            mvarbairroSocio = value
        End Set
    End Property


    Private mvarestadoSocio As String
    Public Property estadoSocio() As String
        Get
            Return mvarestadoSocio
        End Get
        Set(value As String)
            mvarestadoSocio = value
        End Set
    End Property


    Private mvarcidadeSocio As Short
    Public Property cidadeSocio() As Short
        Get
            Return mvarcidadeSocio
        End Get
        Set(value As Short)
            mvarcidadeSocio = value
        End Set
    End Property


    Private mvarcepSocio As Integer
    Public Property cepSocio() As Integer
        Get
            Return mvarcepSocio
        End Get
        Set(value As Integer)
            mvarcepSocio = value
        End Set
    End Property


    Private mvareMailSocio As String
    Public Property eMailSocio() As String
        Get
            Return mvareMailSocio
        End Get
        Set(value As String)
            mvareMailSocio = value
        End Set
    End Property


    Private mvarfoneSocio As Integer
    Public Property foneSocio() As Integer
        Get
            Return mvarfoneSocio
        End Get
        Set(value As Integer)
            mvarfoneSocio = value
        End Set
    End Property


    Private mvarcelSocio As Integer
    Public Property celSocio() As Integer
        Get
            Return mvarcelSocio
        End Get
        Set(value As Integer)
            mvarcelSocio = value
        End Set
    End Property


    Private mvarCPFSocio As Double
    Public Property CPFSocio() As Double
        Get
            Return mvarCPFSocio
        End Get
        Set(value As Double)
            mvarCPFSocio = value
        End Set
    End Property


    Private mvarRGSocio As String
    Public Property RGSocio() As String
        Get
            Return mvarCPFSocio
        End Get
        Set(value As String)
            mvarCPFSocio = value
        End Set
    End Property


    Private mvaremissorRGSocio As String
    Public Property emissorRGSocio() As String
        Get
            Return mvaremissorRGSocio
        End Get
        Set(value As String)
            mvaremissorRGSocio = value
        End Set
    End Property


    Private mvarCNPJSocio As Double
    Public Property CNPJSocio() As Double
        Get
            Return mvarCNPJSocio
        End Get
        Set(value As Double)
            mvarCNPJSocio = value
        End Set
    End Property


    Private mvarRaizCNPJSocio As String
    Public Property RaizCNPJSocio() As String
        Get
            Return mvarCNPJSocio
        End Get
        Set(value As String)
            mvarCNPJSocio = value
        End Set
    End Property


    Private mvarDigCNPJSocio As String
    Public Property DigCNPJSocio() As String
        Get
            Return mvarDigCNPJSocio
        End Get
        Set(value As String)
            mvarDigCNPJSocio = value
        End Set
    End Property


    Private mvartpParticip As Short
    Public Property tpParticip() As Short
        Get
            Return mvartpParticip
        End Get
        Set(value As Short)
            mvartpParticip = value
        End Set
    End Property


    Private mvarparticipacao As Double
    Public Property participacao() As Double
        Get
            Return mvarparticipacao
        End Get
        Set(value As Double)
            mvarparticipacao = value
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
    Public tEmprCompSoc As ADODB.Recordset

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
        tEmprCompSoc = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            tEmprCompSoc.Open("EmprCompSoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprCompSoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprCompSoc.Close()

                If tipo = 1 Then
                    tEmprCompSoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprCompSoc.Open("EmprCompSoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprCompSoc - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprCompSoc.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprCompSoc.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprCompSoc.EOF Then
            idEmpresa = tEmprCompSoc.Fields("idEmpresa").Value
            tpSocio = tEmprCompSoc.Fields("tpSocio").Value
            idSocio = tEmprCompSoc.Fields("idSocio").Value
            nomeSocio = tEmprCompSoc.Fields("nomeSocio").Value
            enderecoSocio = tEmprCompSoc.Fields("enderecoSocio").Value
            complemSocio = tEmprCompSoc.Fields("complemSocio").Value
            bairroSocio = tEmprCompSoc.Fields("bairroSocio").Value
            paisSocio = tEmprCompSoc.Fields("paisSocio").Value
            estadoSocio = tEmprCompSoc.Fields("estadoSocio").Value
            cidadeSocio = tEmprCompSoc.Fields("cidadeSocio").Value
            cepSocio = tEmprCompSoc.Fields("cepSocio").Value
            eMailSocio = tEmprCompSoc.Fields("eMailSocio").Value
            foneSocio = tEmprCompSoc.Fields("foneSocio").Value
            celSocio = tEmprCompSoc.Fields("celSocio").Value
            CPFSocio = tEmprCompSoc.Fields("CPFSocio").Value
            RGSocio = tEmprCompSoc.Fields("RGSocio").Value
            emissorRGSocio = tEmprCompSoc.Fields("emissorRGSocio").Value
            CNPJSocio = tEmprCompSoc.Fields("CNPJSocio").Value
            RaizCNPJSocio = tEmprCompSoc.Fields("RaizCNPJSocio").Value
            DigCNPJSocio = tEmprCompSoc.Fields("DigCNPJSocio").Value
            tpParticip = tEmprCompSoc.Fields("tpParticip").Value
            participacao = tEmprCompSoc.Fields("participacao").Value
            idMoeda = tEmprCompSoc.Fields("idMoeda").Value
            dtInicio = tEmprCompSoc.Fields("dtInicio").Value
            dtFim = tEmprCompSoc.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprCompSoc - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, nome As String, data As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprCompSoc.                         *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * nome       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprCompSoc WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND nomeSocio = '" & nome & "'"
        vSelect = vSelect & " AND Convert(datetime, dtInicio, 112) = '" & Format(data, "yyyymmdd") & "'"
        tEmprCompSoc.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprCompSoc.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprCompSoc.Fields("idEmpresa").Value
            tpSocio = tEmprCompSoc.Fields("tpSocio").Value
            idSocio = tEmprCompSoc.Fields("idSocio").Value
            nomeSocio = tEmprCompSoc.Fields("nomeSocio").Value
            enderecoSocio = tEmprCompSoc.Fields("enderecoSocio").Value
            complemSocio = tEmprCompSoc.Fields("complemSocio").Value
            bairroSocio = tEmprCompSoc.Fields("bairroSocio").Value
            paisSocio = tEmprCompSoc.Fields("paisSocio").Value
            estadoSocio = tEmprCompSoc.Fields("estadoSocio").Value
            cidadeSocio = tEmprCompSoc.Fields("cidadeSocio").Value
            cepSocio = tEmprCompSoc.Fields("cepSocio").Value
            eMailSocio = tEmprCompSoc.Fields("eMailSocio").Value
            foneSocio = tEmprCompSoc.Fields("foneSocio").Value
            celSocio = tEmprCompSoc.Fields("celSocio").Value
            CPFSocio = tEmprCompSoc.Fields("CPFSocio").Value
            RGSocio = tEmprCompSoc.Fields("RGSocio").Value
            emissorRGSocio = tEmprCompSoc.Fields("emissorRGSocio").Value
            CNPJSocio = tEmprCompSoc.Fields("CNPJSocio").Value
            RaizCNPJSocio = tEmprCompSoc.Fields("RaizCNPJSocio").Value
            DigCNPJSocio = tEmprCompSoc.Fields("DigCNPJSocio").Value
            tpParticip = tEmprCompSoc.Fields("tpParticip").Value
            participacao = tEmprCompSoc.Fields("participacao").Value
            idMoeda = tEmprCompSoc.Fields("idMoeda").Value
            dtInicio = tEmprCompSoc.Fields("dtInicio").Value
            dtFim = tEmprCompSoc.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprCompSoc.Close()
                tEmprCompSoc.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprCompSoc - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(vDescric As String) As Integer

        On Error GoTo Einclui
        tEmprCompSoc.AddNew()
        tEmprCompSoc.Fields("idEmpresa").Value = idEmpresa
        tEmprCompSoc.Fields("tpSocio").Value = tpSocio
        tEmprCompSoc.Fields("idSocio").Value = idSocio
        tEmprCompSoc.Fields("nomeSocio").Value = nomeSocio
        tEmprCompSoc.Fields("enderecoSocio").Value = enderecoSocio
        tEmprCompSoc.Fields("complemSocio").Value = complemSocio
        tEmprCompSoc.Fields("bairroSocio").Value = bairroSocio
        tEmprCompSoc.Fields("paisSocio").Value = paisSocio
        tEmprCompSoc.Fields("estadoSocio").Value = estadoSocio
        tEmprCompSoc.Fields("cidadeSocio").Value = cidadeSocio
        tEmprCompSoc.Fields("cepSocio").Value = cepSocio
        tEmprCompSoc.Fields("eMailSocio").Value = eMailSocio
        tEmprCompSoc.Fields("foneSocio").Value = foneSocio
        tEmprCompSoc.Fields("celSocio").Value = celSocio
        tEmprCompSoc.Fields("CPFSocio").Value = CPFSocio
        tEmprCompSoc.Fields("RGSocio").Value = RGSocio
        tEmprCompSoc.Fields("emissorRGSocio").Value = emissorRGSocio
        tEmprCompSoc.Fields("CNPJSocio").Value = CNPJSocio
        tEmprCompSoc.Fields("RaizCNPJSocio").Value = RaizCNPJSocio
        tEmprCompSoc.Fields("DigCNPJSocio").Value = DigCNPJSocio
        tEmprCompSoc.Fields("tpParticip").Value = tpParticip
        tEmprCompSoc.Fields("participacao").Value = participacao
        tEmprCompSoc.Fields("idMoeda").Value = idMoeda
        tEmprCompSoc.Fields("dtInicio").Value = dtInicio
        tEmprCompSoc.Fields("dtFim").Value = dtFim
        tEmprCompSoc.Fields("descric").Value = IIf(vDescric = "", " ", vDescric)
        tEmprCompSoc.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprCompSoc - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(vDescric As String) As Integer

        On Error GoTo Ealtera
        tEmprCompSoc.AddNew()
        tEmprCompSoc.Fields("idEmpresa").Value = idEmpresa
        tEmprCompSoc.Fields("tpSocio").Value = tpSocio
        tEmprCompSoc.Fields("idSocio").Value = idSocio
        tEmprCompSoc.Fields("nomeSocio").Value = nomeSocio
        tEmprCompSoc.Fields("enderecoSocio").Value = enderecoSocio
        tEmprCompSoc.Fields("complemSocio").Value = complemSocio
        tEmprCompSoc.Fields("bairroSocio").Value = bairroSocio
        tEmprCompSoc.Fields("paisSocio").Value = paisSocio
        tEmprCompSoc.Fields("estadoSocio").Value = estadoSocio
        tEmprCompSoc.Fields("cidadeSocio").Value = cidadeSocio
        tEmprCompSoc.Fields("cepSocio").Value = cepSocio
        tEmprCompSoc.Fields("eMailSocio").Value = eMailSocio
        tEmprCompSoc.Fields("foneSocio").Value = foneSocio
        tEmprCompSoc.Fields("celSocio").Value = celSocio
        tEmprCompSoc.Fields("CPFSocio").Value = CPFSocio
        tEmprCompSoc.Fields("RGSocio").Value = RGSocio
        tEmprCompSoc.Fields("emissorRGSocio").Value = emissorRGSocio
        tEmprCompSoc.Fields("CNPJSocio").Value = CNPJSocio
        tEmprCompSoc.Fields("RaizCNPJSocio").Value = RaizCNPJSocio
        tEmprCompSoc.Fields("DigCNPJSocio").Value = DigCNPJSocio
        tEmprCompSoc.Fields("tpParticip").Value = tpParticip
        tEmprCompSoc.Fields("participacao").Value = participacao
        tEmprCompSoc.Fields("idMoeda").Value = idMoeda
        tEmprCompSoc.Fields("dtInicio").Value = dtInicio
        tEmprCompSoc.Fields("dtFim").Value = dtFim
        tEmprCompSoc.Fields("descric").Value = IIf(vDescric = "", " ", vDescric)
        tEmprCompSoc.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprCompSoc - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function extraiMemo() As String
        'Extrai o conteúdo do campo memo txtDescric

        extraiMemo = tEmprCompSoc.Fields("descric").Value

    End Function

    Public Function elimina(empresa As Short, nome As String, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprCompSoc.                         *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * nome       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprCompSoc WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND nomeSocio = '" & nome & "'"
        vSelec = vSelec & " AND Convert(datetime, dtInicio, 112) = '" & Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprCompSoc - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
