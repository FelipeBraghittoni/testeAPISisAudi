Option Explicit On
Public Class EmprSites
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850314"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320314"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0314"
#End Region

#Region "Variaveis de ambiente"

    Private mvaridEmpresa As Short
    Public Property idEmpresa() As Short
        Get
            Return mvaridEmpresa
        End Get
        Set(value As Short)
            mvaridEmpresa = value
        End Set
    End Property


    Private mvaridSite As Short
    Public Property idSite() As Short
        Get
            Return mvaridSite
        End Get
        Set(value As Short)
            mvaridSite = value
        End Set
    End Property


    Private mvarnomeSite As String
    Public Property nomeSite() As String
        Get
            Return mvarnomeSite
        End Get
        Set(value As String)
            mvarnomeSite = value
        End Set
    End Property


    Private mvarfone1Site As Integer
    Public Property fone1site() As Integer
        Get
            Return mvarfone1Site
        End Get
        Set(value As Integer)
            mvarfone1Site = value
        End Set
    End Property


    Private mvarramal1Site As Short
    Public Property ramal1Site() As Short
        Get
            Return mvarramal1Site
        End Get
        Set(value As Short)
            mvarramal1Site = value
        End Set
    End Property


    Private mvarfone2Site As Integer
    Public Property fone2Site() As Integer
        Get
            Return mvarfone2Site
        End Get
        Set(value As Integer)
            mvarfone2Site = value
        End Set
    End Property


    Private mvarramal2Site As Short
    Public Property ramal2Site() As Short
        Get
            Return mvarramal2Site
        End Get
        Set(value As Short)
            mvarramal2Site = value
        End Set
    End Property


    Private mvarfaxSite As Integer
    Public Property faxSite() As Integer
        Get
            Return mvarfaxSite
        End Get
        Set(value As Integer)
            mvarfaxSite = value
        End Set
    End Property


    Private mvarenderecoSite As String
    Public Property enderecoSite() As String
        Get
            Return mvarenderecoSite
        End Get
        Set(value As String)
            mvarenderecoSite = value
        End Set
    End Property


    Private mvarcomplemSite As String
    Public Property complemSite() As String
        Get
            Return mvarcomplemSite
        End Get
        Set(value As String)
            mvarcomplemSite = value
        End Set
    End Property


    Private mvarbairroSite As String
    Public Property bairroSite() As String
        Get
            Return mvarbairroSite
        End Get
        Set(value As String)
            mvarbairroSite = value
        End Set
    End Property


    Private mvarpaisSite As Short
    Public Property paisSite() As Short
        Get
            Return mvarpaisSite
        End Get
        Set(value As Short)
            mvarpaisSite = value
        End Set
    End Property


    Private mvarufSite As String
    Public Property ufSite() As String
        Get
            Return mvarufSite
        End Get
        Set(value As String)
            mvarufSite = value
        End Set
    End Property


    Private mvarcidadeSite As Short
    Public Property cidadeSite() As Short
        Get
            Return mvarcidadeSite
        End Get
        Set(value As Short)
            mvarcidadeSite = value
        End Set
    End Property


    Private mvarcepSite As Integer
    Public Property cepSite() As Integer
        Get
            Return mvarcepSite
        End Get
        Set(value As Integer)
            mvarcepSite = value
        End Set
    End Property


    Private mvarregimeTrabSite As Short
    Public Property regimeTrabSite() As Short
        Get
            Return mvarregimeTrabSite
        End Get
        Set(value As Short)
            mvarregimeTrabSite = value
        End Set
    End Property


    Private mvardiaInicTrabSite As Short
    Public Property diaInicTrabSite() As Short
        Get
            Return mvardiaInicTrabSite
        End Get
        Set(value As Short)
            mvardiaInicTrabSite = value
        End Set
    End Property


    Private mvardiaFimTrabSite As Short
    Public Property diaFimTrabSite() As Short
        Get
            Return mvardiaFimTrabSite
        End Get
        Set(value As Short)
            mvardiaFimTrabSite = value
        End Set
    End Property


    Private mvarrefeitorioSite As String
    Public Property refeitorioSite() As String
        Get
            Return mvarrefeitorioSite
        End Get
        Set(value As String)
            mvarrefeitorioSite = value
        End Set
    End Property


    Private mvarhoraInicTrabSite As String
    Public Property horaInicTrabSite() As String
        Get
            Return mvarhoraInicTrabSite
        End Get
        Set(value As String)
            mvarhoraInicTrabSite = value
        End Set
    End Property


    Private mvarhoraInicDescSite As String
    Public Property horaInicDescSite() As String
        Get
            Return mvarhoraInicDescSite
        End Get
        Set(value As String)
            mvarhoraInicDescSite = value
        End Set
    End Property


    Private mvarhoraFimDescSite As String
    Public Property horaFimDescSite() As String
        Get
            Return mvarhoraFimDescSite
        End Get
        Set(value As String)
            mvarhoraFimDescSite = value
        End Set
    End Property


    Private mvarhoraFimTrabSite As String
    Public Property horaFimTrabSite() As String
        Get
            Return mvarhoraFimTrabSite
        End Get
        Set(value As String)
            mvarhoraFimTrabSite = value
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


    Private mvarraizCNPJSite As Double
    Public Property raizCNPJSite() As Double
        Get
            Return mvarraizCNPJSite
        End Get
        Set(value As Double)
            mvarraizCNPJSite = value
        End Set
    End Property


    Private mvarcompCNPJSite As Short
    Public Property compCNPJSite() As Short
        Get
            Return mvarcompCNPJSite
        End Get
        Set(value As Short)
            mvarcompCNPJSite = value
        End Set
    End Property


    Private mvardigCNPJSite As Short
    Public Property digCNPJSite() As Short
        Get
            Return mvardigCNPJSite
        End Get
        Set(value As Short)
            mvardigCNPJSite = value
        End Set
    End Property


    Private mvartipoPessoa As Short
    Public Property tipoPessoa() As Short
        Get
            Return mvartipoPessoa
        End Get
        Set(value As Short)
            mvartipoPessoa = value
        End Set
    End Property


    Private mvarpadrao As String
    Public Property padrao() As String
        Get
            Return mvarpadrao
        End Get
        Set(value As String)
            mvarpadrao = value
        End Set
    End Property


#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public tEmprSites As ADODB.Recordset
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
        tEmprSites = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tEmprSites.Open("EmprSites", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM EmprSites ORDER BY idEmpresa, nomeSite"
                tEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprSites.Close()

                If tipo = 1 Or tipo = 2 Then
                    tEmprSites.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprSites.Open("EmprSites", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprSites - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Integer) As Long
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tEmprSites.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprSites.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprSites.EOF Then
            idEmpresa = tEmprSites.Fields("idEmpresa").Value
            idSite = tEmprSites.Fields("idSite").Value
            nomeSite = tEmprSites.Fields("nomeSite").Value
            fone1site = tEmprSites.Fields("fone1Site").Value
            ramal1Site = tEmprSites.Fields("ramal1Site").Value
            fone2Site = tEmprSites.Fields("fone2Site").Value
            ramal2Site = tEmprSites.Fields("ramal2Site").Value
            faxSite = tEmprSites.Fields("faxSite").Value
            enderecoSite = tEmprSites.Fields("enderecoSite").Value
            complemSite = tEmprSites.Fields("complemSite").Value
            bairroSite = tEmprSites.Fields("bairroSite").Value
            paisSite = tEmprSites.Fields("paisSite").Value
            ufSite = tEmprSites.Fields("ufSite").Value
            cidadeSite = tEmprSites.Fields("cidadeSite").Value
            cepSite = tEmprSites.Fields("cepSite").Value
            regimeTrabSite = tEmprSites.Fields("regimeTrabSite").Value
            diaInicTrabSite = tEmprSites.Fields("diaInicTrabSite").Value
            diaFimTrabSite = tEmprSites.Fields("diaFimTrabSite").Value
            refeitorioSite = tEmprSites.Fields("refeitorioSite").Value
            horaInicTrabSite = tEmprSites.Fields("horaInicTrabSite").Value
            horaInicDescSite = tEmprSites.Fields("horaInicDescSite").Value
            horaFimDescSite = tEmprSites.Fields("horaFimDescSite").Value
            horaFimTrabSite = tEmprSites.Fields("horaFimTrabSite").Value
            dtInicio = tEmprSites.Fields("dtInicio").Value
            dtFim = tEmprSites.Fields("dtFim").Value
            raizCNPJSite = tEmprSites.Fields("raizCNPJSite").Value
            compCNPJSite = tEmprSites.Fields("compCNPJSite").Value
            digCNPJSite = tEmprSites.Fields("digCNPJSite").Value
            tipoPessoa = tEmprSites.Fields("tipoPessoa").Value
            padrao = tEmprSites.Fields("padrao").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprSites - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, site As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprSites.                           *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * site       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprSites WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idSite = " & site
        tEmprSites.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprSites.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprSites.Fields("idEmpresa").Value
            idSite = tEmprSites.Fields("idSite").Value
            nomeSite = tEmprSites.Fields("nomeSite").Value
            fone1site = tEmprSites.Fields("fone1Site").Value
            ramal1Site = tEmprSites.Fields("ramal1Site").Value
            fone2Site = tEmprSites.Fields("fone2Site").Value
            ramal2Site = tEmprSites.Fields("ramal2Site").Value
            faxSite = tEmprSites.Fields("faxSite").Value
            enderecoSite = tEmprSites.Fields("enderecoSite").Value
            complemSite = tEmprSites.Fields("complemSite").Value
            bairroSite = tEmprSites.Fields("bairroSite").Value
            paisSite = tEmprSites.Fields("paisSite").Value
            ufSite = tEmprSites.Fields("ufSite").Value
            cidadeSite = tEmprSites.Fields("cidadeSite").Value
            cepSite = tEmprSites.Fields("cepSite").Value
            regimeTrabSite = tEmprSites.Fields("regimeTrabSite").Value
            diaInicTrabSite = tEmprSites.Fields("diaInicTrabSite").Value
            diaFimTrabSite = tEmprSites.Fields("diaFimTrabSite").Value
            refeitorioSite = tEmprSites.Fields("refeitorioSite").Value
            horaInicTrabSite = tEmprSites.Fields("horaInicTrabSite").Value
            horaInicDescSite = tEmprSites.Fields("horaInicDescSite").Value
            horaFimDescSite = tEmprSites.Fields("horaFimDescSite").Value
            horaFimTrabSite = tEmprSites.Fields("horaFimTrabSite").Value
            dtInicio = tEmprSites.Fields("dtInicio").Value
            dtFim = tEmprSites.Fields("dtFim").Value
            raizCNPJSite = tEmprSites.Fields("raizCNPJSite").Value
            compCNPJSite = tEmprSites.Fields("compCNPJSite").Value
            digCNPJSite = tEmprSites.Fields("digCNPJSite").Value
            tipoPessoa = tEmprSites.Fields("tipoPessoa").Value
            padrao = tEmprSites.Fields("padrao").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprSites.Close()
                tEmprSites.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprSites - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(empresa As Short, descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprSites.                           *
        '* *                                      *
        '* * empresa = argum. p/pesquisa +        *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM EmprSites WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND nomeSite = '" & descric & "'"
        tEmprSites.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprSites.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprSites.Fields("idEmpresa").Value
            idSite = tEmprSites.Fields("idSite").Value
            nomeSite = tEmprSites.Fields("nomeSite").Value
            fone1site = tEmprSites.Fields("fone1Site").Value
            ramal1Site = tEmprSites.Fields("ramal1Site").Value
            fone2Site = tEmprSites.Fields("fone2Site").Value
            ramal2Site = tEmprSites.Fields("ramal2Site").Value
            faxSite = tEmprSites.Fields("faxSite").Value
            enderecoSite = tEmprSites.Fields("enderecoSite").Value
            complemSite = tEmprSites.Fields("complemSite").Value
            bairroSite = tEmprSites.Fields("bairroSite").Value
            paisSite = tEmprSites.Fields("paisSite").Value
            ufSite = tEmprSites.Fields("ufSite").Value
            cidadeSite = tEmprSites.Fields("cidadeSite").Value
            cepSite = tEmprSites.Fields("cepSite").Value
            regimeTrabSite = tEmprSites.Fields("regimeTrabSite").Value
            diaInicTrabSite = tEmprSites.Fields("diaInicTrabSite").Value
            diaFimTrabSite = tEmprSites.Fields("diaFimTrabSite").Value
            refeitorioSite = tEmprSites.Fields("refeitorioSite").Value
            horaInicTrabSite = tEmprSites.Fields("horaInicTrabSite").Value
            horaInicDescSite = tEmprSites.Fields("horaInicDescSite").Value
            horaFimDescSite = tEmprSites.Fields("horaFimDescSite").Value
            horaFimTrabSite = tEmprSites.Fields("horaFimTrabSite").Value
            dtInicio = tEmprSites.Fields("dtInicio").Value
            dtFim = tEmprSites.Fields("dtFim").Value
            raizCNPJSite = tEmprSites.Fields("raizCNPJSite").Value
            compCNPJSite = tEmprSites.Fields("compCNPJSite").Value
            digCNPJSite = tEmprSites.Fields("digCNPJSite").Value
            tipoPessoa = tEmprSites.Fields("tipoPessoa").Value
            padrao = tEmprSites.Fields("padrao").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprSites.Close()
                tEmprSites.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe EmprSites - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Long

        On Error GoTo Einclui
        tEmprSites.AddNew()
        tEmprSites(0).Value = idEmpresa
        tEmprSites(1).Value = idSite
        tEmprSites(2).Value = nomeSite
        tEmprSites(3).Value = fone1site
        tEmprSites(4).Value = ramal1Site
        tEmprSites(5).Value = fone2Site
        tEmprSites(6).Value = ramal2Site
        tEmprSites(7).Value = faxSite
        tEmprSites(8).Value = enderecoSite
        tEmprSites(9).Value = complemSite
        tEmprSites(10).Value = bairroSite
        tEmprSites(11).Value = paisSite
        tEmprSites(12).Value = ufSite
        tEmprSites(13).Value = cidadeSite
        tEmprSites(14).Value = cepSite
        tEmprSites(15).Value = regimeTrabSite
        tEmprSites(16).Value = diaInicTrabSite
        tEmprSites(17).Value = diaFimTrabSite
        tEmprSites(18).Value = refeitorioSite
        tEmprSites(19).Value = horaInicTrabSite
        tEmprSites(20).Value = horaInicDescSite
        tEmprSites(21).Value = horaFimDescSite
        tEmprSites(22).Value = horaFimTrabSite
        tEmprSites(23).Value = dtInicio
        tEmprSites(24).Value = dtFim
        tEmprSites(25).Value = raizCNPJSite
        tEmprSites(26).Value = compCNPJSite
        tEmprSites(27).Value = digCNPJSite
        tEmprSites.Fields("tipoPessoa").Value = tipoPessoa
        tEmprSites.Fields("padrao").Value = padrao
        tEmprSites.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprSites - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Long

        On Error GoTo Ealtera
        tEmprSites(2).Value = nomeSite
        tEmprSites(3).Value = fone1site
        tEmprSites(4).Value = ramal1Site
        tEmprSites(5).Value = fone2Site
        tEmprSites(6).Value = ramal2Site
        tEmprSites(7).Value = faxSite
        tEmprSites(8).Value = enderecoSite
        tEmprSites(9).Value = complemSite
        tEmprSites(10).Value = bairroSite
        tEmprSites(11).Value = paisSite
        tEmprSites(12).Value = ufSite
        tEmprSites(13).Value = cidadeSite
        tEmprSites(14).Value = cepSite
        tEmprSites(15).Value = regimeTrabSite
        tEmprSites(16).Value = diaInicTrabSite
        tEmprSites(17).Value = diaFimTrabSite
        tEmprSites(18).Value = refeitorioSite
        tEmprSites(19).Value = horaInicTrabSite
        tEmprSites(20).Value = horaInicDescSite
        tEmprSites(21).Value = horaFimDescSite
        tEmprSites(22).Value = horaFimTrabSite
        tEmprSites(23).Value = dtInicio
        tEmprSites(24).Value = dtFim
        tEmprSites(25).Value = raizCNPJSite
        tEmprSites(26).Value = compCNPJSite
        tEmprSites(27).Value = digCNPJSite
        tEmprSites.Fields("tipoPessoa").Value = tipoPessoa
        tEmprSites.Fields("padrao").Value = padrao
        tEmprSites.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprSites - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(empresa As Short, site As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprSites.                           *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * site       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String
        Dim RC As Integer

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro
        'Verifica se não haverá perde de integridade referencial não controlada por código:
        RC = VerificaSite(0, empresa, site)
        If RC <> 0 Then
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM EmprSites WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprSites - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function VerificaSite(tipo As Short, empresa As Short, site As Short, Optional data As Date = Nothing) As Integer

        '* *****************************************************************
        '* Verifica se a data de término não fere a integridade relacional *
        '* tipo = 0 => não verifica data (se um site pode ser eliminado)   *
        '*        1 => verifica data (se um site pode ser encerrado)       *
        '* *****************************************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        'Abre a tabela de mensagens:
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro ao abrir segMensagens")
            VerificaSite = -1
            Exit Function
        End If
        VerificaSite = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        '* *******************************************
        '* Verifica se não fere integridade          *
        '* relacional, quando controlada por código: *
        '* *******************************************
        'Pesquisa em BibAlocEstr:
        vSelec = "SELECT * FROM BibAlocEstr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230' "
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "BibAlocEstr com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " ativo ou com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em BibHAlocEstr:
        vSelec = "SELECT * FROM BibHAlocEstr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "BibHAlocEstr com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em EmprAlocDeptoSite:
        vSelec = "SELECT * FROM EmprAlocDeptoSite WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230' "
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "EmprAlocDeptoSite com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " ativo ou com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em InvAlocEmpr:
        vSelec = "SELECT * FROM InvAlocEmpr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230' "
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "InvAlocEmpr com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " ativo ou com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em InvHAlocEmpr:
        vSelec = "SELECT * FROM InvHAlocEmpr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "InvHAlocEmpr com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em ProjProjetos:
        vSelec = "SELECT * FROM ProjProjetos WHERE localContrProj = 1"
        vSelec = vSelec & " AND idOrganiz = " & empresa
        vSelec = vSelec & " AND idSetorOrg = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProj, 112) = '18991230' "
            vSelec = vSelec & " OR Convert(datetime, dtFinalProj, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "ProjProjetos interno com idOrganiz = " & empresa & " e idSetorOrg = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " ativo ou com data de encerramento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em RDFaturamento:
        vSelec = "SELECT * FROM RDFaturamento WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFaturamento, 112) >= '" & Format(data, "yyyymmdd") & "'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "RDFaturamento com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de faturamento ou vencimento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em RDPagamento:
        vSelec = "SELECT * FROM RDPagamentos WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtPagamento, 112) >= '" & Format(data, "yyyymmdd") & "'"
            vSelec = vSelec & " OR Convert(datetime, dtVencimento, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "RDPagamentos com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de pagamento ou vencimento maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em RDRecDesp:
        vSelec = "SELECT * FROM RDRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtEfetivacao, 112) = '18991230' "
            vSelec = vSelec & " OR Convert(datetime, dtEfetivacao, 112) >= '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "RDRecDesp com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de efetivação não informada ou maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em RDHRecDesp:
        vSelec = "SELECT * FROM RDHRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtEfetivacao, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "RDHRecDesp com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de efetivação maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em EstAlocEstrDeposito:
        vSelec = "SELECT * FROM EstAlocEstrDeposito WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtFim, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "EstAlocEstrDeposito com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data de término maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em EstCompras:
        vSelec = "SELECT * FROM EstCompras WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtCompra, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "EstCompras com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data da compra maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If
        'Pesquisa em EstVendas:
        vSelec = "SELECT * FROM EstVendas WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(datetime, dtVenda, 112) >= '" & Format(data, "yyyymmdd") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            vSelec = "EstVendas com idEmpresa = " & empresa & " e idSite = " & site
            If tipo = 1 Then
                vSelec = vSelec & Chr(10) & " com data da venda maior que " & data
            End If
            MsgBox(vSelec)
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            VerificaSite = RC
            Exit Function
        End If

    End Function


End Class
