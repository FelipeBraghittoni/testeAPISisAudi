Option Explicit On
Imports System.IO

Public Class EmprEmpresas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850310"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320310"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0310"
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


    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvarnomeEmpresa
        End Get
        Set(value As String)
            mvarnomeEmpresa = value
        End Set
    End Property



    Private mvarraizCNPJEmpresa As Integer
    Public Property raizCNPJEmpresa() As Integer
        Get
            Return mvarraizCNPJEmpresa
        End Get
        Set(value As Integer)
            mvarraizCNPJEmpresa = value
        End Set
    End Property


    Private mvarcompCNPJEmpresa As Short
    Public Property compCNPJEmpresa() As Short
        Get
            Return mvarcompCNPJEmpresa
        End Get
        Set(value As Short)
            mvarcompCNPJEmpresa = value
        End Set
    End Property


    Private mvardigCNPJEmpresa As Short
    Public Property digCNPJEmpresa() As Short
        Get
            Return mvardigCNPJEmpresa
        End Get
        Set(value As Short)
            mvardigCNPJEmpresa = value
        End Set
    End Property


    Private mvarenderecoEmpresa As String
    Public Property enderecoEmpresa() As String
        Get
            Return mvarenderecoEmpresa
        End Get
        Set(value As String)
            mvarenderecoEmpresa = value
        End Set
    End Property


    Private mvarcomplemEmpresa As String
    Public Property complemEmpresa() As String
        Get
            Return mvarcomplemEmpresa
        End Get
        Set(value As String)
            mvarcomplemEmpresa = value
        End Set
    End Property


    Private mvarbairroEmpresa As String
    Public Property bairroEmpresa() As String
        Get
            Return mvarbairroEmpresa
        End Get
        Set(value As String)
            mvarbairroEmpresa = value
        End Set
    End Property


    Private mvarpaisEmpresa As Short
    Public Property paisEmpresa() As Short
        Get
            Return mvarpaisEmpresa
        End Get
        Set(value As Short)
            mvarpaisEmpresa = value
        End Set
    End Property


    Private mvarUFEmpresa As String
    Public Property UFEmpresa() As String
        Get
            Return mvarUFEmpresa
        End Get
        Set(value As String)
            mvarUFEmpresa = value
        End Set
    End Property


    Private mvarcidadeEmpresa As Short
    Public Property cidadeEmpresa() As Short
        Get
            Return mvarcidadeEmpresa
        End Get
        Set(value As Short)
            mvarcidadeEmpresa = value
        End Set
    End Property


    Private mvarcepEmpresa As Integer
    Public Property cepEmpresa() As Integer
        Get
            Return mvarcepEmpresa
        End Get
        Set(value As Integer)
            mvarcepEmpresa = value
        End Set
    End Property


    Private mvarinternetEmpresa As String
    Public Property internetEmpresa() As String
        Get
            Return mvarinternetEmpresa
        End Get
        Set(value As String)
            mvarinternetEmpresa = value
        End Set
    End Property


    Private mvaremailEmpresa As String
    Public Property emailEmpresa() As String
        Get
            Return mvaremailEmpresa
        End Get
        Set(value As String)
            mvaremailEmpresa = value
        End Set
    End Property


    Private mvarfone1Empresa As Integer
    Public Property fone1Empresa() As Integer
        Get
            Return mvarfone1Empresa
        End Get
        Set(value As Integer)
            mvarfone1Empresa = value
        End Set
    End Property


    Private mvarfone2Empresa As Integer
    Public Property fone2Empresa() As Integer
        Get
            Return mvarfone2Empresa
        End Get
        Set(value As Integer)
            mvarfone2Empresa = value
        End Set
    End Property


    Private mvarfaxEmpresa As Integer
    Public Property faxEmpresa() As Integer
        Get
            Return mvarfaxEmpresa
        End Get
        Set(value As Integer)
            mvarfaxEmpresa = value
        End Set
    End Property


    Private mvartpLogo1Empresa As String
    Public Property tpLogo1Empresa() As String
        Get
            Return mvartpLogo1Empresa
        End Get
        Set(value As String)
            mvartpLogo1Empresa = value
        End Set
    End Property


    Private mvartpLogo2Empresa As String
    Public Property tpLogo2Empresa() As String
        Get
            Return mvartpLogo2Empresa
        End Get
        Set(value As String)
            mvartpLogo2Empresa = value
        End Set
    End Property


    Private mvarstatusEmpresa As Short
    Public Property statusEmpresa() As Short
        Get
            Return mvarstatusEmpresa
        End Get
        Set(value As Short)
            mvarstatusEmpresa = value
        End Set
    End Property

    Private mvaridHolding As Short
    Public Property idHolding() As Short
        Get
            Return mvaridHolding
        End Get
        Set(value As Short)
            mvaridHolding = value
        End Set
    End Property


    Private mvarnomeFantasia As String
    Public Property nomeFantasia() As String
        Get
            Return mvarnomeFantasia
        End Get
        Set(value As String)
            mvarnomeFantasia = value
        End Set
    End Property


    Private mvaridTpJuridico As Short
    Public Property idTpJuridico() As Short
        Get
            Return mvaridTpJuridico
        End Get
        Set(value As Short)
            mvaridTpJuridico = value
        End Set
    End Property


    Private mvaridRamoAtiv As Short
    Public Property idRamoAtiv() As Short
        Get
            Return mvaridRamoAtiv
        End Get
        Set(value As Short)
            mvaridRamoAtiv = value
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


    Private mvaridRegFiscal As Short
    Public Property idRegFiscal() As Short
        Get
            Return mvaridRegFiscal
        End Get
        Set(value As Short)
            mvaridRegFiscal = value
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

#End Region

#Region "instancias de classes"

    Private mvarEmprCargos As EmprCargos
    Public Property EmprCargos() As EmprCargos
        Get
            Return mvarEmprCargos
        End Get
        Set(value As EmprCargos)
            mvarEmprCargos = value
        End Set
    End Property


    Private mvarEmprCompSoc As EmprCompSoc
    Public Property EmprCompSoc() As EmprCompSoc
        Get
            Return mvarEmprCompSoc
        End Get
        Set(value As EmprCompSoc)
            mvarEmprCompSoc = value
        End Set
    End Property


    Private mvarEmprDoctos As EmprDoctos
    Public Property EmprDoctos() As EmprDoctos
        Get
            Return mvarEmprDoctos
        End Get
        Set(value As EmprDoctos)
            mvarEmprDoctos = value
        End Set
    End Property


    Private mvarEmprOcorrenc As EmprOcorrenc
    Public Property EmprOcorrenc() As EmprOcorrenc
        Get
            Return mvarEmprOcorrenc
        End Get
        Set(value As EmprOcorrenc)
            mvarEmprOcorrenc = value
        End Set
    End Property


    Private mvarEmprRegEmpr As EmprRegEmpr
    Public Property EmprRegEmpr() As EmprRegEmpr
        Get
            Return mvarEmprRegEmpr
        End Get
        Set(value As EmprRegEmpr)
            mvarEmprRegEmpr = value
        End Set
    End Property


    Private mvarEmprSites As EmprSites
    Public Property EmprSites() As EmprSites
        Get
            Return mvarEmprSites
        End Get
        Set(value As EmprSites)
            mvarEmprSites = value
        End Set
    End Property


    Private mvarEmprDeptos As EmprDeptos
    Public Property EmprDeptos() As EmprDeptos
        Get
            Return mvarEmprDeptos
        End Get
        Set(value As EmprDeptos)
            mvarEmprDeptos = value
        End Set
    End Property


#End Region

#Region "Conexão com banco"

    Public Db As ADODB.Connection
    Public tEmprEmpresas As ADODB.Recordset

#End Region


    Private Sub Class_Initialize()

        'create the mEmprOcorrenc object when the Empresa class is created
        mvarEmprOcorrenc = New EmprOcorrenc
        'create the mEmprDoctos object when the Empresa class is created
        mvarEmprDoctos = New EmprDoctos
        'create the mEmprDeptos object when the EmprEmpresas class is created
        mvarEmprDeptos = New EmprDeptos
        'create the mEmprCargos object when the EmprEmpresas class is created
        mvarEmprCargos = New EmprCargos
        'create the mEmprRegEmpr object when the EmprEmpresas class is created
        mvarEmprRegEmpr = New EmprRegEmpr
        'create the mEmprSites object when the EmprEmpresas class is created
        mvarEmprSites = New EmprSites
        'create the mEmprCompSoc object when the EmprEmpresas class is created
        mvarEmprCompSoc = New EmprCompSoc

    End Sub

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
        tEmprEmpresas = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            tEmprEmpresas.Open("EmprEmpresas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                    'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM EmprEmpresas ORDER BY nomeEmpresa"
                tEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprEmpresas.Close()

                If tipo = 1 Or tipo = 2 Then
                    tEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprEmpresas.Open("EmprEmpresas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprEmpresas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprEmpresas.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprEmpresas.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprEmpresas.EOF Then
            idEmpresa = tEmprEmpresas.Fields("idEmpresa").Value
            nomeEmpresa = tEmprEmpresas.Fields("nomeEmpresa").Value
            raizCNPJEmpresa = tEmprEmpresas.Fields("raizCNPJEmpresa").Value
            compCNPJEmpresa = tEmprEmpresas.Fields("compCNPJEmpresa").Value
            digCNPJEmpresa = tEmprEmpresas.Fields("digCNPJEmpresa").Value
            enderecoEmpresa = tEmprEmpresas.Fields("enderecoEmpresa").Value
            complemEmpresa = tEmprEmpresas.Fields("complemEmpresa").Value
            bairroEmpresa = tEmprEmpresas.Fields("bairroEmpresa").Value
            paisEmpresa = tEmprEmpresas.Fields("paisEmpresa").Value
            UFEmpresa = tEmprEmpresas.Fields("UFEmpresa").Value
            cidadeEmpresa = tEmprEmpresas.Fields("cidadeEmpresa").Value
            cepEmpresa = tEmprEmpresas.Fields("cepEmpresa").Value
            internetEmpresa = tEmprEmpresas.Fields("internetEmpresa").Value
            emailEmpresa = tEmprEmpresas.Fields("emailEmpresa").Value
            fone1Empresa = tEmprEmpresas.Fields("fone1Empresa").Value
            fone2Empresa = tEmprEmpresas.Fields("fone2Empresa").Value
            faxEmpresa = tEmprEmpresas.Fields("faxEmpresa").Value
            tpLogo1Empresa = tEmprEmpresas.Fields("tpLogo1Empresa").Value
            tpLogo2Empresa = tEmprEmpresas.Fields("tpLogo2Empresa").Value
            statusEmpresa = tEmprEmpresas.Fields("statusEmpresa").Value
            idHolding = tEmprEmpresas.Fields("idHolding").Value
            nomeFantasia = tEmprEmpresas.Fields("nomeFantasia").Value
            idTpJuridico = tEmprEmpresas.Fields("idTpJuridico").Value
            idRamoAtiv = tEmprEmpresas.Fields("idRamoAtiv").Value
            dtInicio = tEmprEmpresas.Fields("dtInicio").Value
            dtFim = tEmprEmpresas.Fields("dtFim").Value
            idRegFiscal = tEmprEmpresas.Fields("idRegFiscal").Value
            tipoPessoa = tEmprEmpresas.Fields("tipoPessoa").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprEmpresas - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function localiza(idEmpresa As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprEmpresas.                        *
        '* *                                      *
        '* * idEmpresa  = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprEmpresas WHERE idEmpresa = " & idEmpresa
        tEmprEmpresas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprEmpresas.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprEmpresas.Fields("idEmpresa").Value
            nomeEmpresa = tEmprEmpresas.Fields("nomeEmpresa").Value
            raizCNPJEmpresa = tEmprEmpresas.Fields("raizCNPJEmpresa").Value
            compCNPJEmpresa = tEmprEmpresas.Fields("compCNPJEmpresa").Value
            digCNPJEmpresa = tEmprEmpresas.Fields("digCNPJEmpresa").Value
            enderecoEmpresa = tEmprEmpresas.Fields("enderecoEmpresa").Value
            complemEmpresa = tEmprEmpresas.Fields("complemEmpresa").Value
            bairroEmpresa = tEmprEmpresas.Fields("bairroEmpresa").Value
            paisEmpresa = tEmprEmpresas.Fields("paisEmpresa").Value
            UFEmpresa = tEmprEmpresas.Fields("UFEmpresa").Value
            cidadeEmpresa = tEmprEmpresas.Fields("cidadeEmpresa").Value
            cepEmpresa = tEmprEmpresas.Fields("cepEmpresa").Value
            internetEmpresa = tEmprEmpresas.Fields("internetEmpresa").Value
            emailEmpresa = tEmprEmpresas.Fields("emailEmpresa").Value
            fone1Empresa = tEmprEmpresas.Fields("fone1Empresa").Value
            fone2Empresa = tEmprEmpresas.Fields("fone2Empresa").Value
            faxEmpresa = tEmprEmpresas.Fields("faxEmpresa").Value
            tpLogo1Empresa = tEmprEmpresas.Fields("tpLogo1Empresa").Value
            tpLogo2Empresa = tEmprEmpresas.Fields("tpLogo2Empresa").Value
            statusEmpresa = tEmprEmpresas.Fields("statusEmpresa").Value
            idHolding = tEmprEmpresas.Fields("idHolding").Value
            nomeFantasia = tEmprEmpresas.Fields("nomeFantasia").Value
            idTpJuridico = tEmprEmpresas.Fields("idTpJuridico").Value
            idRamoAtiv = tEmprEmpresas.Fields("idRamoAtiv").Value
            dtInicio = tEmprEmpresas.Fields("dtInicio").Value
            dtFim = tEmprEmpresas.Fields("dtFim").Value
            idRegFiscal = tEmprEmpresas.Fields("idRegFiscal").Value
            tipoPessoa = tEmprEmpresas.Fields("tipoPessoa").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprEmpresas.Close()
                tEmprEmpresas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprEmpresas - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function
    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprEmpresas.                        *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM EmprEmpresas WHERE nomeEmpresa = '" & descric & "'"
        tEmprEmpresas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprEmpresas.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprEmpresas.Fields("idEmpresa").Value
            nomeEmpresa = tEmprEmpresas.Fields("nomeEmpresa").Value
            raizCNPJEmpresa = tEmprEmpresas.Fields("raizCNPJEmpresa").Value
            compCNPJEmpresa = tEmprEmpresas.Fields("compCNPJEmpresa").Value
            digCNPJEmpresa = tEmprEmpresas.Fields("digCNPJEmpresa").Value
            enderecoEmpresa = tEmprEmpresas.Fields("enderecoEmpresa").Value
            complemEmpresa = tEmprEmpresas.Fields("complemEmpresa").Value
            bairroEmpresa = tEmprEmpresas.Fields("bairroEmpresa").Value
            paisEmpresa = tEmprEmpresas.Fields("paisEmpresa").Value
            UFEmpresa = tEmprEmpresas.Fields("UFEmpresa").Value
            cidadeEmpresa = tEmprEmpresas.Fields("cidadeEmpresa").Value
            cepEmpresa = tEmprEmpresas.Fields("cepEmpresa").Value
            internetEmpresa = tEmprEmpresas.Fields("internetEmpresa").Value
            emailEmpresa = tEmprEmpresas.Fields("emailEmpresa").Value
            fone1Empresa = tEmprEmpresas.Fields("fone1Empresa").Value
            fone2Empresa = tEmprEmpresas.Fields("fone2Empresa").Value
            faxEmpresa = tEmprEmpresas.Fields("faxEmpresa").Value
            tpLogo1Empresa = tEmprEmpresas.Fields("tpLogo1Empresa").Value
            tpLogo2Empresa = tEmprEmpresas.Fields("tpLogo2Empresa").Value
            statusEmpresa = tEmprEmpresas.Fields("statusEmpresa").Value
            idHolding = tEmprEmpresas.Fields("idHolding").Value
            nomeFantasia = tEmprEmpresas.Fields("nomeFantasia").Value
            idTpJuridico = tEmprEmpresas.Fields("idTpJuridico").Value
            idRamoAtiv = tEmprEmpresas.Fields("idRamoAtiv").Value
            dtInicio = tEmprEmpresas.Fields("dtInicio").Value
            dtFim = tEmprEmpresas.Fields("dtFim").Value
            idRegFiscal = tEmprEmpresas.Fields("idRegFiscal").Value
            tipoPessoa = tEmprEmpresas.Fields("tipoPessoa").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprEmpresas.Close()
                tEmprEmpresas.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe EmprEmpresas - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(picture1 As Integer, vAux1 As String, picture2 As Integer, vAux2 As String, missao As String) As Integer

        'Define Variáveis:
        Dim vAux As String

        On Error GoTo Einclui
        tEmprEmpresas.AddNew()
        tEmprEmpresas(0).Value = idEmpresa
        tEmprEmpresas(1).Value = nomeEmpresa
        tEmprEmpresas(2).Value = raizCNPJEmpresa
        tEmprEmpresas(3).Value = compCNPJEmpresa
        tEmprEmpresas(4).Value = digCNPJEmpresa
        tEmprEmpresas(5).Value = enderecoEmpresa
        tEmprEmpresas(6).Value = complemEmpresa
        tEmprEmpresas(7).Value = bairroEmpresa
        tEmprEmpresas(8).Value = paisEmpresa
        tEmprEmpresas(9).Value = UFEmpresa
        tEmprEmpresas(10).Value = cidadeEmpresa
        tEmprEmpresas(11).Value = cepEmpresa
        tEmprEmpresas(12).Value = internetEmpresa
        tEmprEmpresas(13).Value = emailEmpresa
        tEmprEmpresas(14).Value = fone1Empresa
        tEmprEmpresas(15).Value = fone2Empresa
        tEmprEmpresas(16).Value = faxEmpresa
        'Carrega Logotipo 1
        If picture1 = 0 Then
            tpLogo1Empresa = " "
            tEmprEmpresas(17).Value = ""
        Else
            'Carrega o conteúdo do arquivo na coluna da tabela:
            CarregaFigura(vAux1, 17)
            'Elimina o arquivo do disco:
            Kill(vAux1)
        End If
        tEmprEmpresas(18).Value = tpLogo1Empresa
        'Carrega Logotipo 2
        If picture2 = 0 Then
            tpLogo2Empresa = " "
            tEmprEmpresas(19).Value = ""
        Else
            'Carrega o conteúdo do arquivo na coluna da tabela:
            CarregaFigura(vAux2, 19)
            'Elimina o arquivo do disco:
            Kill(vAux2)
        End If
        tEmprEmpresas(20).Value = tpLogo2Empresa
        'Missão da Empresa:
        tEmprEmpresas(21).Value = IIf(missao = "", " ", missao)
        tEmprEmpresas(22).Value = statusEmpresa
        tEmprEmpresas(23).Value = idHolding
        tEmprEmpresas(24).Value = nomeFantasia
        tEmprEmpresas(25).Value = idTpJuridico
        tEmprEmpresas(26).Value = idRamoAtiv
        tEmprEmpresas(27).Value = dtInicio
        tEmprEmpresas(28).Value = dtFim
        tEmprEmpresas(29).Value = idRegFiscal
        tEmprEmpresas.Fields("tipoPessoa").Value = tipoPessoa
        tEmprEmpresas.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprEmpresas - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function altera(picture1 As Integer, vAux1 As String, picture2 As Integer, vAux2 As String, missao As String) As Integer

        'Define variáveis:
        Dim vAux As String

        On Error GoTo Ealtera
        tEmprEmpresas(1).Value = nomeEmpresa
        tEmprEmpresas(2).Value = raizCNPJEmpresa
        tEmprEmpresas(3).Value = compCNPJEmpresa
        tEmprEmpresas(4).Value = digCNPJEmpresa
        tEmprEmpresas(5).Value = enderecoEmpresa
        tEmprEmpresas(6).Value = complemEmpresa
        tEmprEmpresas(7).Value = bairroEmpresa
        tEmprEmpresas(8).Value = paisEmpresa
        tEmprEmpresas(9).Value = UFEmpresa
        tEmprEmpresas(10).Value = cidadeEmpresa
        tEmprEmpresas(11).Value = cepEmpresa
        tEmprEmpresas(12).Value = internetEmpresa
        tEmprEmpresas(13).Value = emailEmpresa
        tEmprEmpresas(14).Value = fone1Empresa
        tEmprEmpresas(15).Value = fone2Empresa
        tEmprEmpresas(16).Value = faxEmpresa
        'Carrega Logotipo 1
        If picture1 = 0 Then
            tpLogo1Empresa = " "
            tEmprEmpresas(17).Value = ""
        Else
            'Carrega o conteúdo do arquivo na coluna da tabela:
            CarregaFigura(vAux1, 17)
            'Elimina o arquivo do disco:
            Kill(vAux1)
        End If
        tEmprEmpresas(18).Value = tpLogo1Empresa
        'Carrega Logotipo 2
        If picture2 = 0 Then
            tpLogo2Empresa = " "
            tEmprEmpresas(19).Value = ""
        Else
            'Carrega o conteúdo do arquivo na coluna da tabela:
            CarregaFigura(vAux2, 19)
            'Elimina o arquivo do disco:
            Kill(vAux2)
        End If
        tEmprEmpresas(20).Value = tpLogo2Empresa
        'Missão da Empresa:
        tEmprEmpresas(21).Value = IIf(missao = "", " ", missao)
        tEmprEmpresas(22).Value = statusEmpresa
        tEmprEmpresas(23).Value = idHolding
        tEmprEmpresas(24).Value = nomeFantasia
        tEmprEmpresas(25).Value = idTpJuridico
        tEmprEmpresas(26).Value = idRamoAtiv
        tEmprEmpresas(27).Value = dtInicio
        tEmprEmpresas(28).Value = dtFim
        tEmprEmpresas(29).Value = idRegFiscal
        tEmprEmpresas.Fields("tipoPessoa").Value = tipoPessoa
        tEmprEmpresas.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprEmpresas - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Sub GravaArqFig(colF As Short, nomeTabF As String) tEmprEmpresas(colF).GetChunk(ChunkSize * 2)
        'Extrai o conteúdo de um campo binário, em um arquivo

        'Define as variáveis:
        Dim J As Short
        Dim ChunkSize As Integer
        Dim CurSize As Integer
        'Dim CurChunk As String

        Dim CurChunk() As Byte
        Dim fs As FileStream
        Dim writer As BinaryWriter

        On Error Resume Next
        ChunkSize = 8192    'Define o tamanho de cada pedaço
        J = FreeFile()        'Obtém um número livre de arquivo
        'Abre o arquivo como binário:
        fs = File.Open(nomeTabF, FileMode.Create)
        writer = New BinaryWriter(fs)
        Do
            CurChunk = tEmprEmpresas(colF).GetChunk(ChunkSize * 2)
            ' Write each byte
            For Each value As Byte In CurChunk
                writer.Write(value)
            Next
            If CurChunk.Length < ChunkSize Then Exit Do
        Loop
        writer.Close()
        fs.Close()

    End Sub

    Public Sub CarregaFigura(nomeTabF As String, colF As Short)
        'Grava o conteúdo do arquivo no campo binário

        'Define as variáveis:
        Dim TotalSize As Integer
        Dim CurChunk As String
        Dim ChunkSize As Integer
        Dim J As Integer

        ChunkSize = 8192    'Define o tamanho de cada pedaço
        J = FreeFile()        'Obtém um número livre de arquivo
        'Abre o arquivo como binário:
        Open(nomeTabF) For Binary As #J
    TotalSize = LOF(J)
        Do While Not EOF(J)
            If TotalSize - Seek(J) < ChunkSize Then
                ChunkSize = TotalSize - Seek(J) + 10
            End If
            CurChunk = String$(ChunkSize + 1, 32)
        'Lê o pedaço do arquivo:
        Get #J, , CurChunk
        'Grava o pedaço no final da coluna:
        tEmprEmpresas(colF).AppendChunk(CurChunk)
        Loop
        Close #J

End Sub

    Public Function LeMissao() As String

        LeMissao = tEmprEmpresas.Fields("missaoEmpresa").Value

    End Function

    Public Function elimina(empresa As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprEmpresas.                        *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro ao acessar a Tabela segMensagens")
            elimina = 3
            Exit Function
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        '* *******************************************
        '* Verifica se não fere integridade
        '* relacional, quando controlada por código: *
        '* *******************************************
        'Pesquisa em BibAlocEstr:
        vSelec = "SELECT * FROM BibAlocEstr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela BibAlocEstr - não eliminado")
                elimina = 10
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em BibHAlocEstr:
        vSelec = "SELECT * FROM BibHAlocEstr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela BibHAlocEstr - não eliminado")
                elimina = 11
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em CandVagas:
        vSelec = "SELECT * FROM CandVagas WHERE (idOrganiz = " & empresa
        vSelec = vSelec & " AND localTrabVaga = 1)"
        vSelec = vSelec & " OR idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela CandVagas - não eliminado")
                elimina = 12
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ColExtensaoBeneficio:
        vSelec = "SELECT * FROM ColExtensaoBeneficio WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ColExtensaoBeneficio - não eliminado")
                elimina = 13
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorAlocProjetos:
        vSelec = "SELECT * FROM OcorAlocProjetos WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND localTrabColab = " & 1
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorAlocProjetos - não eliminado")
                elimina = 14
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorApontHoras:
        vSelec = "SELECT * FROM OcorApontHoras WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorApontHoras - não eliminado")
                elimina = 15
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorAumProm:
        vSelec = "SELECT * FROM OcorAumProm WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorAumProm - não eliminado")
                elimina = 16
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorAvaliacEmpr:
        vSelec = "SELECT * FROM OcorAvaliacEmpr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorAvaliacEmpr - não eliminado")
                elimina = 17
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorCursosColab:
        vSelec = "SELECT * FROM OcorCursosColab WHERE idInstituicao = " & empresa
        vSelec = vSelec & " AND tpInstituicao = 1"
        vSelec = vSelec & " AND Convert(datetime, dtFinal, 112) = '18991230'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorCursosColab - não eliminado")
                elimina = 18
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorDeptoColab:
        vSelec = "SELECT * FROM OcorDeptoColab WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorDeptoColab - não eliminado")
                elimina = 19
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorEmpresaColab:
        vSelec = "SELECT * FROM OcorEmpresaColab WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorEmpresaColab - não eliminado")
                elimina = 20
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorSaldoHoras:
        vSelec = "SELECT * FROM OcorSaldoHoras WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorSaldoHoras - não eliminado")
                elimina = 21
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHAlocProjetos:
        vSelec = "SELECT * FROM OcorHAlocProjetos WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND localTrabColab = " & 1
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHAlocProjetos - não eliminado")
                elimina = 22
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHApontHoras:
        vSelec = "SELECT * FROM OcorHApontHoras WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHApontHoras - não eliminado")
                elimina = 23
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHAumProm:
        vSelec = "SELECT * FROM OcorHAumProm WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHAumProm - não eliminado")
                elimina = 24
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHAvaliacEmpr:
        vSelec = "SELECT * FROM OcorHAvaliacEmpr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHAvaliacEmpr - não eliminado")
                elimina = 25
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHCursosColab:
        vSelec = "SELECT * FROM OcorHCursosColab WHERE idInstituicao = " & empresa
        vSelec = vSelec & " AND tpInstituicao = 1"
        vSelec = vSelec & " AND Convert(datetime, dtFinal, 112) = '18991230'"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHCursosColab - não eliminado")
                elimina = 26
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHDeptoColab:
        vSelec = "SELECT * FROM OcorHDeptoColab WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHDeptoColab - não eliminado")
                elimina = 27
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHEmpresaColab:
        vSelec = "SELECT * FROM OcorHEmpresaColab WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHEmpresaColab - não eliminado")
                elimina = 28
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em OcorHSaldoHoras:
        vSelec = "SELECT * FROM OcorHSaldoHoras WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela OcorHSaldoHoras - não eliminado")
                elimina = 29
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ContrCadastro:
        vSelec = "SELECT * FROM ContrCadastro WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ContrCadastro - não eliminado")
                elimina = 30
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ContrColabContrato:
        vSelec = "SELECT * FROM ContrColabContrato WHERE idEntidade = " & empresa
        vSelec = vSelec & " AND (entidadesPartic = 0"
        vSelec = vSelec & " OR entidadesPartic = 3"
        vSelec = vSelec & " OR entidadesPartic = 4)"
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ContrColabContrato - não eliminado")
                elimina = 31
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ContrProcuracoes:
        vSelec = "SELECT * FROM ContrProcuracoes WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ContrProcuracoes - não eliminado")
                elimina = 32
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em DirDoctosDir:
        vSelec = "SELECT * FROM DirDoctosDir WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela DirDoctosDir - não eliminado")
                elimina = 33
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprAlocCargos:
        vSelec = "SELECT * FROM EmprAlocCargos WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprAlocCargos - não eliminado")
                elimina = 34
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprAlocDeptoSite:
        vSelec = "SELECT * FROM EmprAlocDeptoSite WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprAlocDeptoSite - não eliminado")
                elimina = 35
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprAlocGerDepto:
        vSelec = "SELECT * FROM EmprAlocGerDepto WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprAlocGerDepto - não eliminado")
                elimina = 36
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprAvaliacClie:
        vSelec = "SELECT * FROM EmprAvaliacClie WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprAvaliacClie - não eliminado")
                elimina = 37
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprAvaliacForn:
        vSelec = "SELECT * FROM EmprAvaliacForn WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprAvaliacForn - não eliminado")
                elimina = 38
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprCompSoc:
        vSelec = "SELECT * FROM EmprCompSoc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprCompSoc - não eliminado")
                elimina = 39
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprDeptos:
        vSelec = "SELECT * FROM EmprDeptos WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprDeptos - não eliminado")
                elimina = 40
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprDoctos:
        vSelec = "SELECT * FROM EmprDoctos WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprDoctos - não eliminado")
                elimina = 41
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprOcorrenc:
        vSelec = "SELECT * FROM EmprOcorrenc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprOcorrenc - não eliminado")
                elimina = 42
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprRegEmpr:
        vSelec = "SELECT * FROM EmprRegEmpr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprRegEmpr - não eliminado")
                elimina = 43
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EmprSites:
        vSelec = "SELECT * FROM EmprSites WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EmprSites - não eliminado")
                elimina = 44
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em EstDeposito:
        vSelec = "SELECT * FROM EstDeposito WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela EstDeposito - não eliminado")
                elimina = 45
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em GerFeriados:
        vSelec = "SELECT * FROM GerFeriados WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND tpLocal = " & 1
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela GerFeriados - não eliminado")
                elimina = 46
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em InvAlocEmpr:
        vSelec = "SELECT * FROM InvAlocEmpr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela InvAlocEmpr - não eliminado")
                elimina = 47
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em InvCadastro:
        vSelec = "SELECT * FROM InvCadastro WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela InvCadastro - não eliminado")
                elimina = 48
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em InvHAlocEmpr:
        vSelec = "SELECT * FROM InvHAlocEmpr WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela InvHAlocEmpr - não eliminado")
                elimina = 49
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em InvHCadastro:
        vSelec = "SELECT * FROM InvHCadastro WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela InvHCadastro - não eliminado")
                elimina = 50
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ParCadastro:
        vSelec = "SELECT * FROM ParCadastro WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ParCadastro - não eliminado")
                elimina = 51
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ProjProjetos:
        vSelec = "SELECT * FROM ProjProjetos WHERE (idOrganiz = " & empresa
        vSelec = vSelec & " AND localContrProj = 1)"
        vSelec = vSelec & " OR idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ProjProjetos - não eliminado")
                elimina = 52
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em ProjHorasConsumidas:
        vSelec = "SELECT * FROM ProjHorasConsumidas WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela ProjHorasConsumidas - não eliminado")
                elimina = 53
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDCadCtasBco:
        vSelec = "SELECT * FROM RDCadCtasBco WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDCadCtasBco - não eliminado")
                elimina = 54
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDEstrutImp:
        vSelec = "SELECT * FROM RDEstrutImp WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDEstrutImp - não eliminado")
                elimina = 55
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDFaturamento:
        vSelec = "SELECT * FROM RDFaturamento WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDFaturamento - não eliminado")
                elimina = 56
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDFluxoCaixa:
        vSelec = "SELECT * FROM RDFluxoCaixa WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDFluxoCaixa - não eliminado")
                elimina = 57
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHFluxoCaixa:
        vSelec = "SELECT * FROM RDHFluxoCaixa WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHFluxoCaixa - não eliminado")
                elimina = 58
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDNotasExplicat:
        vSelec = "SELECT * FROM RDNotasExplicat WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDNotasExplicat - não eliminado")
                elimina = 59
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDPagamentos:
        vSelec = "SELECT * FROM RDPagamentos WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDPagamentos - não eliminado")
                elimina = 60
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDRecDesp:
        vSelec = "SELECT * FROM RDRecDesp WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDRecDesp - não eliminado")
                elimina = 61
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHRecDesp:
        vSelec = "SELECT * FROM RDHRecDesp WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHRecDesp - não eliminado")
                elimina = 62
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDRecDespEliminados:
        vSelec = "SELECT * FROM RDRecDespEliminados WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDRecDespEliminados - não eliminado")
                elimina = 63
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoBanco:
        vSelec = "SELECT * FROM RDSaldoBanco WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoBanco - não eliminado")
                elimina = 64
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoBancoOrc:
        vSelec = "SELECT * FROM RDSaldoBancoOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoBancoOrc - não eliminado")
                elimina = 65
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoBanco:
        vSelec = "SELECT * FROM RDHSaldoBanco WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoBanco - não eliminado")
                elimina = 66
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoBancoOrc:
        vSelec = "SELECT * FROM RDHSaldoBancoOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoBancoOrc - não eliminado")
                elimina = 67
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoCaixa:
        vSelec = "SELECT * FROM RDSaldoCaixa WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoCaixa - não eliminado")
                elimina = 68
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoCaixaOrc:
        vSelec = "SELECT * FROM RDSaldoCaixaOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoCaixaOrc - não eliminado")
                elimina = 69
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoCaixa:
        vSelec = "SELECT * FROM RDHSaldoCaixa WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoCaixa - não eliminado")
                elimina = 70
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoCaixaOrc:
        vSelec = "SELECT * FROM RDHSaldoCaixaOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoCaixaOrc - não eliminado")
                elimina = 71
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoRecDesp:
        vSelec = "SELECT * FROM RDSaldoRecDesp WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoRecDesp - não eliminado")
                elimina = 72
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDSaldoRecDespOrc:
        vSelec = "SELECT * FROM RDSaldoRecDespOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDSaldoRecDespOrc - não eliminado")
                elimina = 73
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoRecDesp:
        vSelec = "SELECT * FROM RDHSaldoRecDesp WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoRecDesp - não eliminado")
                elimina = 74
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If
        'Pesquisa em RDHSaldoRecDespOrc:
        vSelec = "SELECT * FROM RDHSaldoRecDespOrc WHERE idEmpresa = " & empresa
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then
                MsgBox("Há referência na tabela RDHSaldoRecDespOrc - não eliminado")
                elimina = 75
                Exit Function
            End If
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = 2
            Exit Function
        End If

        vSelec = "DELETE FROM EmprEmpresas WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprEmpresas - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
