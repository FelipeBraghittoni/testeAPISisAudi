Option Explicit On
Public Class repEmprEmpresas
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850325"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320325"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0325"

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

    Private mvarCNPJEmpr As String
    Public Property CNPJEmpr() As String
        Get
            Return mvarCNPJEmpr
        End Get
        Set(value As String)
            mvarCNPJEmpr = value
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

    Private mvarpaisEmpresa As String
    Public Property paisEmpresa() As String
        Get
            Return mvarpaisEmpresa
        End Get
        Set(value As String)
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

    Private mvarcidadeEmpresa As String
    Public Property cidadeEmpresa() As String
        Get
            Return mvarUFEmpresa
        End Get
        Set(value As String)
            mvarUFEmpresa = value
        End Set
    End Property

    Private mvarcCepEmpresa As String
    Public Property cCepEmpresa() As String
        Get
            Return mvarcCepEmpresa
        End Get
        Set(value As String)
            mvarcCepEmpresa = value
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

    Private mvardFone1Empresa As String
    Public Property dFone1Empresa() As String
        Get
            Return mvardFone1Empresa
        End Get
        Set(value As String)
            mvardFone1Empresa = value
        End Set
    End Property

    Private mvardFone2Empresa As String
    Public Property dFone2Empresa() As String
        Get
            Return mvardFone2Empresa
        End Get
        Set(value As String)
            mvardFone2Empresa = value
        End Set
    End Property

    Private mvardFaxEmpresa As String
    Public Property dFaxEmpresa() As String
        Get
            Return mvardFaxEmpresa
        End Get
        Set(value As String)
            mvardFaxEmpresa = value
        End Set
    End Property

    Private mvarstatus As String
    Public Property status() As String
        Get
            Return mvarstatus
        End Get
        Set(value As String)
            mvarstatus = value
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

    Private mvarnomeTpJuridico As String
    Public Property nomeTpJuridico() As String
        Get
            Return mvarnomeTpJuridico
        End Get
        Set(value As String)
            mvarnomeTpJuridico = value
        End Set
    End Property

    Private mvarnomeRamoAtiv As String
    Public Property nomeRamoAtiv() As String
        Get
            Return mvarnomeRamoAtiv
        End Get
        Set(value As String)
            mvarnomeRamoAtiv = value
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

    Private mvardDtFim As String
    Public Property dDtFim() As String
        Get
            Return mvardDtFim
        End Get
        Set(value As String)
            mvardDtFim = value
        End Set
    End Property

    Private mvarnomeRegFiscal As String
    Public Property nomeRegFiscal() As String
        Get
            Return mvarnomeRegFiscal
        End Get
        Set(value As String)
            mvarnomeRegFiscal = value
        End Set
    End Property

    Private mvartipoPessoa As String
    Public Property tipoPessoa() As String
        Get
            Return mvartipoPessoa
        End Get
        Set(value As String)
            mvartipoPessoa = value
        End Set
    End Property


#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public trepEmprEmpresas As ADODB.Recordset
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
        trepEmprEmpresas = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            trepEmprEmpresas.Open("repEmprEmpresas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                    'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM repEmprEmpresas ORDER BY nomeEmpresa"
                trepEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprEmpresas.Close()

                If tipo = 1 Or tipo = 2 Then
                    trepEmprEmpresas.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprEmpresas.Open("repEmprEmpresas", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprEmpresas - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprEmpresas.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprEmpresas.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprEmpresas.EOF Then
            idEmpresa = trepEmprEmpresas.Fields("idEmpresa").Value
            nomeEmpresa = trepEmprEmpresas.Fields("nomeEmpresa").Value
            CNPJEmpr = trepEmprEmpresas.Fields("CNPJEmpr").Value
            enderecoEmpresa = trepEmprEmpresas.Fields("enderecoEmpresa").Value
            complemEmpresa = trepEmprEmpresas.Fields("complemEmpresa").Value
            bairroEmpresa = trepEmprEmpresas.Fields("bairroEmpresa").Value
            paisEmpresa = trepEmprEmpresas.Fields("paisEmpresa").Value
            UFEmpresa = trepEmprEmpresas.Fields("UFEmpresa").Value
            cidadeEmpresa = trepEmprEmpresas.Fields("cidadeEmpresa").Value
            cCepEmpresa = trepEmprEmpresas.Fields("cCepEmpresa").Value
            internetEmpresa = trepEmprEmpresas.Fields("internetEmpresa").Value
            emailEmpresa = trepEmprEmpresas.Fields("emailEmpresa").Value
            dFone1Empresa = trepEmprEmpresas.Fields("dFone1Empresa").Value
            dFone2Empresa = trepEmprEmpresas.Fields("dFone2Empresa").Value
            dFaxEmpresa = trepEmprEmpresas.Fields("dFaxEmpresa").Value
            status = trepEmprEmpresas.Fields("status").Value
            idHolding = trepEmprEmpresas.Fields("idHolding").Value
            nomeFantasia = trepEmprEmpresas.Fields("nomeFantasia").Value
            nomeTpJuridico = trepEmprEmpresas.Fields("nomeTpJuridico").Value
            nomeRamoAtiv = trepEmprEmpresas.Fields("nomeRamoAtiv").Value
            dtInicio = trepEmprEmpresas.Fields("dtInicio").Value
            dDtFim = trepEmprEmpresas.Fields("dDtFim").Value
            nomeRegFiscal = trepEmprEmpresas.Fields("nomeRegFiscal").Value
            tipoPessoa = trepEmprEmpresas.Fields("tipoPessoa").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprEmpresas - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function LeMissao() As String

        LeMissao = trepEmprEmpresas.Fields("missaoEmpresa").Value

    End Function
End Class
