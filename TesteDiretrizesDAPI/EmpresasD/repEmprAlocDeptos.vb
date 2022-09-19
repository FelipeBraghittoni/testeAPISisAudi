Option Explicit On
Public Class repEmprAlocDeptos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850318"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320318"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0318"
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


    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvarnomeEmpresa
        End Get
        Set(value As String)
            mvarnomeEmpresa = value
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


    Private mvaridDepto As Short
    Public Property idDepto() As Short
        Get
            Return mvaridDepto
        End Get
        Set(value As Short)
            mvaridDepto = value
        End Set
    End Property


    Private mvarnomeDepto As String
    Public Property nomeDepto() As String
        Get
            Return mvarnomeDepto
        End Get
        Set(value As String)
            mvarnomeDepto = value
        End Set
    End Property


    Private mvarnomeGerente As String
    Public Property nomeGerente() As String
        Get
            Return mvarnomeGerente
        End Get
        Set(value As String)
            mvarnomeGerente = value
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


    Private mvardFone1Depto As String
    Public Property dFone1Depto() As String
        Get
            Return mvardFone1Depto
        End Get
        Set(value As String)
            mvardFone1Depto = value
        End Set
    End Property


    Private mvarramal1Depto As Short
    Public Property ramal1Depto() As Short
        Get
            Return mvarramal1Depto
        End Get
        Set(value As Short)
            mvarramal1Depto = value
        End Set
    End Property


    Private mvardFone2Depto As String
    Public Property dFone2Depto() As String
        Get
            Return mvardFone2Depto
        End Get
        Set(value As String)
            mvardFone2Depto = value
        End Set
    End Property


    Private mvarramal2Depto As Short
    Public Property ramal2Depto() As Short
        Get
            Return mvarramal2Depto
        End Get
        Set(value As Short)
            mvarramal2Depto = value
        End Set
    End Property


    Private mvardFaxDepto As String
    Public Property dFaxDepto() As String
        Get
            Return mvardFaxDepto
        End Get
        Set(value As String)
            mvardFaxDepto = value
        End Set
    End Property

#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public trepEmprAlocDeptos As ADODB.Recordset
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
        trepEmprAlocDeptos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            trepEmprAlocDeptos.Open("repEmprAlocDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM repEmprAlocDeptos ORDER BY idEmpresa, nomeDepto"
                trepEmprAlocDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepEmprAlocDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprAlocDeptos.Close()

                If tipo = 1 Or tipo = 2 Then
                    trepEmprAlocDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprAlocDeptos.Open("repEmprAlocDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprAlocDeptos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprAlocDeptos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprAlocDeptos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprAlocDeptos.EOF Then
            idEmpresa = trepEmprAlocDeptos.Fields("idEmpresa").Value
            nomeEmpresa = trepEmprAlocDeptos.Fields("nomeEmpresa").Value
            idSite = trepEmprAlocDeptos.Fields("idSite").Value
            nomeSite = trepEmprAlocDeptos.Fields("nomeSite").Value
            idDepto = trepEmprAlocDeptos.Fields("idDepto").Value
            nomeDepto = trepEmprAlocDeptos.Fields("nomeDepto").Value
            nomeGerente = trepEmprAlocDeptos.Fields("nomeGerente").Value
            dtInicio = trepEmprAlocDeptos.Fields("dtInicio").Value
            dDtFim = trepEmprAlocDeptos.Fields("dDtFim").Value
            dFone1Depto = trepEmprAlocDeptos.Fields("dFone1Depto").Value
            ramal1Depto = trepEmprAlocDeptos.Fields("ramal1Depto").Value
            dFone2Depto = trepEmprAlocDeptos.Fields("dFone2Depto").Value
            ramal2Depto = trepEmprAlocDeptos.Fields("ramal2Depto").Value
            dFaxDepto = trepEmprAlocDeptos.Fields("dFaxDepto").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprAlocDeptos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
