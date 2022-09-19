Option Explicit On
Public Class repProjAlocHoras
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850731"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320731"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0731"
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

    Private mvarnomeProjeto As String
    Public Property nomeProjeto() As String
        Get
            Return mvarnomeProjeto
        End Get
        Set(value As String)
            mvarnomeProjeto = value
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

    Private mvarnomeAtividade As String
    Public Property nomeAtividade() As String
        Get
            Return mvaridAtividade
        End Get
        Set(value As String)
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

    Private mvarnomeSAtividade As String
    Public Property nomeSAtividade() As String
        Get
            Return mvarnomeSAtividade
        End Get
        Set(value As String)
            mvarnomeSAtividade = value
        End Set
    End Property

    Private mvaridPropostaExt As String
    Public Property idPropostaExt() As String
        Get
            Return mvaridPropostaExt
        End Get
        Set(value As String)
            mvaridPropostaExt = value
        End Set
    End Property

    Private mvardataInicio As String
    Public Property dataInicio() As String
        Get
            Return mvardataInicio
        End Get
        Set(value As String)
            mvardataInicio = value
        End Set
    End Property

    Private mvardataFim As String
    Public Property dataFim() As String
        Get
            Return mvardataFim
        End Get
        Set(value As String)
            mvardataFim = value
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
    Public Property idusuario() As String
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
    Public Property dcHoras() As Date
        Get
            Return mvardcHoras
        End Get
        Set(value As Date)
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
    Public trepProjAlocHoras As ADODB.Recordset
#End Region


    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjAlocHoras    *
        '* * 1 = OpenDynaset                   *
        '* *************************************

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
        trepProjAlocHoras = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjAlocHoras como OpenTable
                trepProjAlocHoras.Open("repProjAlocHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjAlocHoras.Close()

                Select Case tipo
                    Case 0      'Abre repProjAlocHoras como OpenTable
                        trepProjAlocHoras.Open("repProjAlocHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjAlocHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjAlocHoras - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjAlocHoras.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjAlocHoras.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjAlocHoras.EOF Then
            idProjeto = trepProjAlocHoras.Fields("idProjeto").Value
            nomeProjeto = trepProjAlocHoras.Fields("nomeProjeto").Value
            dataInicio = trepProjAlocHoras.Fields("dataInicio").Value
            dataFim = trepProjAlocHoras.Fields("dataFim").Value
            idAtividade = trepProjAlocHoras.Fields("idAtividade").Value
            nomeAtividade = trepProjAlocHoras.Fields("nomeAtividade").Value
            idSAtividade = trepProjAlocHoras.Fields("idSAtividade").Value
            nomeSAtividade = trepProjAlocHoras.Fields("nomeSAtividade").Value
            idPropostaExt = trepProjAlocHoras.Fields("idPropostaExt").Value
            dtInicio = trepProjAlocHoras.Fields("dtInicio").Value
            dtFim = trepProjAlocHoras.Fields("dtFim").Value
            horas = trepProjAlocHoras.Fields("horas").Value
            idusuario = trepProjAlocHoras.Fields("idUsuario").Value
            dtHsAlteracao = trepProjAlocHoras.Fields("dtHsAlteracao").Value
            dcHoras = trepProjAlocHoras.Fields("dcHoras").Value
            origem = trepProjAlocHoras.Fields("origem").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjAlocHoras - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaTxt() As String

        CarregaTxt = trepProjAlocHoras.Fields("descricAlocHoras").Value

    End Function

End Class
