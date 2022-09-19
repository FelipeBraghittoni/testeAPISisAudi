Option Explicit On
Public Class repProjOcor
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850737"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320737"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0737"
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

    Private mvardescricFatos As String
    Public Property descricFatos() As String
        Get
            Return mvardescricFatos
        End Get
        Set(value As String)
            mvardescricFatos = value
        End Set
    End Property

    Private mvardataOcorrenc As String
    Public Property dataOcorrenc() As String
        Get
            Return mvardataOcorrenc
        End Get
        Set(value As String)
            mvardataOcorrenc = value
        End Set
    End Property

    Private mvardataSolucao As String
    Public Property dataSolucao() As String
        Get
            Return mvardataSolucao
        End Get
        Set(value As String)
            mvardataSolucao = value
        End Set
    End Property

    Private mvarnomeColab As String
    Public Property nomeColab() As String
        Get
            Return mvarnomeColab
        End Get
        Set(value As String)
            mvarnomeColab = value
        End Set
    End Property

    Private mvardtOcorrenc As Date
    Public Property dtOcorrenc() As Date
        Get
            Return mvardtOcorrenc
        End Get
        Set(value As Date)
            mvardtOcorrenc = value
        End Set
    End Property

    Private mvardtSolucao As Date
    Public Property dtSolucao() As Date
        Get
            Return mvardtSolucao
        End Get
        Set(value As Date)
            mvardtSolucao = value
        End Set
    End Property

    Private mvaridColabFato As Single
    Public Property idColabFato() As Single
        Get
            Return mvaridColabFato
        End Get
        Set(value As Single)
            mvaridColabFato = value
        End Set
    End Property

#End Region

#Region "Conexão de banco"
    Public Db As ADODB.Connection
    Public trepProjOcor As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjOcor         *
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
        trepProjOcor = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjOcor como OpenTable
                trepProjOcor.Open("repProjOcor", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjOcor.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjOcor.Close()

                Select Case tipo
                    Case 0      'Abre repProjOcor como OpenTable
                        trepProjOcor.Open("repProjOcor", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic,)
                    Case 1      'Abre como OpenDynaset
                        trepProjOcor.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjOcor - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjOcor.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjOcor.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjOcor.EOF Then
            idProjeto = trepProjOcor.Fields("idProjeto").Value
            nomeProjeto = trepProjOcor.Fields("nomeProjeto").Value
            dataOcorrenc = trepProjOcor.Fields("dataOcorrenc").Value
            dataSolucao = trepProjOcor.Fields("dataSolucao").Value
            descricFatos = trepProjOcor.Fields("descricFatos").Value
            nomeColab = trepProjOcor.Fields("nomeColab").Value
            dtOcorrenc = trepProjOcor.Fields("dtGeracaoFato").Value
            dtSolucao = trepProjOcor.Fields("dtSolucaoFato").Value
            idColabFato = trepProjOcor.Fields("idColabFato").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjOcor - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaTxtGeracao() As String

        CarregaTxtGeracao = trepProjOcor.Fields("txtGeracaoFato").Value

    End Function

    Public Function CarregaTxtSolucao() As String

        CarregaTxtSolucao = trepProjOcor.Fields("txtSolucaoFato").Value

    End Function
End Class
