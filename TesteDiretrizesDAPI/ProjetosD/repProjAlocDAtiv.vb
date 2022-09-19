Option Explicit On
Public Class repProjAlocDAtiv
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850730"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320730"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0730"
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

    Private mvarnomeAtividade As String
    Public Property nomeAtividade() As String
        Get
            Return mvarnomeAtividade
        End Get
        Set(value As String)
            mvarnomeAtividade = value
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

    Private mvarnomeSAtividade As String
    Public Property nomeSAtividade() As String
        Get
            Return mvarnomeSAtividade
        End Get
        Set(value As String)
            mvarnomeSAtividade = value
        End Set
    End Property

    Private mvarnomeDetAtiv As String
    Public Property nomeDetAtiv() As String
        Get
            Return mvarnomeDetAtiv
        End Get
        Set(value As String)
            mvarnomeDetAtiv = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public trepProjAlocDAtiv As ADODB.Recordset

#End Region


    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjAlocDAtiv    *
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
        trepProjAlocDAtiv = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjAlocDAtiv como OpenTable
                trepProjAlocDAtiv.Open("repProjAlocDAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjAlocDAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjAlocDAtiv.Close()

                Select Case tipo
                    Case 0      'Abre repProjAlocDAtiv como OpenTable
                        trepProjAlocDAtiv.Open("repProjAlocDAtiv", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjAlocDAtiv.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjAlocDAtiv - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjAlocDAtiv.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjAlocDAtiv.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjAlocDAtiv.EOF Then
            idProjeto = trepProjAlocDAtiv.Fields("idProjeto").Value
            nomeProjeto = trepProjAlocDAtiv.Fields("nomeProjeto").Value
            dataInicio = trepProjAlocDAtiv.Fields("dataInicio").Value
            dataFim = trepProjAlocDAtiv.Fields("dataFim").Value
            nomeAtividade = trepProjAlocDAtiv.Fields("nomeAtividade").Value
            nomeSAtividade = trepProjAlocDAtiv.Fields("nomeSAtividade").Value
            nomeDetAtiv = trepProjAlocDAtiv.Fields("nomeDetAtiv").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjAlocDAtiv - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
