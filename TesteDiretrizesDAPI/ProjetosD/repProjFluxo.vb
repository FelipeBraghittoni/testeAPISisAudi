Option Explicit On
Public Class repProjFluxo
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850735"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320735"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0735"
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

    Private mvarnomeColab As String
    Public Property nomeColab() As String
        Get
            Return mvarnomeColab
        End Get
        Set(value As String)
            mvarnomeColab = value
        End Set
    End Property

    Private mvardataStatus As String
    Public Property dataStatus() As String
        Get
            Return mvardataStatus
        End Get
        Set(value As String)
            mvardataStatus = value
        End Set
    End Property

    Private mvarnomeStatus As String
    Public Property nomeStatus() As String
        Get
            Return mvarnomeStatus
        End Get
        Set(value As String)
            mvarnomeStatus = value
        End Set
    End Property

    Private mvardtStatus As Date
    Public Property dtStatus() As Date
        Get
            Return mvardtStatus
        End Get
        Set(value As Date)
            mvardtStatus = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"

    Public Db As ADODB.Connection
    Public trepProjFluxo As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjFluxo        *
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
        trepProjFluxo = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjFluxo como OpenTable
                trepProjFluxo.Open("repProjFluxo", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjFluxo.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjFluxo.Close()

                Select Case tipo
                    Case 0      'Abre repProjFluxo como OpenTable
                        trepProjFluxo.Open("repProjFluxo", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjFluxo.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjFluxo - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjFluxo.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjFluxo.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjFluxo.EOF Then
            idProjeto = trepProjFluxo.Fields("idProjeto").Value
            nomeProjeto = trepProjFluxo.Fields("nomeProjeto").Value
            dataStatus = trepProjFluxo.Fields("dataStatus").Value
            nomeStatus = trepProjFluxo.Fields("nomeStatus").Value
            nomeColab = trepProjFluxo.Fields("nomeColab").Value
            dtStatus = trepProjFluxo.Fields("dtStatus").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjFluxo - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaTxt() As String

        CarregaTxt = trepProjFluxo.Fields("txtStatus").Value

    End Function
End Class
