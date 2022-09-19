Option Explicit On

Public Class repEmprOcorrenc
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850326"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320326"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0326"
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

    Private mvaridTpOcorrenc As Short
    Public Property idTpOcorrenc() As String
        Get
            Return mvaridTpOcorrenc
        End Get
        Set(value As String)
            mvaridTpOcorrenc = value
        End Set
    End Property

    Private mvarnomeTpOcorrenc As String
    Public Property nomeTpOcorrenc() As String
        Get
            Return mvarnomeTpOcorrenc
        End Get
        Set(value As String)
            mvarnomeTpOcorrenc = value
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

#End Region

#Region "Conexão com banco"
    Public trepEmprOcorrenc As ADODB.Recordset
    Public Db As ADODB.Connection
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
        trepEmprOcorrenc = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            trepEmprOcorrenc.Open("repEmprOcorrenc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                trepEmprOcorrenc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprOcorrenc.Close()

                If tipo = 1 Then
                    trepEmprOcorrenc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprOcorrenc.Open("repEmprOcorrenc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprOcorrenc - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprOcorrenc.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprOcorrenc.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprOcorrenc.EOF Then
            idEmpresa = trepEmprOcorrenc.Fields("idEmpresa").Value
            nomeEmpresa = trepEmprOcorrenc.Fields("nomeEmpresa").Value
            idTpOcorrenc = trepEmprOcorrenc.Fields("idTpOcorrenc").Value
            nomeTpOcorrenc = trepEmprOcorrenc.Fields("nomeTpOcorrenc").Value
            dtOcorrenc = trepEmprOcorrenc.Fields("dtOcorrenc").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprOcorrenc - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String

        carregaMemos = trepEmprOcorrenc.Fields("txtOcorrenc").Value

    End Function
End Class
