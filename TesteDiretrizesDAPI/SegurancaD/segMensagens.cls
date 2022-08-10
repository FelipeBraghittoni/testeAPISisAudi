Option Explicit On

Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum

<ComClass(segMensagens.ClassId, segMensagens.InterfaceId, segMensagens.EventsId)>
Public Class segMensagens

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859007"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329007"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9007"
#End Region



    Public vCurDir As String
    'local variable(s) to hold property value(s)
    'Private mvartipoMsg As Integer 'local copy
    'Private mvarcodigoMsg As Integer 'local copy
    'Private mvardescricMsg As String 'local copy

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection
    Public tMensagens As ADODB.Recordset

    Private mvardescricMsg As String
    Public Property descricMsg() As String
        Get
            Return mvardescricMsg
        End Get
        Set(ByVal value As String)
            mvardescricMsg = value
        End Set
    End Property

    Private mvarcodigoMsg As Integer
    Public Property codigoMsg() As Integer
        Get
            Return mvarcodigoMsg
        End Get
        Set(ByVal value As Integer)
            mvarcodigoMsg = value
        End Set
    End Property

    Private mvartipoMsg As Integer
    Public Property tipoMsg() As Integer
        Get
            Return mvartipoMsg
        End Get
        Set(ByVal value As Integer)
            mvartipoMsg = value
        End Set
    End Property
    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    '###########################

    Public Function dbConecta(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset             *
        '* ****************************************

        Dim strConnect As String, strSQL As String

        On Error GoTo EdbConecta

        dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            'Cria uma instância de ADODB.Connection:
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            'strConnect = "driver={SQL Server};server=localhost\SQLEXPRESS;database=auditeste;uid=sa;pwd=auditeste"
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tMensagens = New ADODB.Recordset
        'Abre a tabela:
        Dim vSelect As String = ""
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            tMensagens.Open("Mensagens", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tMensagens.Open(vSelec, Db, adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tMensagens.Close()
                tMensagens.Open(vSelect, Db, adOpenDynamic)
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                'MsgBox("Classe Mensagens - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
                Return ("Classe Mensagens - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Integer) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tMensagens.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tMensagens.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tMensagens.EOF Then
            tipoMsg = tMensagens.Fields("tipoMsg").Value
            codigoMsg = tMensagens.Fields("codigoMsg").Value
            descricMsg = tMensagens.Fields("descricMsg").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe Mensagens - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(pCodigoMsg As String, atualiz As Integer) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * Mensagens.                           *
        '* *                                      *
        '* * cliente    = identif. para pesquisa  *
        '* * site       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************

        Dim vSelect As String

        On Error GoTo Elocaliza

        localiza = 0    'ReturnCode se não houver nenhum problema
        vSelect = "SELECT * FROM Mensagens WHERE codigoMsg = '" & pCodigoMsg & "'"
        tMensagens.Open(vSelect, Db, adOpenDynamic)

        'tMensagens.Seek pCodigoMsg, adSeekFirstEQ
        'Não encontrou:
        If tMensagens.EOF Then
            localiza = 1021
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            tipoMsg = tMensagens.Fields("tipoMsg").Value
            codigoMsg = tMensagens.Fields("codigoMsg").Value
            descricMsg = tMensagens.Fields("descricMsg").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tMensagens.Close()
                tMensagens.Open(vSelect, Db, adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe Mensagens - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cTipo As Integer, cCodigo As Integer, cMsg As String) As Integer

        On Error GoTo Einclui
        tMensagens.AddNew()
        tMensagens(0).Value = cTipo
        tMensagens(1).Value = cCodigo
        tMensagens(2).Value = cMsg
        tMensagens.Update()
        inclui = 0

Einclui:
        If Err.Number Then inclui = Err.Number

    End Function

    Public Function altera(cTipo As Integer, cCodigo As Integer, cMsg As String) As Integer

        On Error GoTo Ealtera
        tMensagens(2).Value = cMsg
        tMensagens.Update()
        altera = 0

Ealtera:
        If Err.Number Then altera = Err.Number

    End Function

    Public Sub exibeMsg(Optional ptipoMsg As Integer = 0, Optional ByVal pCodigoMsg As Integer = 0)
        'Exibe mensagem para o usuário

        'Variável para guardar o título:
        Dim vTitulo As String
        Dim vSelect As String

        On Error GoTo EexibeMsg

        'Pesquisa a mensagem na tabela:
        If ptipoMsg <> 0 Then   'Foi informado o tipo de mensagem
            vSelect = "SELECT * FROM Mensagens WHERE codigoMsg = " & pCodigoMsg
            vSelect = vSelect & " ORDER BY TipoMsg"
        Else                    'Não foi informado o tipo de mensagem
            vSelect = "SELECT * FROM Mensagens WHERE codigoMsg = " & pCodigoMsg
        End If
        tMensagens.Open(vSelect, Db, adOpenDynamic)
        'Verifica se encontrou na tabela:
        If Not tMensagens.EOF Then  'Encontrou a Mensagem
            If ptipoMsg = 0 Then ptipoMsg = tMensagens(0).Value
            If ptipoMsg = 16 Then
                vTitulo = "Erro!"
            Else
                If ptipoMsg = 48 Then
                    vTitulo = "Atenção!"
                Else
                    vTitulo = "Aviso!"
                End If
            End If
            'Exibe a mensagen:
            MsgBox(tMensagens(2).Value, ptipoMsg, vTitulo)
        Else        'Se não encontrou a mensagem
            vTitulo = "Código da Mensagem = " & pCodigoMsg & vbCrLf
            ''''vTitulo = vTitulo & Error(pCodigoMsg) & vbCrLf 'linha removida Claudio
            vTitulo = vTitulo & vbCrLf & "Favor informar o código da mensagem ao suporte técnico!"
            MsgBox(vTitulo, 48, "Atenção!")
        End If

EexibeMsg:
        If Err.Number Then
            If Err.Number = 3705 Then
                tMensagens.Close()
                tMensagens.Open(vSelect, Db, adOpenDynamic)
                Resume Next
            Else
                MsgBox("Classe Mensagens - exibeMsg" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Sub
End Class
