'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "segServicos"
'Attribute VB_GlobalNameSpace = True
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit On
Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum
Imports ADODB
Imports Microsoft.VisualBasic.FileIO
<ComClass(segServicos.ClassId, segServicos.InterfaceId, segServicos.EventsId)>
Public Class segServicos

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859008"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329008"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9008"
#End Region

    'local variable(s) to hold property value(s)
    'Private mvarnumServico As Integer 'local copy
    'Private mvarcodSistema As Single 'local copy
    'Private mvarcodModulo As Single 'local copy
    'Private mvarcodSubModulo As Single 'local copy
    'Private mvardescricServico As String 'local copy
    'Private mvarnumServReorg As Integer 'local copy

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection

    Public tServicos As ADODB.Recordset

    Private mvarnumServReorg As Integer
    Public Property numServReorg() As Integer
        Get
            Return mvarnumServReorg
        End Get
        Set(ByVal value As Integer)
            mvarnumServReorg = value
        End Set
    End Property

    Private mvardescricServico As String
    Public Property descricServico() As String
        Get
            Return mvardescricServico
        End Get
        Set(ByVal value As String)
            mvardescricServico = value
        End Set
    End Property

    Private mvarcodSubModulo As Single
    Public Property codSubModulo() As Single
        Get
            Return mvarcodSubModulo
        End Get
        Set(ByVal value As Single)
            mvarcodSubModulo = value
        End Set
    End Property

    Private mvarcodModulo As Single
    Public Property codModulo() As Single
        Get
            Return mvarcodModulo
        End Get
        Set(ByVal value As Single)
            mvarcodModulo = value
        End Set
    End Property

    Private mvarcodSistema As Single
    Public Property codSistema() As Single
        Get
            Return mvarcodSistema
        End Get
        Set(ByVal value As Single)
            mvarcodSistema = value
        End Set
    End Property

    Private mvarnumServico As Integer
    Public Property numServico() As Integer
        Get
            Return mvarnumServico
        End Get
        Set(ByVal value As Integer)
            mvarnumServico = value
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
        '* * 1 = OpenDynaset                      *
        '* ****************************************

        Dim strConnect As String, strSQL As String, vSelect As String = ""

        On Error GoTo EdbConecta

        dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            'Cria uma instância de ADODB.Connection:
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tServicos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            tServicos.Open("Servicos", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tServicos.Open(vSelec, Db, adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tServicos.Close()
                tServicos.Open(vSelect, Db, adOpenDynamic)
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe Servicos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function RetornaNumServico(Num As Integer) As String
        '* ********************************************************
        '* Busca um serviço específico (Num), na tabela tServiço. *
        '* Utiliza o índice por Num (IndCodServico).
        '* Retorna composição do serviço pesquisado
        '* ********************************************************

        Dim vSelect As String
        On Error GoTo ERetorna
        vSelect = "SELECT * FROM Servicos WHERE numServico = " & Num

        tServicos.Open(vSelect, Db, adOpenDynamic)

        'tServicos.Seek Num, adSeekFirstEQ
        If tServicos.EOF Then       'Não encontrou o servico
            RetornaNumServico = ""
            Exit Function
        Else
            RetornaNumServico = Convert.ToInt32(tServicos(1).Value).ToString("00")  'codSistema
            RetornaNumServico = RetornaNumServico & Convert.ToInt32(tServicos(2).Value).ToString("00") 'codModulo
            RetornaNumServico = RetornaNumServico & Convert.ToInt32(tServicos(3).Value).ToString("00")    'codSubSistema
            'Para controlar níveis de agrupamento da TreeView:
            If tServicos(3).Value <> 0 Then   'codSubModulo
                RetornaNumServico = RetornaNumServico & "3"
            Else
                If tServicos(2).Value <> 0 Then 'codModulo
                    RetornaNumServico = RetornaNumServico & "2"
                Else                                  'codSistema
                    RetornaNumServico = RetornaNumServico & "1"
                End If
            End If
            RetornaNumServico = RetornaNumServico & Convert.ToInt32(tServicos(0).Value).ToString("000")   'Número do serviço
            RetornaNumServico = RetornaNumServico & tServicos(4).Value                  'Descrição do serviço
        End If

ERetorna:
        If Err.Number Then
            If Err.Number = 3705 Then
                tServicos.Close()
                tServicos.Open(vSelect, Db, adOpenDynamic)
            End If
            Resume Next
        End If

    End Function

    Public Function localizaNumServico(Num As Integer) As Integer
        '* ********************************************************
        '* Busca um serviço específico (Num), na tabela tServiço. *
        '* Utiliza o índice por Num (IndCodServico).              *
        '* ********************************************************

        Dim vSelect As String
        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM Servicos WHERE numServico = " & Num

        tServicos.Open(vSelect, Db, adOpenDynamic)

        'tServicos.Seek Num, adSeekFirstEQ
        If tServicos.EOF Then       'Não encontrou o servico
            localizaNumServico = 1013
            Exit Function
        Else
            localizaNumServico = 0
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tServicos.Close()
                tServicos.Open(vSelect, Db, adOpenDynamic)
            End If
            Resume Next
        End If

    End Function

    Public Function localizaClasServico(cSistema As Integer, cModulo As Integer, cSubModulo As Integer) As Integer
        '* ********************************************************************
        '* Busca um serviço específico (cSistema+cModulo+cSubModulo),         *
        '* na tabela tServiço.                                                *
        '* Utiliza o índice por csistema+cModulo+cSubModulo (IndClasServico). *
        '* ********************************************************************

        Dim vSelect As String
        On Error GoTo ElocalizaClasServico
        vSelect = "SELECT * FROM Servicos WHERE codSistema = " & cSistema
        vSelect = vSelect & " AND codModulo = " & cModulo
        vSelect = vSelect & " AND codSubModulo = " & cSubModulo

        tServicos.Open(vSelect, Db, adOpenDynamic)

        'tServicos.Seek Array(cSistema, cModulo, cSubModulo), adSeekFirstEQ
        If tServicos.EOF Then       'Não encontrou o servico
            localizaClasServico = 1013
            Exit Function
        Else
            localizaClasServico = 0
        End If

ElocalizaClasServico:
        If Err.Number Then
            If Err.Number = 3705 Then
                tServicos.Close()
                tServicos.Open(vSelect, Db, adOpenDynamic)
            End If
            Resume Next
        End If

    End Function

    Public Function leSeq(tipo As Integer) As Integer
        '* *****************************
        '* Lê sequencialmente tServico *
        '* tipo = tipo de leitura:     *
        '* 0 = primeira vez            *
        '* 1 = lê o próximo registro   *
        '* *****************************

        On Error GoTo EleSeq
        leSeq = 0    'ReturnCode se não houver nenhum problema
        If Not tServicos.EOF Then
            If tipo = 0 Then    'Não é a primeira leitura
                tServicos.MoveNext()
            End If
        End If
        If tServicos.EOF Then
            leSeq = 1016
            Exit Function
        Else
            numServico = tServicos.Fields("numServico").Value
            codSistema = tServicos.Fields("codSistema").Value
            codModulo = tServicos.Fields("codModulo").Value
            codSubModulo = tServicos.Fields("codSubModulo").Value
            descricServico = tServicos.Fields("descrServico").Value
            numServReorg = tServicos.Fields("numServReorg").Value
        End If

EleSeq:
        If Err.Number Then leSeq = Str(Err.Number)

    End Function

    Public Function leServicoSeq(tipo As Integer) As String
        '* *****************************
        '* Lê sequencialmente tServico *
        '* tipo = tipo de leitura:     *
        '* 0 = primeira vez            *
        '* 1 = lê o próximo registro   *
        '* *****************************

        On Error GoTo EleServicoSeq
        leServicoSeq = 0    'ReturnCode se não houver nenhum problema
        If Not tServicos.EOF Then
            If tipo <> 0 Then    'Não é a primeira leitura
                tServicos.MoveNext()
            End If
        End If
        If tServicos.EOF Then
            leServicoSeq = ""
            Exit Function
        Else
            leServicoSeq = Convert.ToInt32(tServicos(1).Value).ToString("00") & Convert.ToInt32(tServicos(2).Value).ToString("00") & Convert.ToInt32(tServicos(3).Value).ToString("00")
            leServicoSeq = leServicoSeq & Convert.ToInt32(tServicos(0).Value).ToString("000") & tServicos(4).Value
        End If

EleServicoSeq:
        If Err.Number Then leServicoSeq = Str(Err.Number)

    End Function

    Public Sub mudaIndiceServ(tipo As Integer)
        'Altera o tipo de índice a ser utilizado na tabela de serviços:

        If tipo = 0 Then
            'tServicos.Index = "indClasServico"
        Else
            'tServicos.Index = "indCodServico"
        End If

    End Sub

    Public Function criaServico(cNumero As Integer, cCodSistema As Integer, cCodModulo As Integer, cCodSubModulo As Integer, cDescServico As String) As Integer

        On Error GoTo EcriaServico
        tServicos.AddNew()
        tServicos(0).Value = cNumero
        tServicos(1).Value = cCodSistema
        tServicos(2).Value = cCodModulo
        tServicos(3).Value = cCodSubModulo
        tServicos(4).Value = cDescServico
        tServicos.Fields("numServReorg").Value = numServReorg
        tServicos.Update()
        criaServico = 0

EcriaServico:
        If Err.Number Then criaServico = Err.Number

    End Function

    Public Function alteraServico(cNumero As Integer, cCodSistema As Integer, cCodModulo As Integer, cCodSubModulo As Integer, cDescServico As String) As Integer

        On Error GoTo EalteraServico
        tServicos(0).Value = cNumero
        tServicos(1).Value = cCodSistema
        tServicos(2).Value = cCodModulo
        tServicos(3).Value = cCodSubModulo
        tServicos(4).Value = cDescServico
        tServicos.Fields("numServReorg").Value = numServReorg
        tServicos.Update()
        alteraServico = 0

EalteraServico:
        If Err.Number Then
            alteraServico = Err.Number
            MsgBox(Err.Number & " - " & Err.Description & Chr(10) & "Serviço: " & cNumero)
        End If

    End Function
End Class
