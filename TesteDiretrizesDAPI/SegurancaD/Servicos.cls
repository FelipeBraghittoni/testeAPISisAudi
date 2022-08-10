'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "Servicos"
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
Imports ADODB.SeekEnum
Imports ADODB
Imports Microsoft.VisualBasic.FileIO
Imports System.IO

<ComClass(Servicos.ClassId, Servicos.InterfaceId, Servicos.EventsId)>
Public Class Servicos

    'Public Db As ADODB.Connection
    'Public tDirTpDiretriz As ADODB.Recordset

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859004"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329004"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9004"
#End Region

    ''local variable(s) to hold property value(s)
    'Private mvarnumServico As Integer 'local copy
    'Private mvarcodGrupo As Integer 'local copy
    'Private mvarcodModulo As Integer 'local copy
    'Private mvarcodSubModulo As Integer 'local copy
    'Private mvardescricServico As String 'local copy

    ''local variable(s) to hold property value(s)
    'Private mvarcodAtividade As Integer 'local copy
    'Private mvartipoUsuario As Integer 'local copy
    'Private mvarnivelAgrup As Integer 'local copy
    'Private mvarnumFigura As Integer 'local copy
    'Private mvarnumServico_Novo As Integer 'local copy
    'Private mvarcodMódulo_Novo As Integer 'local copy
    'Private mvarcodSubModulo_Novo As Integer 'local copy
    'Private mvarcodGrupo_Novo As Integer 'local copy
    'Private mvarcodAtividade_Novo As Integer 'local copy

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection
    Public tServicos As ADODB.Recordset

    Private mvarcodAtividade_Novo As Integer
    Public Property codAtividade_Novo() As Integer
        Get
            Return mvarcodAtividade_Novo
        End Get
        Set(ByVal value As Integer)
            mvarcodAtividade_Novo = value
        End Set
    End Property

    Private mvarcodGrupo_Novo As Integer
    Public Property codGrupo_Novo() As Integer
        Get
            Return mvarcodGrupo_Novo
        End Get
        Set(ByVal value As Integer)
            mvarcodGrupo_Novo = value
        End Set
    End Property

    Private mvarcodSubModulo_Novo As Integer
    Public Property codSubModulo_Novo() As Integer
        Get
            Return mvarcodSubModulo_Novo
        End Get
        Set(ByVal value As Integer)
            mvarcodSubModulo_Novo = value
        End Set
    End Property

    Private mvarcodMódulo_Novo As Integer
    Public Property codMódulo_Novo() As Integer
        Get
            Return mvarcodMódulo_Novo
        End Get
        Set(ByVal value As Integer)
            mvarcodMódulo_Novo = value
        End Set
    End Property

    Private mvarnumServico_Novo As Integer
    Public Property numServico_Novo() As Integer
        Get
            Return mvarnumServico_Novo
        End Get
        Set(ByVal value As Integer)
            mvarnumServico_Novo = value
        End Set
    End Property

    Private mvarnumFigura As Integer
    Public Property numFigura() As Integer
        Get
            Return mvarnumFigura
        End Get
        Set(ByVal value As Integer)
            mvarnumFigura = value
        End Set
    End Property

    Private mvarnivelAgrup As Integer
    Public Property nivelAgrup() As Integer
        Get
            Return mvarnivelAgrup
        End Get
        Set(ByVal value As Integer)
            mvarnivelAgrup = value
        End Set
    End Property

    Private mvartipoUsuario As Integer
    Public Property tipoUsuario() As Integer
        Get
            Return mvartipoUsuario
        End Get
        Set(ByVal value As Integer)
            mvartipoUsuario = value
        End Set
    End Property

    Private mvarcodAtividade As Integer
    Public Property codAtividade() As Integer
        Get
            Return mvarcodAtividade
        End Get
        Set(ByVal value As Integer)
            mvarcodAtividade = value
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

    Private mvarcodSubModulo As Integer
    Public Property codSubModulo() As Integer
        Get
            Return mvarcodSubModulo
        End Get
        Set(ByVal value As Integer)
            mvarcodSubModulo = value
        End Set
    End Property

    Private mvarcodModulo As Integer
    Public Property codModulo() As Integer
        Get
            Return mvarcodModulo
        End Get
        Set(ByVal value As Integer)
            mvarcodModulo = value
        End Set
    End Property

    Private mvarcodGrupo As Integer
    Public Property codGrupo() As Integer
        Get
            Return mvarcodGrupo
        End Get
        Set(ByVal value As Integer)
            mvarcodGrupo = value
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

        Dim strConnect As String

        On Error GoTo EdbConecta

        dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            'Cria uma instância de ADODB.Connection:
            Db = New ADODB.Connection
            '* ********************************************
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            '* ********************************************
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
                tServicos.Open(vSelec, Db, adOpenDynamic)
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe Servicos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNumServico(Num As Integer) As Integer
        '* ********************************************************
        '* Busca um serviço específico (Num), na tabela tServiço. *
        '* Utiliza o índice por Num (IndCodServico).              *
        '* ********************************************************

        tServicos.Seek(Num, adSeekFirstEQ)
        If tServicos.EOF Then       'Não encontrou o servico
            localizaNumServico = 1013
            Exit Function
        Else
            localizaNumServico = 0
        End If

    End Function

    Public Function localizaClasServico(cSistema As Integer, cModulo As Integer, cSubModulo As Integer) As Integer
        '* ********************************************************************
        '* Busca um serviço específico (cSistema+cModulo+cSubModulo),         *
        '* na tabela tServiço.                                                *
        '* Utiliza o índice por csistema+cModulo+cSubModulo (IndClasServico). *
        '* ********************************************************************
        Dim Arr As Integer() = {cSistema, cModulo, cSubModulo}
        tServicos.Seek(Arr, adSeekFirstEQ)
        'tServicos.Seek(Array(cSistema, cModulo, cSubModulo), adSeekFirstEQ)
        If tServicos.EOF Then       'Não encontrou o servico
            localizaClasServico = 1013
            Exit Function
        Else
            localizaClasServico = 0
        End If

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

    Public Function criaServico(cNumero As Integer, ccodGrupo As Integer, cCodModulo As Integer, cCodSubModulo As Integer, cDescServico As String) As Integer

        On Error GoTo EcriaServico
        tServicos.AddNew()
        tServicos(0).Value = cNumero
        tServicos(1).Value = ccodGrupo
        tServicos(2).Value = cCodModulo
        tServicos(3).Value = cCodSubModulo
        tServicos(4).Value = cDescServico
        tServicos.Update()
        criaServico = 0

EcriaServico:
        If Err.Number Then criaServico = Err.Number

    End Function

    Public Function alteraServico(cNumero As Integer, ccodGrupo As Integer, cCodModulo As Integer, cCodSubModulo As Integer, cDescServico As String) As Integer

        On Error GoTo EalteraServico
        tServicos(0).Value = cNumero
        tServicos(1).Value = ccodGrupo
        tServicos(2).Value = cCodModulo
        tServicos(3).Value = cCodSubModulo
        tServicos(4).Value = cDescServico
        tServicos.Update()
        alteraServico = 0

EalteraServico:
        If Err.Number Then alteraServico = Err.Number

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
        If Not tServicos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tServicos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tServicos.EOF Then
            numServico = tServicos.Fields("numServico").Value
            codModulo = tServicos.Fields("codMódulo").Value
            codSubModulo = tServicos.Fields("codSubModulo").Value
            codGrupo = tServicos.Fields("codGrupo").Value
            codAtividade = tServicos.Fields("codAtividade").Value
            descricServico = tServicos.Fields("descrServico").Value
            tipoUsuario = tServicos.Fields("tipoUsuario").Value
            nivelAgrup = tServicos.Fields("nivelAgrup").Value
            numFigura = tServicos.Fields("numFigura").Value
            numServico_Novo = tServicos.Fields("numServico_Novo").Value
            codMódulo_Novo = tServicos.Fields("codMódulo_Novo").Value
            codSubModulo_Novo = tServicos.Fields("codSubModulo_Novo").Value
            codGrupo_Novo = tServicos.Fields("codGrupo_Novo").Value
            codAtividade_Novo = tServicos.Fields("codAtividade_Novo").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe Servicos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
