'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 1  'Persistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "FuncoesGerais"
'Attribute VB_GlobalNameSpace = True
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
'Attribute VB_Description = "Componente Gerais - Super-Classe FuncoesGerais."
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit On
Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum

<ComClass(FuncoesGerais.ClassId, FuncoesGerais.InterfaceId, FuncoesGerais.EventsId)>
Public Class FuncoesGerais

    'Public ... Variáveis aqui

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859005"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329005"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9005"
#End Region

    'local variable(s) to hold property value(s)
    'Private mvaridColaborador As Single 'local copy
    'Private mvaridEmpresa As Integer 'local copy

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection
    Public tPesqReg As ADODB.Recordset

    'Public idEmpresa As Integer
    Private mvaridEmpresa As Integer
    Public Property idEmpresa() As Integer
        Get
            Return mvaridEmpresa
        End Get
        Set(ByVal value As Integer)
            mvaridEmpresa = value
        End Set
    End Property


    Private mvaridColaborador As Single
    Public Property idColaborador() As Single
        Get
            Return mvaridColaborador
        End Get
        Set(ByVal value As Single)
            mvaridColaborador = value
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

    Public Function pesqRegistros(vSelec As String) As Integer
        '* *****************************************
        '* * Pesquisa se a tabela contém registros *
        '* * com a condição especificada em vSelec *
        '* *****************************************

        Dim strConnect As String, strSQL As String
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        On Error GoTo EPesqRegistros

        pesqRegistros = 0   'ReturnCode se não encontrar o registro na tabela

        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        cdSeguranca1 = New SegurancaD.cdSeguranca1
        'Abre o DB:
        strConnect = cdSeguranca1.LeDADOSsys(1)
        Db.Open(strConnect)
        'Cria uma instância de ADODB.Recordset
        tPesqReg = New ADODB.Recordset
        'Abre a tabela:
        tPesqReg.Open(vSelec, Db, adOpenDynamic)

        'Verifica se a tabela possui dados nas condições de vSelec:
        If Not tPesqReg.EOF Then pesqRegistros = 1

EPesqRegistros:
        If Err.Number Then
            If Err.Number = 3705 Then
                tPesqReg.Close()
                tPesqReg.Open(vSelec, Db, adOpenDynamic)
                Resume Next
            Else
                pesqRegistros = 2
                MsgBox("Classe FuncoesGerais - pesqRegistros" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function pesqRegistrosAudiHoras(vSelec As String) As Integer
        '* *****************************************
        '* * Pesquisa se a tabela contém registros *
        '* * com a condição especificada em vSelec *
        '* * Aplicativo AudiHoras                  *
        '* *****************************************

        Dim strConnect As String, strSQL As String
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        On Error GoTo EPesqRegistros

        pesqRegistrosAudiHoras = 0   'ReturnCode se não encontrar o registro na tabela

        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        cdSeguranca1 = New SegurancaD.cdSeguranca1
        'Abre o DB - AudiHoras:
        strConnect = cdSeguranca1.LeDADOSsys(2)
        Db.Open(strConnect)
        'Cria uma instância de ADODB.Recordset
        tPesqReg = New ADODB.Recordset
        'Abre a tabela:
        tPesqReg.Open(vSelec, Db, adOpenDynamic)

        'Verifica se a tabela possui dados nas condições de vSelec:
        If Not tPesqReg.EOF Then pesqRegistrosAudiHoras = 1

EPesqRegistros:
        If Err.Number Then
            If Err.Number = 3705 Then
                tPesqReg.Close()
                tPesqReg.Open(vSelec, Db, adOpenDynamic)
                Resume Next
            Else
                pesqRegistrosAudiHoras = 2
                MsgBox("Classe FuncoesGerais - pesqRegistros" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Integer) As Integer
        '* ****************************************
        '* * Lê sequencialmente, a tabela         *
        '* * OcorAumProm                          *
        '* * Utilizado em frmCargos               *
        '* * Retorna valor de idColaborador       *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tPesqReg.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tPesqReg.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tPesqReg.EOF Then
            idColaborador = tPesqReg.Fields("idColaborador").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe FuncoesGerais - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function dataServ() As Date
        '* **********************************************
        '* Retorna a data do Servidor do Banco de Dados *
        '* **********************************************

        Dim Db As ADODB.Connection
        Dim strConnect As String
        Dim rsTime As ADODB.Recordset
        Dim sSQL As String
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        Db = New ADODB.Connection
        cdSeguranca1 = New SegurancaD.cdSeguranca1
        strConnect = cdSeguranca1.LeDADOSsys(1)
        Db.Open(strConnect)
        rsTime = New ADODB.Recordset
        sSQL = "SELECT getdate() as ServerTime"
        rsTime.Open(sSQL, Db, adOpenDynamic)
        dataServ = rsTime.Fields("ServerTime").Value
    End Function
End Class
