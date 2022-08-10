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

<ComClass(TmpErros.ClassId, TmpErros.InterfaceId, TmpErros.EventsId)>
Public Class TmpErros

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859010"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329010"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9010"
#End Region

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection
    Public tTmpErros As ADODB.Recordset

    Private mvardtProc As Date
    Public Property dtProc() As Date
        Get
            Return mvardtProc
        End Get
        Set(ByVal value As Date)
            mvardtProc = value
        End Set
    End Property

    Private mvarPeriodo As String
    Public Property Periodo() As String
        Get
            Return mvarPeriodo
        End Get
        Set(ByVal value As String)
            mvarPeriodo = value
        End Set
    End Property

    Private mvarTpProblema As String
    Public Property TpProblema() As String
        Get
            Return mvarTpProblema
        End Get
        Set(ByVal value As String)
            mvarTpProblema = value
        End Set
    End Property

    Private mvarDescricao As String
    Public Property Descricao() As String
        Get
            Return mvarDescricao
        End Get
        Set(ByVal value As String)
            mvarDescricao = value
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

    Public Function DbConecta(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset                      *
        '* ****************************************

        On Error GoTo EdbConecta

        DbConecta = 0   'ReturnCode se não houver nenhum problema

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
        tTmpErros = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tTmpErros.Open("TmpRDErros", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                DbConecta = 1014
            Else
                tTmpErros.Open(vSelec, Db, adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = -2147217865 Then
                DbConecta = 8888
                Exit Function
            End If
            If Err.Number = 3705 Then
                tTmpErros.Close()

                If tipo = 1 Then
                    tTmpErros.Open(vSelec, Db, adOpenDynamic)
                Else
                    tTmpErros.Open("TmpRDErros", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
                End If
                Resume Next
            Else
                DbConecta = Err.Number
                If Err.Number = 3024 Then DbConecta = 1046
                MsgBox("Classe TmpErros - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Sub DbDesConecta()

        tTmpErros.Close()

    End Sub

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
        If Not tTmpErros.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tTmpErros.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tTmpErros.EOF Then
            dtProc = tTmpErros.Fields("dtProc").Value
            Periodo = tTmpErros.Fields("Periodo").Value
            TpProblema = tTmpErros.Fields("TpProblema").Value
            Descricao = tTmpErros.Fields("Descricao").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe RDTmpErros - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class




