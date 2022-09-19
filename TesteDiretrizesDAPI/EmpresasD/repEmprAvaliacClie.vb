﻿Option Explicit On
Public Class repEmprAvaliacClie
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850319"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320319"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0319"
#End Region

#Region "Variaveis de Ambiente"
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


    Private mvaridCliente As Short
    Public Property idCliente() As Short
        Get
            Return mvaridCliente
        End Get
        Set(value As Short)
            mvaridCliente = value
        End Set
    End Property


    Private mvarnomeCliente As String
    Public Property nomeCliente() As String
        Get
            Return mvarnomeCliente
        End Get
        Set(value As String)
            mvarnomeCliente = value
        End Set
    End Property


    Private mvarnomeAvaliador As String
    Public Property nomeAvaliador() As String
        Get
            Return mvarnomeAvaliador
        End Get
        Set(value As String)
            mvarnomeAvaliador = value
        End Set
    End Property


    Private mvarcargoAvaliador As String
    Public Property cargoAvaliador() As String
        Get
            Return mvarcargoAvaliador
        End Get
        Set(value As String)
            mvarcargoAvaliador = value
        End Set
    End Property


    Private mvardtAvaliacao As Date
    Public Property dtAvaliacao() As Date
        Get
            Return mvardtAvaliacao
        End Get
        Set(value As Date)
            mvardtAvaliacao = value
        End Set
    End Property
#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tAvaliacClie As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable                     *
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
        tAvaliacClie = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tAvaliacClie.Open("repEmprAvaliacClie", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tAvaliacClie.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tAvaliacClie.Close()

                If tipo = 1 Then
                    tAvaliacClie.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tAvaliacClie.Open("repEmprAvaliacClie", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprAvaliacClie - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tAvaliacClie.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tAvaliacClie.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tAvaliacClie.EOF Then
            idEmpresa = tAvaliacClie.Fields("idEmpresa").Value
            nomeEmpresa = tAvaliacClie.Fields("nomeEmpresa").Value
            idCliente = tAvaliacClie.Fields("idCliente").Value
            nomeCliente = tAvaliacClie.Fields("nomeCliente").Value
            nomeAvaliador = tAvaliacClie.Fields("nomeAvaliador").Value
            cargoAvaliador = tAvaliacClie.Fields("cargoAvaliador").Value
            dtAvaliacao = tAvaliacClie.Fields("dtAvaliacao").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprAvaliacClie - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String

        carregaMemos = tAvaliacClie.Fields("txtAvaliacao").Value

    End Function
End Class