﻿Option Explicit On
Public Class EmprOcorrenc

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850312"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320312"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0312"

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


    Private mvaridTpOcorrenc As Short
    Public Property idTpOcorrenc() As Short
        Get
            Return mvaridTpOcorrenc
        End Get
        Set(value As Short)
            mvaridTpOcorrenc = value
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

    Public tEmprOcorrenc As ADODB.Recordset
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
        tEmprOcorrenc = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            tEmprOcorrenc.Open("EmprOcorrenc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprOcorrenc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprOcorrenc.Close()

                If tipo = 1 Then
                    tEmprOcorrenc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprOcorrenc.Open("EmprOcorrenc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprOcorrenc - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprOcorrenc.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprOcorrenc.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprOcorrenc.EOF Then
            idEmpresa = tEmprOcorrenc.Fields("idEmpresa").Value
            idTpOcorrenc = tEmprOcorrenc.Fields("idTpOcorrenc").Value
            dtOcorrenc = tEmprOcorrenc.Fields("dtOcorrenc").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprOcorrenc - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, tpOcor As Short, data As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprOcorrenc.                        *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * tpOcor    = identif. para pesquisa   *
        '* * data    = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprOcorrenc WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idTpOcorrenc = " & tpOcor
        vSelect = vSelect & " AND Convert(datetime, dtOcorrenc, 112) = '" & Format(data, "yyyymmdd") & "'"
        tEmprOcorrenc.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprOcorrenc.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprOcorrenc.Fields("idEmpresa").Value
            idTpOcorrenc = tEmprOcorrenc.Fields("idTpOcorrenc").Value
            dtOcorrenc = tEmprOcorrenc.Fields("dtOcorrenc").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprOcorrenc.Close()
                tEmprOcorrenc.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprOcorrenc - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(ocorrenc As String) As Integer

        'Define Variáveis:
        Dim vAux As String

        On Error GoTo Einclui
        tEmprOcorrenc.AddNew()
        tEmprOcorrenc(0).Value = idEmpresa
        tEmprOcorrenc(1).Value = idTpOcorrenc
        tEmprOcorrenc(2).Value = dtOcorrenc
        'Texto da Ocorrência:
        tEmprOcorrenc(3).Value = IIf(ocorrenc = "", " ", ocorrenc)
        tEmprOcorrenc.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprOcorrenc - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(ocorrenc As String) As Integer

        'Define Variáveis:
        Dim vAux As String

        On Error GoTo Ealtera
        'Texto da Ocorrência:
        tEmprOcorrenc(3).Value = IIf(ocorrenc = "", " ", ocorrenc)
        tEmprOcorrenc.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprOcorrenc - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String

        carregaMemos = tEmprOcorrenc.Fields("txtOcorrenc").Value

    End Function

    Public Function elimina(empresa As Short, tipo As Short, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprOcorrenc.                        *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprOcorrenc WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idTpOcorrenc = " & tipo
        vSelec = vSelec & " AND Convert(datetime, dtOcorrenc, 112) = '" & Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprOcorrenc - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
