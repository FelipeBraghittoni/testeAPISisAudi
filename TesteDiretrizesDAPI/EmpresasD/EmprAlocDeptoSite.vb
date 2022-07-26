﻿Option Explicit On
Public Class EmprAlocDeptoSite
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850302"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320302"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0302"
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


    Private mvaridSite As Short
    Public Property idSite() As Short
        Get
            Return mvaridSite
        End Get
        Set(value As Short)
            mvaridSite = value
        End Set
    End Property


    Private mvaridDepto As Short
    Public Property idDepto() As Short
        Get
            Return mvaridDepto
        End Get
        Set(value As Short)
            mvaridDepto = value
        End Set
    End Property


    Private mvardtInicio As Date
    Public Property dtInicio() As Date
        Get
            Return mvardtInicio
        End Get
        Set(value As Date)
            mvardtInicio = value
        End Set
    End Property


    Private mvardtFim As Date
    Public Property dtFim() As Date
        Get
            Return mvardtFim
        End Get
        Set(value As Date)
            mvardtFim = value
        End Set
    End Property


    Private mvarfone1Depto As Integer
    Public Property fone1Depto() As Integer
        Get
            Return mvarfone1Depto
        End Get
        Set(value As Integer)
            mvarfone1Depto = value
        End Set
    End Property


    Private mvarramal1Depto As Short
    Public Property ramal1Depto() As Short
        Get
            Return mvarramal1Depto
        End Get
        Set(value As Short)
            mvarramal1Depto = value
        End Set
    End Property


    Private mvarfone2Depto As Integer
    Public Property fone2Depto() As Integer
        Get
            Return mvarfone2Depto
        End Get
        Set(value As Integer)
            mvarfone2Depto = value
        End Set
    End Property


    Private mvarramal2Depto As Short
    Public Property ramal2Depto() As Short
        Get
            Return mvarramal2Depto
        End Get
        Set(value As Short)
            mvarramal2Depto = value
        End Set
    End Property


    Private mvarfaxDepto As Integer
    Public Property faxDepto() As Integer
        Get
            Return mvarfaxDepto
        End Get
        Set(value As Integer)
            mvarfaxDepto = value
        End Set
    End Property


#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public tEmprAlocDeptoSite As ADODB.Recordset

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
        tEmprAlocDeptoSite = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            tEmprAlocDeptoSite.Open("EmprAlocDeptoSite", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprAlocDeptoSite.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocDeptoSite.Close()

                If tipo = 1 Then
                    tEmprAlocDeptoSite.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprAlocDeptoSite.Open("EmprAlocDeptoSite", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprAlocDeptoSite - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprAlocDeptoSite.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprAlocDeptoSite.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprAlocDeptoSite.EOF Then
            idEmpresa = tEmprAlocDeptoSite.Fields("idEmpresa").Value
            idSite = tEmprAlocDeptoSite.Fields("idSite").Value
            idDepto = tEmprAlocDeptoSite.Fields("idDepto").Value
            dtInicio = tEmprAlocDeptoSite.Fields("dtInicio").Value
            dtFim = tEmprAlocDeptoSite.Fields("dtFim").Value
            fone1Depto = tEmprAlocDeptoSite.Fields("fone1Depto").Value
            ramal1Depto = tEmprAlocDeptoSite.Fields("ramal1Depto").Value
            fone2Depto = tEmprAlocDeptoSite.Fields("fone2Depto").Value
            ramal2Depto = tEmprAlocDeptoSite.Fields("ramal2Depto").Value
            faxDepto = tEmprAlocDeptoSite.Fields("faxDepto").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprAlocDeptoSite - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, site As Short, depto As Short, data As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprAlocDeptoSite.                   *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * site       = identif. para pesquisa  *
        '* * depto      = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprAlocDeptoSite WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idSite = " & site
        vSelect = vSelect & " AND idDepto = " & depto
        vSelect = vSelect & " AND Convert(datetime, dtInicio, 112) = '" & Format(data, "yyyymmdd") & "'"
        tEmprAlocDeptoSite.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprAlocDeptoSite.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprAlocDeptoSite.Fields("idEmpresa").Value
            idSite = tEmprAlocDeptoSite.Fields("idSite").Value
            idDepto = tEmprAlocDeptoSite.Fields("idDepto").Value
            dtInicio = tEmprAlocDeptoSite.Fields("dtInicio").Value
            dtFim = tEmprAlocDeptoSite.Fields("dtFim").Value
            fone1Depto = tEmprAlocDeptoSite.Fields("fone1Depto").Value
            ramal1Depto = tEmprAlocDeptoSite.Fields("ramal1Depto").Value
            fone2Depto = tEmprAlocDeptoSite.Fields("fone2Depto").Value
            ramal2Depto = tEmprAlocDeptoSite.Fields("ramal2Depto").Value
            faxDepto = tEmprAlocDeptoSite.Fields("faxDepto").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocDeptoSite.Close()
                tEmprAlocDeptoSite.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprAlocDeptoSite - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tEmprAlocDeptoSite.AddNew()
        tEmprAlocDeptoSite(0).Value = idEmpresa
        tEmprAlocDeptoSite(1).Value = idSite
        tEmprAlocDeptoSite(2).Value = idDepto
        tEmprAlocDeptoSite(3).Value = dtInicio
        tEmprAlocDeptoSite(4).Value = dtFim
        tEmprAlocDeptoSite(5).Value = fone1Depto
        tEmprAlocDeptoSite(6).Value = ramal1Depto
        tEmprAlocDeptoSite(7).Value = fone2Depto
        tEmprAlocDeptoSite(8).Value = ramal2Depto
        tEmprAlocDeptoSite(9).Value = faxDepto
        tEmprAlocDeptoSite.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprAlocDeptoSite - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tEmprAlocDeptoSite(4).Value = dtFim
        tEmprAlocDeptoSite(5).Value = fone1Depto
        tEmprAlocDeptoSite(6).Value = ramal1Depto
        tEmprAlocDeptoSite(7).Value = fone2Depto
        tEmprAlocDeptoSite(8).Value = ramal2Depto
        tEmprAlocDeptoSite(9).Value = faxDepto
        tEmprAlocDeptoSite.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprAlocDeptoSite - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(empresa As Short, site As Short, depto As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprAlocDeptoSite.                   *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * site       = identif. para pesquisa  *
        '* * depto      = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprAlocDeptoSite WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idSite = " & site
        vSelec = vSelec & " AND idDepto = " & depto
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprAlocDeptoSite - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function





End Class
