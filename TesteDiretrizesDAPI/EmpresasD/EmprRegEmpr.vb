Option Explicit On
Public Class EmprRegEmpr
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850313"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320313"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0313"
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


    Private mvaridCadDocto As Short
    Public Property idCadDocto() As Short
        Get
            Return mvaridCadDocto
        End Get
        Set(value As Short)
            mvaridCadDocto = value
        End Set
    End Property


    Private mvarnumDoctoEmpr As String
    Public Property numDoctoEmpr() As String
        Get
            Return mvarnumDoctoEmpr
        End Get
        Set(value As String)
            mvarnumDoctoEmpr = value
        End Set
    End Property


    Private mvarcomplemen1Docto As String
    Public Property complemen1Docto() As String
        Get
            Return mvarcomplemen1Docto
        End Get
        Set(value As String)
            mvarcomplemen1Docto = value
        End Set
    End Property


    Private mvarcomplemen2Docto As String
    Public Property complemen2Docto() As String
        Get
            Return mvarcomplemen2Docto
        End Get
        Set(value As String)
            mvarcomplemen2Docto = value
        End Set
    End Property

    Private mvarcomplemen3Docto As String
    Public Property complemen3Docto() As String
        Get
            Return mvarcomplemen3Docto
        End Get
        Set(value As String)
            mvarcomplemen3Docto = value
        End Set
    End Property

#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public tEmprRegEmpr As ADODB.Recordset

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
        tEmprRegEmpr = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            tEmprRegEmpr.Open("EmprRegEmpr", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprRegEmpr.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprRegEmpr.Close()

                If tipo = 1 Then
                    tEmprRegEmpr.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprRegEmpr.Open("EmprRegEmpr", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprRegEmpr - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprRegEmpr.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprRegEmpr.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprRegEmpr.EOF Then
            idEmpresa = tEmprRegEmpr.Fields("idEmpresa").Value
            idCadDocto = tEmprRegEmpr.Fields("idCadDocto").Value
            numDoctoEmpr = tEmprRegEmpr.Fields("numDoctoEmpr").Value
            complemen1Docto = tEmprRegEmpr.Fields("complemen1Docto").Value
            complemen2Docto = tEmprRegEmpr.Fields("complemen2Docto").Value
            complemen3Docto = tEmprRegEmpr.Fields("complemen3Docto").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprRegEmpr - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empr As Short, docto As Short, numDocto As String, complem1 As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprRegEmpr.                         *
        '* *                                      *
        '* * empr    = identif. para pesquisa     *
        '* * docto   = identif. para pesquisa     *
        '* * numdocto= identif. para pesquisa     *
        '* * complem1= identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprRegEmpr WHERE idEmpresa = " & empr
        vSelect = vSelect & " AND idCadDocto = " & docto
        vSelect = vSelect & " AND numDoctoEmpr = '" & numDocto & "'"
        vSelect = vSelect & " AND complemen1Docto = '" & complem1 & "'"
        tEmprRegEmpr.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprRegEmpr.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprRegEmpr.Fields("idEmpresa").Value
            idCadDocto = tEmprRegEmpr.Fields("idCadDocto").Value
            numDoctoEmpr = tEmprRegEmpr.Fields("numDoctoEmpr").Value
            complemen1Docto = tEmprRegEmpr.Fields("complemen1Docto").Value
            complemen2Docto = tEmprRegEmpr.Fields("complemen2Docto").Value
            complemen3Docto = tEmprRegEmpr.Fields("complemen3Docto").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprRegEmpr.Close()
                tEmprRegEmpr.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprRegEmpr - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        'Define Variáveis:
        Dim vAux As String

        On Error GoTo Einclui
        tEmprRegEmpr.AddNew()
        tEmprRegEmpr(0).Value = idEmpresa
        tEmprRegEmpr(1).Value = idCadDocto
        tEmprRegEmpr(2).Value = numDoctoEmpr
        tEmprRegEmpr(3).Value = complemen1Docto
        tEmprRegEmpr(4).Value = complemen2Docto
        tEmprRegEmpr(5).Value = complemen3Docto
        tEmprRegEmpr.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprRegEmpr - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        'Define Variáveis:
        Dim vAux As String

        On Error GoTo Ealtera
        tEmprRegEmpr(2).Value = numDoctoEmpr
        tEmprRegEmpr(3).Value = complemen1Docto
        tEmprRegEmpr(4).Value = complemen2Docto
        tEmprRegEmpr(5).Value = complemen3Docto
        tEmprRegEmpr.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprRegEmpr - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
    Public Function elimina(empresa As Short, tipo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprRegEmpr.                         *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprRegEmpr WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idCadDocto = " & tipo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprRegEmpr - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
