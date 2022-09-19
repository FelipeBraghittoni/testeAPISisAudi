Option Explicit On
Public Class ProjTpNaturFatur
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850727"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320727"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0727"
#End Region

#Region "VAriaveis de ambiente"
    Private mvaridTpNaturFatur As Short
    Public Property idTpNaturFatur() As Short
        Get
            Return mvaridTpNaturFatur
        End Get
        Set(value As Short)
            mvaridTpNaturFatur = value
        End Set
    End Property

    Private mvarnomeTpNaturFatur As String
    Public Property nomeTpNaturFatur() As String
        Get
            Return mvarnomeTpNaturFatur
        End Get
        Set(value As String)
            mvarnomeTpNaturFatur = value
        End Set
    End Property

#End Region

#Region "conexão com banco"
    Public Db As ADODB.Connection
    Public tProjTpNaturFatur As ADODB.Recordset
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
        '* * 2 = OpenTable, índice IndiceNome     *
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
        tProjTpNaturFatur = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjTpNaturFatur.Open("ProjTpNaturFatur", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjTpNaturFatur ORDER BY nomeTpNaturFatur"
                tProjTpNaturFatur.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjTpNaturFatur.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjTpNaturFatur.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjTpNaturFatur.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjTpNaturFatur.Open("ProjTpNaturFatur", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjTpNaturFatur - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjTpNaturFatur.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjTpNaturFatur.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjTpNaturFatur.EOF Then
            idTpNaturFatur = tProjTpNaturFatur.Fields("idTpNaturFatur").Value
            nomeTpNaturFatur = tProjTpNaturFatur.Fields("nomeTpNaturFatur").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjTpNaturFatur - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(tipoFatos As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjTpNaturFatur.                    *
        '* *                                      *
        '* * tipoFatos  = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjTpNaturFatur WHERE idTpNaturFatur = " & tipoFatos
        tProjTpNaturFatur.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjTpNaturFatur.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idTpNaturFatur = tProjTpNaturFatur.Fields("idTpNaturFatur").Value
            nomeTpNaturFatur = tProjTpNaturFatur.Fields("nomeTpNaturFatur").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjTpNaturFatur.Close()
                tProjTpNaturFatur.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjTpNaturFatur - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjTpNaturFatur                     *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjTpNaturFatur WHERE nomeTpNaturFatur = '" & descric & "'"
        tProjTpNaturFatur.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjTpNaturFatur.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idTpNaturFatur = tProjTpNaturFatur.Fields("idTpNaturFatur").Value
            nomeTpNaturFatur = tProjTpNaturFatur.Fields("nomeTpNaturFatur").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjTpNaturFatur.Close()
                tProjTpNaturFatur.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjTpNaturFatur - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tProjTpNaturFatur.AddNew()
        tProjTpNaturFatur(0).Value = cCodigo
        tProjTpNaturFatur(1).Value = cDescricao
        tProjTpNaturFatur.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjTpNaturFatur - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tProjTpNaturFatur(1).Value = cDescricao
        tProjTpNaturFatur.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjTpNaturFatur - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjTpNaturFatur.                    *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ProjTpNaturFatur WHERE "
        vSelec = vSelec & " idTpNaturFatur = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjTpNaturFatur - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
