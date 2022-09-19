Option Explicit On
Public Class ProjFatosRelev
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850714"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320714"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0714"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
        End Set
    End Property

    Private mvardtGeracaoFato As Date
    Public Property dtGeracaoFato() As Date
        Get
            Return mvardtGeracaoFato
        End Get
        Set(value As Date)
            mvardtGeracaoFato = value
        End Set
    End Property

    Private mvaridTpFato As Short
    Public Property idTpFato() As Short
        Get
            Return mvaridTpFato
        End Get
        Set(value As Short)
            mvaridTpFato = value
        End Set
    End Property

    Private mvaridColabFato As Single
    Public Property idColabFato() As Single
        Get
            Return mvaridColabFato
        End Get
        Set(value As Single)
            mvaridColabFato = value
        End Set
    End Property
    Private mvardtSolucaoFato As Date
    Public Property dtSolucaoFato() As Date
        Get
            Return mvardtSolucaoFato
        End Get
        Set(value As Date)
            mvardtSolucaoFato = value
        End Set
    End Property
#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tProjFatosRelev As ADODB.Recordset
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
        tProjFatosRelev = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjFatosRelev.Open("ProjFatosRelev", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjFatosRelev.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjFatosRelev.Close()

                If tipo = 1 Then
                    tProjFatosRelev.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjFatosRelev.Open("ProjFatosRelev", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjFatosRelev - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjFatosRelev.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjFatosRelev.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjFatosRelev.EOF Then
            idProjeto = tProjFatosRelev.Fields("idProjeto").Value
            dtGeracaoFato = tProjFatosRelev.Fields("dtGeracaoFato").Value
            idTpFato = tProjFatosRelev.Fields("idTpFato").Value
            idColabFato = tProjFatosRelev.Fields("idColabFato").Value
            dtSolucaoFato = tProjFatosRelev.Fields("dtSolucaoFato").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjFatosRelev - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(projeto As Double, data As Date, tipo As Short, colab As Single, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjFluxo.                           *
        '* *                                      *
        '* * projeto   = identif. para pesquisa + *
        '* * data      = identif. para pesquisa + *
        '* * tipo      = identif. para pesquisa + *
        '* * colab     = identif. para pesquisa   *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjFatosRelev WHERE idProjeto = " & projeto
        vSelect = vSelect & " AND Convert(datetime, dtGeracaoFato, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        vSelect = vSelect & " AND idTpFato = " & tipo
        vSelect = vSelect & " AND idColabFato = " & colab
        tProjFatosRelev.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjFatosRelev.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjFatosRelev.Fields("idProjeto").Value
            dtGeracaoFato = tProjFatosRelev.Fields("dtGeracaoFato").Value
            idTpFato = tProjFatosRelev.Fields("idTpFato").Value
            idColabFato = tProjFatosRelev.Fields("idColabFato").Value
            dtSolucaoFato = tProjFatosRelev.Fields("dtSolucaoFato").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjFatosRelev.Close()
                tProjFatosRelev.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjFatosRelev - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function
    Public Function inclui(ocorrenc As String, solucao As String) As Integer

        On Error GoTo Einclui
        tProjFatosRelev.AddNew()
        tProjFatosRelev(0).Value = idProjeto
        tProjFatosRelev(1).Value = dtGeracaoFato
        tProjFatosRelev(2).Value = idTpFato
        tProjFatosRelev(3).Value = idColabFato
        tProjFatosRelev(4).Value = IIf(ocorrenc = "", " ", ocorrenc)
        tProjFatosRelev(5).Value = dtSolucaoFato
        tProjFatosRelev(6).Value = IIf(solucao = "", " ", solucao)
        tProjFatosRelev.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjFatosRelev - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function altera(ocorrenc As String, solucao As String) As Integer

        On Error GoTo Ealtera
        tProjFatosRelev(4).Value = IIf(ocorrenc = "", " ", ocorrenc)
        tProjFatosRelev(5).Value = dtSolucaoFato
        tProjFatosRelev(6).Value = IIf(solucao = "", " ", solucao)
        tProjFatosRelev.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjFatosRelev - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaOcorrenc() As String

        CarregaOcorrenc = tProjFatosRelev.Fields("txtGeracaoFato").Value

    End Function

    Public Function CarregaSolucao() As String

        CarregaSolucao = tProjFatosRelev.Fields("txtSolucaoFato").Value

    End Function

    Public Function elimina(projeto As Double, data As Date, tipo As Short, colab As Single) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjFatosRelev.                      *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* * colab      = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ProjFatosRelev WHERE "
        vSelec = vSelec & " idProjeto = " & projeto
        vSelec = vSelec & " AND Convert(datetime, dtGeracaoFato, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        vSelec = vSelec & " AND idTpFato = " & tipo
        vSelec = vSelec & " AND idColabFato = " & colab
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjFatosRelev - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function



End Class
