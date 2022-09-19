Option Explicit On
Public Class ProjFluxo
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850715"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320715"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0715"
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

    Private mvaridStatusProjeto As Short
    Public Property idStatusProjeto() As Short
        Get
            Return mvaridStatusProjeto
        End Get
        Set(value As Short)
            mvaridStatusProjeto = value
        End Set
    End Property

    Private mvaridColabStatus As Single
    Public Property idColabStatus() As Single
        Get
            Return mvaridColabStatus
        End Get
        Set(value As Single)
            mvaridColabStatus = value
        End Set
    End Property

    Private mvardtStatus As Date
    Public Property dtStatus() As Date
        Get
            Return mvardtStatus
        End Get
        Set(value As Date)
            mvardtStatus = value
        End Set
    End Property

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tProjFluxo As ADODB.Recordset
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
        tProjFluxo = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjFluxo.Open("ProjFluxo", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tProjFluxo.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjFluxo.Close()

                If tipo = 1 Then
                    tProjFluxo.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjFluxo.Open("ProjFluxo", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjFluxo - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjFluxo.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjFluxo.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjFluxo.EOF Then
            idProjeto = tProjFluxo.Fields("idProjeto").Value
            idStatusProjeto = tProjFluxo.Fields("idStatusProjeto").Value
            idColabStatus = tProjFluxo.Fields("idColabStatus").Value
            dtStatus = tProjFluxo.Fields("dtStatus").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjFluxo - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(projeto As Double, status As Short, data As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjFluxo.                           *
        '* *                                      *
        '* * projeto = identif. para pesquisa +   *
        '* * status  = identif. para pesquisa     *
        '* * data    = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjFluxo WHERE idProjeto = " & projeto
        vSelect = vSelect & " AND idStatusProjeto = " & status
        vSelect = vSelect & " AND Convert(datetime, dtStatus, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        tProjFluxo.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjFluxo.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProjeto = tProjFluxo.Fields("idProjeto").Value
            idStatusProjeto = tProjFluxo.Fields("idStatusProjeto").Value
            idColabStatus = tProjFluxo.Fields("idColabStatus").Value
            dtStatus = tProjFluxo.Fields("dtStatus").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjFluxo.Close()
                tProjFluxo.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjFluxo - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(status As String) As Integer

        On Error GoTo Einclui
        tProjFluxo.AddNew()
        tProjFluxo(0).Value = idProjeto
        tProjFluxo(1).Value = idStatusProjeto
        tProjFluxo(2).Value = idColabStatus
        tProjFluxo(3).Value = dtStatus
        tProjFluxo(4).Value = IIf(status = "", " ", status)
        tProjFluxo.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjFluxo - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(status As String) As Integer

        On Error GoTo Ealtera
        tProjFluxo(2).Value = idColabStatus
        tProjFluxo(4).Value = IIf(status = "", " ", status)
        tProjFluxo.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjFluxo - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaMemos() As String

        CarregaMemos = tProjFluxo.Fields("txtStatus").Value

    End Function

    Public Function elimina(projeto As Double, tipo As Short, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjFluxo.                           *
        '* *                                      *
        '* * projeto    = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ProjFluxo WHERE "
        vSelec = vSelec & " idProjeto = " & projeto
        vSelec = vSelec & " AND idStatusProjeto = " & tipo
        vSelec = vSelec & " AND Convert(datetime, dtStatus, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjFluxo - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
