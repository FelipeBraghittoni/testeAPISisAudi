Option Explicit On
Public Class EmprAlocGerDepto
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850303"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320303"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0303"
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


    Private mvaridDepto As Short
    Public Property idDepto() As Short
        Get
            Return mvaridDepto
        End Get
        Set(value As Short)
            mvaridDepto = value
        End Set
    End Property


    Private mvaridColabResp As Single
    Public Property idColabResp() As Single
        Get
            Return mvaridColabResp
        End Get
        Set(value As Single)
            mvaridColabResp = value
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

#End Region

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public tEmprAlocGerDepto As ADODB.Recordset
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
        tEmprAlocGerDepto = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            tEmprAlocGerDepto.Open("EmprAlocGerDepto", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tEmprAlocGerDepto.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocGerDepto.Close()

                If tipo = 1 Then
                    tEmprAlocGerDepto.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprAlocGerDepto.Open("EmprAlocGerDepto", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprAlocGerDepto - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprAlocGerDepto.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprAlocGerDepto.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprAlocGerDepto.EOF Then
            idEmpresa = tEmprAlocGerDepto.Fields("idEmpresa").Value
            idDepto = tEmprAlocGerDepto.Fields("idDepto").Value
            idColabResp = tEmprAlocGerDepto.Fields("idColabResp").Value
            dtInicio = tEmprAlocGerDepto.Fields("dtInicio").Value
            dtFim = tEmprAlocGerDepto.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprAlocGerDepto - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, depto As Short, colaborador As Short, data As Date, atualiz As Short) As Integer
        '* *****************************************
        '* * Localiza um registro específico,      *
        '* * baseado na Chave Primária da tabela   *
        '* * EmprAlocGerDepto.                     *
        '* *                                       *
        '* * empresa     = identif. para pesquisa  *
        '* * depto       = identif. para pesquisa  *
        '* * colaborador = identif. para pesquisa  *
        '* * data        = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe  *
        '* *****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprAlocGerDepto WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idDepto = " & depto
        vSelect = vSelect & " AND idColabResp = " & colaborador
        vSelect = vSelect & " AND Convert(datetime, dtInicio, 112) = '" & Format(data, "yyyymmdd") & "'"
        tEmprAlocGerDepto.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprAlocGerDepto.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprAlocGerDepto.Fields("idEmpresa").Value
            idDepto = tEmprAlocGerDepto.Fields("idDepto").Value
            idColabResp = tEmprAlocGerDepto.Fields("idColabResp").Value
            dtInicio = tEmprAlocGerDepto.Fields("dtInicio").Value
            dtFim = tEmprAlocGerDepto.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprAlocGerDepto.Close()
                tEmprAlocGerDepto.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprAlocGerDepto - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui() As Integer

        On Error GoTo Einclui
        tEmprAlocGerDepto.AddNew()
        tEmprAlocGerDepto.Fields("idEmpresa").Value = idEmpresa
        tEmprAlocGerDepto.Fields("idDepto").Value = idDepto
        tEmprAlocGerDepto.Fields("idColabResp").Value = idColabResp
        tEmprAlocGerDepto.Fields("dtInicio").Value = dtInicio
        tEmprAlocGerDepto.Fields("dtFim").Value = dtFim
        tEmprAlocGerDepto.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprAlocGerDepto - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera() As Integer

        On Error GoTo Ealtera
        tEmprAlocGerDepto.Fields("dtFim").Value = dtFim
        tEmprAlocGerDepto.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprAlocGerDepto - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(empresa As Short, depto As Short, colaborador As Short) As Integer
        '* *****************************************
        '* * Elimina um registro da tabela         *
        '* * EmprAlocGerDepto.                     *
        '* *                                       *
        '* * empresa     = identif. para pesquisa  *
        '* * colaborador = identif. para pesquisa  *
        '* * depto       = identif. para pesquisa  *
        '* *****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprAlocGerDepto WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        vSelec = vSelec & " AND idColabResp = " & colaborador
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprAlocGerDepto - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
