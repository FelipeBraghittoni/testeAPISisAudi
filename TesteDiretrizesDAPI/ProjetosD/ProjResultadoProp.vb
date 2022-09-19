Option Explicit On
Public Class ProjResultadoProp
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850721"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320721"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0721"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridResultado As Short
    Public Property idResultado() As Short
        Get
            Return mvaridResultado
        End Get
        Set(value As Short)
            mvaridResultado = value
        End Set
    End Property

    Private mvardescricResultado As String
    Public Property descricResultado() As String
        Get
            Return mvardescricResultado
        End Get
        Set(value As String)
            mvardescricResultado = value
        End Set
    End Property
#End Region

#Region "conexão de banco"
    Public Db As ADODB.Connection
    Public tProjResultadoProp As ADODB.Recordset
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
        tProjResultadoProp = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjResultadoProp.Open("ProjResultadoProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjResultadoProp ORDER BY descricResultado"
                tProjResultadoProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjResultadoProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjResultadoProp.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjResultadoProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjResultadoProp.Open("ProjResultadoProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjResultadoProp - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tProjResultadoProp.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjResultadoProp.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjResultadoProp.EOF Then
            idResultado = tProjResultadoProp.Fields("idResultado").Value
            descricResultado = tProjResultadoProp.Fields("descricResultado").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjResultadoProp - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(identif As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjResultadoProp                    *
        '* *                                      *
        '* * identif = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjResultadoProp WHERE idResultado = " & identif
        tProjResultadoProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjResultadoProp.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idResultado = tProjResultadoProp.Fields("idResultado").Value
            descricResultado = tProjResultadoProp.Fields("descricResultado").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjResultadoProp.Close()
                tProjResultadoProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjResultadoProp - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjResultadoProp                    *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjResultadoProp WHERE descricResultado = '" & descric & "'"
        tProjResultadoProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjResultadoProp.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idResultado = tProjResultadoProp.Fields("idResultado").Value
            descricResultado = tProjResultadoProp.Fields("descricResultado").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjResultadoProp.Close()
                tProjResultadoProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjResultadoProp - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tProjResultadoProp.AddNew()
        tProjResultadoProp(0).Value = cCodigo
        tProjResultadoProp(1).Value = cDescricao
        tProjResultadoProp.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjResultadoProp - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tProjResultadoProp(1).Value = cDescricao
        tProjResultadoProp.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjResultadoProp - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjResultadoProp                    *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ProjResultadoProp WHERE "
        vSelec = vSelec & " idResultado = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjResultadoProp - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
