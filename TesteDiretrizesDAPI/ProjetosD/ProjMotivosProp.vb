Option Explicit On
Public Class ProjMotivosProp
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f55378507038"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320738"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0738"
#End Region

#Region "variaveis de ambiente"
    'local variable(s) to hold property value(s)
    Private mvaridMotivoProp As Short
    Public Property idMotivoProp() As Short
        Get
            Return mvaridMotivoProp
        End Get
        Set(value As Short)
            mvaridMotivoProp = value
        End Set
    End Property

    Private mvardescricMotivoProp As String
    Public Property descricMotivoProp() As String
        Get
            Return mvardescricMotivoProp
        End Get
        Set(value As String)
            mvardescricMotivoProp = value
        End Set
    End Property

#End Region

#Region "conexão com banco"
    'Vari�veis para abrir as tabelas relacionadas �s classes:
    Public Db As ADODB.Connection
    Public tProjMotivosProp As ADODB.Recordset
#End Region


    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou n�o o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = n�o abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable, �ndice ChavePrimaria  *
        '* * 1 = OpenDynaset                      *
        '* * 2 = OpenTable, �ndice IndiceNome     *
        '* ****************************************

        On Error GoTo EdbConecta

        dbConecta = 0   'ReturnCode se n�o houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            '* ********************************************
            '* Cria uma inst�ncia de ADODB.Connection:    *
            '* ********************************************
            Dim strConnect As String
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            Db.Open(strConnect)
        End If

        'Cria uma inst�ncia de ADODB.Recordset
        tProjMotivosProp = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjMotivosProp.Open("ProjMotivosProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM ProjMotivosProp ORDER BY descricMotivoProp"
                tProjMotivosProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjMotivosProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjMotivosProp.Close()

                If tipo = 1 Or tipo = 2 Then
                    tProjMotivosProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjMotivosProp.Open("ProjMotivosProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe ProjMotivosProp - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Short) As Integer
        '* ****************************************
        '* * L� sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se � a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se n�o houver nenhum problema
        'Se n�o chegou no final do arquivo:
        If Not tProjMotivosProp.EOF Then
            If vPrimVez = 0 Then    'N�o � a primeira vez
                tProjMotivosProp.MoveNext()    'l� 1 linha
            End If
        End If
        'Se n�o chegou no final do arquivo, carrega propriedades:
        If Not tProjMotivosProp.EOF Then
            idMotivoProp = tProjMotivosProp.Fields("idMotivoProp").Value
            descricMotivoProp = tProjMotivosProp("descricMotivoProp").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjMotivosProp - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(identif As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro espec�fico,     *
        '* * baseado na Chave Prim�ria da tabela  *
        '* * ProjMotivosProp                      *
        '* *                                      *
        '* * identif = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjMotivosProp WHERE idMotivoProp = " & identif
        tProjMotivosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se n�o houver nenhum problema

        'N�o encontrou:
        If tProjMotivosProp.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idMotivoProp = tProjMotivosProp.Fields("idMotivoProp").Value
            descricMotivoProp = tProjMotivosProp.Fields("descricMotivoProp").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjMotivosProp.Close()
                tProjMotivosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjMotivosProp - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro espec�fico,     *
        '* * baseado na Chave Prim�ria da tabela  *
        '* * ProjMotivosProp                      *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM ProjMotivosProp WHERE descricMotivoProp = '" & descric & "'"
        tProjMotivosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se n�o houver nenhum problema

        'N�o encontrou:
        If tProjMotivosProp.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idMotivoProp = tProjMotivosProp.Fields("idMotivoProp").Value
            descricMotivoProp = tProjMotivosProp.Fields("descricMotivoProp").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjMotivosProp.Close()
                tProjMotivosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe ProjMotivosProp - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tProjMotivosProp.AddNew()
        tProjMotivosProp(0).Value = cCodigo
        tProjMotivosProp(1).Value = cDescricao
        tProjMotivosProp.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjMotivosProp - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tProjMotivosProp(1).Value = cDescricao
        tProjMotivosProp.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjMotivosProp - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjMotivosProp                      *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se n�o houver erro

        vSelec = "DELETE FROM ProjMotivosProp WHERE "
        vSelec = vSelec & " idMotivoProp = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjMotivosProp - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
