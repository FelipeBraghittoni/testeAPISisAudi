Option Explicit On

<ComClass(DirTpDiretriz.ClassId, DirTpDiretriz.InterfaceId, DirTpDiretriz.EventsId)>
Public Class DirTpDiretriz

    Private mvartpDiretriz As Integer 'local copy
    Private mvardescricTpDiretriz As String 'local copy

    Public Db As ADODB.Connection
    Public tDirTpDiretriz As ADODB.Recordset
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850101"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320101"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0101"
#End Region

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        MyBase.New()
    End Sub
    '###########################

    Public Property descricTpDiretriz() As String
        Get
            Return mvardescricTpDiretriz
        End Get
        Set(ByVal value As String)
            mvardescricTpDiretriz = value
        End Set
    End Property

    Public Property tpDiretriz() As Integer
        Get
            Return mvartpDiretriz
        End Get
        Set(ByVal value As Integer)
            mvartpDiretriz = value
        End Set
    End Property

    Public Function dbConecta(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As Integer
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
            'strConnect = cdSeguranca1.LeDADOSsys(1)

            strConnect = "Provider = SQLOLEDB;Data Source=DESKTOP-HV9333S;Initial Catalog=Auditeste;User ID=felipe.rozzi2;Password=Felipe1999#"

            'MsgBox(System.AppDomain.CurrentDomain.BaseDirectory())
            'MsgBox(strConnect)
            'strConnect = "driver={SQL Server};server=localhost\SQLEXPRESS;database=auditeste;uid=sa;pwd=auditeste"
            'strConnect = "driver={SQL Server};server=localhost\SQLEXPRESS;database=auditeste;uid=sa;pwd=audi952637" 'locahost Zé Antonio
            'strConnect = "Provider=SQLOLEDB;Data Source=AudiTeste051,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=audi952637;" 'Zé Antonio mesmo acima para VM
            'strConnect = "Provider=SQLOLEDB;Data Source=192.168.1.106,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=audi952637;" 'Zé Antonio mesmo acima para VM
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tDirTpDiretriz = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tDirTpDiretriz.Open("dbo.DirTpDiretriz", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic,)
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "Select * FROM DirTpDiretriz ORDER BY descricTpDiretriz"
                tDirTpDiretriz.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tDirTpDiretriz.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number <> 0 Then '#AJUSTE
            If Err.Number = 3705 Then '#AJUSTE Err().Number().Equals() -> Err.Number = 
                tDirTpDiretriz.Close()
                If tipo = 1 Or tipo = 2 Then
                    tDirTpDiretriz.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tDirTpDiretriz.Open("DirTpDiretriz", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                End If
                Resume Next
            Else
                dbConecta = Err.Number '#AJUSTE Err().Number -> Err.Number
                If Err.Number = 3024 Then dbConecta = 1046 '#AJUSTE Err().Number().Equals() -> Err.Number = 
                '#                MsgBox "Classe DirTpDiretriz - dbConecta" & Chr$(13) & Chr$(10) & Err() & " - " & Error()
                MsgBox("Classe DirTpDiretriz - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Integer) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tDirTpDiretriz.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tDirTpDiretriz.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tDirTpDiretriz.EOF Then
            'Dim aux As String = tDirTpDiretriz.Fields("tpDiretriz").ToString
            tpDiretriz = tDirTpDiretriz.Fields("tpDiretriz").Value '#AJUSTE ToString -> Value
            descricTpDiretriz = tDirTpDiretriz.Fields("descricTpDiretriz").Value '#AJUSTE ToString -> Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err().Equals(True) Then
            leSeq = Err().Number
            MsgBox("Classe DirTpDiretriz - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(tipo As Integer, atualiz As Integer) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * DirTpDiretriz.                       *
        '* *                                      *
        '* * tipo       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "Select * FROM DirTpDiretriz WHERE tpDiretriz = " & tipo
        tDirTpDiretriz.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
        localiza = 0    'ReturnCode se não houver nenhum problema
        'Não encontrou:
        If tDirTpDiretriz.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            tpDiretriz = tDirTpDiretriz.Fields("tpDiretriz").Value '#AJUSTE ToString -> Value
            descricTpDiretriz = tDirTpDiretriz.Fields("descricTpDiretriz").Value '#AJUSTE ToString -> Value
        End If

Elocaliza:
        If Err.Number <> 0 Then
            If Err.Number = 3705 Then
                tDirTpDiretriz.Close()
                tDirTpDiretriz.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err().Number
                MsgBox("Classe DirTpDiretriz - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If
    End Function


    Public Function localizaNome(descric As String, atualiz As Integer) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * DirTpDiretriz.                       *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "Select * FROM DirTpDiretriz WHERE descricTpDiretriz = '" & descric & "'"
        tDirTpDiretriz.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tDirTpDiretriz.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            tpDiretriz = tDirTpDiretriz.Fields("tpDiretriz").Value '#AJUSTE ToString -> Value
            descricTpDiretriz = tDirTpDiretriz.Fields("descricTpDiretriz").Value '#AJUSTE ToString -> Value

        End If

ElocalizaNome:
        If Err.Number <> 0 Then
            If Err.Number = 3705 Then
                tDirTpDiretriz.Close()
                tDirTpDiretriz.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err().Number
                MsgBox("Classe DirTpDiretriz - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function
    Public Function inclui(cCodigo As Integer, cDescricao As String) As Integer
        On Error GoTo Einclui
        'tDirTpDiretriz.AddNew(0 = cCodigo, 1 = cDescricao) 
        tDirTpDiretriz.AddNew()
        tDirTpDiretriz(0).Value = cCodigo '#Ajuste
        tDirTpDiretriz(1).Value = cDescricao '#Ajuste
        tDirTpDiretriz.Update()
        inclui = 0

Einclui:
        If Err.Number <> 0 Then
            inclui = Err.Number
            MsgBox("Classe DirTpDiretriz - inclui" & vbCrLf & Err.Number & " - " & Err.Description) '#AJUSTE
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        'tDirTpDiretriz.Update(1 = cDescricao)
        tDirTpDiretriz(1).Value = cDescricao '#Ajuste
        tDirTpDiretriz.Update()
        altera = 0

Ealtera:
        If Err.Number <> 0 Then
            altera = Err.Number
            MsgBox("Classe DirTpDiretriz - altera" & vbCrLf & Err.Number & " - " & Err.Description) '#AJUSTE
        End If

    End Function

    Public Function elimina(codigo As Integer) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * DirTpDiretriz.                       *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM DirTpDiretriz WHERE "
        vSelec = vSelec & " tpDiretriz = " & codigo
        Db.Execute(vSelec, ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number <> 0 Then
            elimina = Err.Number
            MsgBox("Classe DirTpDiretriz - elimina" & vbCrLf & Err.Number & " - " & Err.Description) '#AJUSTE
        End If

    End Function
End Class


