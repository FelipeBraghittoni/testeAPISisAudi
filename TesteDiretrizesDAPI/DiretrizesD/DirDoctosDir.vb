Option Explicit On
Imports System.IO

<ComClass(DirDoctosDir.ClassId, DirDoctosDir.InterfaceId, DirDoctosDir.EventsId)>
Public Class DirDoctosDir

    Public Db As ADODB.Connection
    Public tDirDoctosDir As ADODB.Recordset

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850102"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320102"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0102"
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

    Private mvaridDiretriz As Integer
    Public Property idDiretriz() As Integer
        Get
            Return mvaridDiretriz
        End Get
        Set(ByVal value As Integer)
            mvaridDiretriz = value
        End Set
    End Property

    Private mvarnomeDiretriz As String
    Public Property nomeDiretriz() As String
        Get
            Return mvarnomeDiretriz
        End Get
        Set(ByVal value As String)
            mvarnomeDiretriz = value
        End Set
    End Property

    Private mvartipodocto1 As String
    Public Property tipodocto1() As String
        Get
            Return mvartipodocto1
        End Get
        Set(ByVal value As String)
            mvartipodocto1 = value
        End Set
    End Property

    Private mvartipodocto2 As String
    Public Property tipodocto2() As String
        Get
            Return mvartipodocto2
        End Get
        Set(ByVal value As String)
            mvartipodocto2 = value
        End Set
    End Property

    Private mvardtRegistro As Date
    Public Property dtRegistro() As Date
        Get
            Return mvardtRegistro
        End Get
        Set(ByVal value As Date)
            mvardtRegistro = value
        End Set
    End Property

    Private mvartpDoctoDiretriz As String
    Public Property tpDoctoDiretriz() As String
        Get
            Return mvartpDoctoDiretriz
        End Get
        Set(ByVal value As String)
            mvartpDoctoDiretriz = value
        End Set
    End Property

    Private mvardtFim As Date
    Public Property dtFim() As Date
        Get
            Return mvardtFim
        End Get
        Set(ByVal value As Date)
            mvardtFim = value
        End Set
    End Property

    Private mvarnomeDocto1 As String
    Public Property nomeDocto1() As String
        Get
            Return mvarnomeDocto1
        End Get
        Set(ByVal value As String)
            mvarnomeDocto1 = value
        End Set
    End Property

    Private mvarnomeDocto2 As String
    Public Property nomeDocto2() As String
        Get
            Return mvarnomeDocto2
        End Get
        Set(ByVal value As String)
            mvarnomeDocto2 = value
        End Set
    End Property

    Private mvaridEmpresa As Integer
    Public Property idEmpresa() As Integer
        Get
            Return mvaridEmpresa
        End Get
        Set(ByVal value As Integer)
            mvaridEmpresa = value
        End Set
    End Property


    Private mvaridDepto As Integer
    Public Property idDepto() As Integer
        Get
            Return mvaridDepto
        End Get
        Set(ByVal value As Integer)
            mvaridDepto = value
        End Set
    End Property

    Private mvaridProjeto As Integer
    Public Property idProjeto() As Integer
        Get
            Return mvaridProjeto
        End Get
        Set(ByVal value As Integer)
            mvaridProjeto = value
        End Set
    End Property

    Private mvartpDiretriz As Integer
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
        '* * 0 = OpenTable                        *
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
            'MsgBox(System.AppDomain.CurrentDomain.BaseDirectory())
            'MsgBox(strConnect)
            'strConnect = "driver={SQL Server};server=localhost\SQLEXPRESS;database=auditeste;uid=sa;pwd=auditeste"
            'strConnect = "driver={SQL Server};server=localhost\SQLEXPRESS;database=auditeste;uid=sa;pwd=audi952637" 'locahost Zé Antonio
            'strConnect = "Provider=SQLOLEDB;Data Source=AudiTeste051,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=audi952637;" 'Zé Antonio mesmo acima para VM
            'strConnect = "Provider=SQLOLEDB;Data Source=192.168.1.106,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=audi952637;" 'Zé Antonio mesmo acima para VM
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tDirDoctosDir = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tDirDoctosDir.Open("DirDoctosDir", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic,)
        Else                'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tDirDoctosDir.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tDirDoctosDir.Close()

                If tipo = 1 Or tipo = 2 Then
                    tDirDoctosDir.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tDirDoctosDir.Open("DirDoctosDir", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, ) 'adCmdTableDirect
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe DirDoctosDir - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tDirDoctosDir.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tDirDoctosDir.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tDirDoctosDir.EOF Then
            idDiretriz = tDirDoctosDir.Fields("idDiretriz").Value
            nomeDiretriz = tDirDoctosDir.Fields("nomeDiretriz").Value
            tpDoctoDiretriz = tDirDoctosDir.Fields("tpDoctoDiretriz").Value
            dtRegistro = tDirDoctosDir.Fields("dtRegistro").Value
            dtFim = tDirDoctosDir.Fields("dtFim").Value
            nomeDocto1 = tDirDoctosDir.Fields("nomeDocto1").Value
            If tDirDoctosDir(6).ActualSize > 0 Then
                If tDirDoctosDir.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tDirDoctosDir.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tDirDoctosDir.Fields("nomeDocto2").Value
            If tDirDoctosDir(9).ActualSize > 0 Then
                If tDirDoctosDir.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tDirDoctosDir.Fields("tipodocto2").Value
                End If
            Else
                tipodocto2 = " "
            End If
            idEmpresa = tDirDoctosDir.Fields("idEmpresa").Value
            idDepto = tDirDoctosDir.Fields("idDepto").Value
            idProjeto = tDirDoctosDir.Fields("idProjeto").Value
            tpDiretriz = tDirDoctosDir.Fields("tpDiretriz").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe DirDoctosDir - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(diretriz As Double, atualiz As Integer) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * DirDoctosDir.                        *
        '* *                                      *
        '* * diretriz   = identif. para pesquisa +*
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM DirDoctosDir WHERE idDiretriz = " & diretriz
        tDirDoctosDir.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tDirDoctosDir.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idDiretriz = tDirDoctosDir.Fields("idDiretriz").Value
            nomeDiretriz = tDirDoctosDir.Fields("nomeDiretriz").Value
            tpDoctoDiretriz = tDirDoctosDir.Fields("tpDoctoDiretriz").Value
            dtRegistro = tDirDoctosDir.Fields("dtRegistro").Value
            dtFim = tDirDoctosDir.Fields("dtFim").Value
            nomeDocto1 = tDirDoctosDir.Fields("nomeDocto1").Value
            If tDirDoctosDir(6).ActualSize > 0 Then 'se arquivo existe
                If tDirDoctosDir.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tDirDoctosDir.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tDirDoctosDir.Fields("nomeDocto2").Value
            If tDirDoctosDir(9).ActualSize > 0 Then
                If tDirDoctosDir.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tDirDoctosDir.Fields("tipodocto2").Value
                End If
            Else
                tipodocto2 = " "
            End If
            idEmpresa = tDirDoctosDir.Fields("idEmpresa").Value
            idDepto = tDirDoctosDir.Fields("idDepto").Value
            idProjeto = tDirDoctosDir.Fields("idProjeto").Value
            tpDiretriz = tDirDoctosDir.Fields("tpDiretriz").Value
        End If
Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tDirDoctosDir.Close()
                tDirDoctosDir.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe DirDoctosDir - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtDocto As String) As Integer

        On Error GoTo Einclui
        tDirDoctosDir.AddNew()
        tDirDoctosDir(0).Value = idDiretriz
        tDirDoctosDir(1).Value = nomeDiretriz
        tDirDoctosDir(2).Value = tpDoctoDiretriz
        tDirDoctosDir(3).Value = dtRegistro
        tDirDoctosDir(4).Value = dtFim
        tDirDoctosDir(5).Value = IIf(txtDocto = "", " ", txtDocto)
        If doc1 = True Then
            ArmazenaDB(path1, 6)
            tDirDoctosDir(7).Value = nomeDocto1
            tDirDoctosDir(8).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tDirDoctosDir(6).Value = ""
            tDirDoctosDir(7).Value = " "
            tDirDoctosDir(8).Value = " "
        End If
        If doc2 = True Then
            ArmazenaDB(path2, 9)
            tDirDoctosDir(10).Value = nomeDocto2
            tDirDoctosDir(11).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tDirDoctosDir(9).Value = ""
            tDirDoctosDir(10).Value = " "
            tDirDoctosDir(11).Value = " "
        End If
        tDirDoctosDir(12).Value = idEmpresa
        tDirDoctosDir(13).Value = idDepto
        tDirDoctosDir(14).Value = idProjeto
        tDirDoctosDir(15).Value = tpDiretriz
        tDirDoctosDir.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe DirDoctosDir - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtDocto As String) As Integer

        On Error GoTo Ealtera
        tDirDoctosDir(1).Value = nomeDiretriz
        tDirDoctosDir(2).Value = tpDoctoDiretriz
        tDirDoctosDir(3).Value = dtRegistro
        tDirDoctosDir(4).Value = dtFim
        tDirDoctosDir(5).Value = IIf(txtDocto = "", " ", txtDocto)
        If doc1 = True Then
            If path1 <> "" Then
                ArmazenaDB(path1, 6)
            End If
            tDirDoctosDir(7).Value = nomeDocto1
            tDirDoctosDir(8).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tDirDoctosDir(6).Value = ""
            tDirDoctosDir(7).Value = " "
            tDirDoctosDir(8).Value = " "
        End If
        If doc2 = True Then
            If path2 <> "" Then
                ArmazenaDB(path2, 9)
            End If
            tDirDoctosDir(10).Value = nomeDocto2
            tDirDoctosDir(11).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tDirDoctosDir(9).Value = ""
            tDirDoctosDir(10).Value = " "
            tDirDoctosDir(11).Value = " "
        End If
        tDirDoctosDir(12).Value = idEmpresa
        tDirDoctosDir(13).Value = idDepto
        tDirDoctosDir(14).Value = idProjeto
        tDirDoctosDir(15).Value = tpDiretriz
        tDirDoctosDir.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe DirDoctosDir - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaMemos() As String

        CarregaMemos = tDirDoctosDir.Fields("txtDocto").Value

    End Function

    Public Sub ArmazenaDB(nomeTabF As String, colF As Integer)
        'Grava o conteúdo do arquivo no campo binário

        Dim bytes = My.Computer.FileSystem.ReadAllBytes(nomeTabF)

        tDirDoctosDir(colF).AppendChunk(bytes)

    End Sub

    Public Sub GravaArq(colF As Integer, nomeTabF As String)
        'Extrai o conteúdo de um campo binário, em um arquivo

        'Define as variáveis:
        Dim ChunkSize As Integer = 32000
        Dim CurChunk() As Byte
        Dim fs As FileStream
        Dim writer As BinaryWriter

        fs = File.Open(nomeTabF, FileMode.Create)
        writer = New BinaryWriter(fs)
        Do
            CurChunk = tDirDoctosDir(colF).GetChunk(ChunkSize)
            ' Write each byte
            For Each value As Byte In CurChunk
                writer.Write(value)
            Next
            If CurChunk.Length < ChunkSize Then Exit Do
        Loop
        writer.Close()
        fs.Close()

    End Sub

    Public Function elimina(diretriz As Double) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * DirDoctosDir.                        *
        '* *                                      *
        '* * diretriz   = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM DirDoctosDir WHERE "
        vSelec = vSelec & " idDiretriz = " & diretriz
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe DirDoctosDir - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class



