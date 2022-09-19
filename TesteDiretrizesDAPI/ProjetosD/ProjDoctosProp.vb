Option Explicit On
Imports System.IO

Public Class ProjDoctosProp
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850713"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320713"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0713"
#End Region

#Region "Variaveis de ambiente"

    Private mvaridProposta As String
    Public Property idProposta() As String
        Get
            Return mvaridProposta
        End Get
        Set(value As String)
            mvaridProposta = value
        End Set
    End Property

    Private mvartpDoctoProposta As Short
    Public Property tpDoctoProposta() As Short
        Get
            Return mvartpDoctoProposta
        End Get
        Set(value As Short)
            mvartpDoctoProposta = value
        End Set
    End Property

    Private mvardtRegistro As Date
    Public Property dtRegistro() As Date
        Get
            Return mvardtRegistro
        End Get
        Set(value As Date)
            mvardtRegistro = value
        End Set
    End Property

    Private mvartipodocto1 As String
    Public Property tipodocto1() As String
        Get
            Return mvartipodocto1
        End Get
        Set(value As String)
            mvartipodocto1 = value
        End Set
    End Property

    Private mvartipodocto2 As String
    Public Property tipodocto2() As String
        Get
            Return mvartipodocto2
        End Get
        Set(value As String)
            mvartipodocto2 = value
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

    Private mvarnomeDocto1 As String
    Public Property nomeDocto1() As String
        Get
            Return mvarnomeDocto1
        End Get
        Set(value As String)
            mvarnomeDocto1 = value
        End Set
    End Property
    Private mvarnomeDocto2 As String
    Public Property nomeDocto2() As String
        Get
            Return mvarnomeDocto2
        End Get
        Set(value As String)
            mvarnomeDocto2 = value
        End Set
    End Property
#End Region

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tProjDoctosProp As ADODB.Recordset
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
        tProjDoctosProp = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tProjDoctosProp.Open("ProjDoctosProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 1 Then    'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tProjDoctosProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            Else
                If tipo = 3 Then    'Fecha a tabela
                    tProjDoctosProp.Close()
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjDoctosProp.Close()

                If tipo = 1 Then
                    tProjDoctosProp.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tProjDoctosProp.Open("ProjDoctosProp", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                If tipo <> 3 And Err.Number = 3704 Then
                    dbConecta = Err.Number
                    If Err.Number = 3024 Then dbConecta = 1046
                    MsgBox("Classe ProjDoctosProp - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
                Else
                    Resume Next
                End If
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
        If Not tProjDoctosProp.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tProjDoctosProp.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tProjDoctosProp.EOF Then
            idProposta = tProjDoctosProp.Fields("idProposta").Value
            tpDoctoProposta = tProjDoctosProp.Fields("tpDoctoProposta").Value
            dtRegistro = tProjDoctosProp.Fields("dtRegistro").Value
            dtFim = tProjDoctosProp.Fields("dtFim").Value
            nomeDocto1 = tProjDoctosProp.Fields("nomeDocto1").Value
            If (tProjDoctosProp(5).ActualSize) > 0 Then
                If tProjDoctosProp.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tProjDoctosProp.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tProjDoctosProp.Fields("nomeDocto2").Value
            If (tProjDoctosProp(8).ActualSize) > 0 Then
                If tProjDoctosProp.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tProjDoctosProp.Fields("tipodocto2").Value
                End If
            Else
                tipodocto2 = " "
            End If
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe ProjDoctosProp - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(proposta As String, tipo As Short, vData As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ProjDoctosProp.                      *
        '* *                                      *
        '* * projeto   = identif. para pesquisa + *
        '* * tipo       = identif. para pesquisa+ *
        '* * vdata      = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ProjDoctosProp WHERE idProposta = '" & proposta & "'"
        vSelect = vSelect & " AND tpDoctoProposta = " & tipo
        vSelect = vSelect & " AND Convert(datetime, dtRegistro, 112) = '" & String.Format(vData, "yyyymmdd") & "'"
        tProjDoctosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tProjDoctosProp.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idProposta = tProjDoctosProp.Fields("idProposta").Value
            tpDoctoProposta = tProjDoctosProp.Fields("tpDoctoProposta").Value
            dtRegistro = tProjDoctosProp.Fields("dtRegistro").Value
            dtFim = tProjDoctosProp.Fields("dtFim").Value
            nomeDocto1 = tProjDoctosProp.Fields("nomeDocto1").Value
            If (tProjDoctosProp(5).ActualSize) > 0 Then
                If tProjDoctosProp.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tProjDoctosProp.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tProjDoctosProp.Fields("nomeDocto2").Value
            If (tProjDoctosProp(8).ActualSize) > 0 Then
                If tProjDoctosProp.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tProjDoctosProp.Fields("tipodocto2").Value
                End If
            Else
                tipodocto2 = " "
            End If
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tProjDoctosProp.Close()
                tProjDoctosProp.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ProjDoctosProp - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtDocto As String) As Integer

        On Error GoTo Einclui
        tProjDoctosProp.AddNew()
        tProjDoctosProp(0).Value = idProposta
        tProjDoctosProp(1).Value = tpDoctoProposta
        tProjDoctosProp(2).Value = dtRegistro
        tProjDoctosProp(3).Value = dtFim
        tProjDoctosProp(4).Value = IIf(txtDocto = "", " ", txtDocto)
        If doc1 = True Then
            ArmazenaDB(path1, 5)
            tProjDoctosProp(6).Value = nomeDocto1
            tProjDoctosProp(7).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tProjDoctosProp(5).Value = ""
            tProjDoctosProp(6).Value = " "
            tProjDoctosProp(7).Value = " "
        End If
        If doc2 = True Then
            ArmazenaDB(path2, 8)
            tProjDoctosProp(9).Value = nomeDocto2
            tProjDoctosProp(10).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tProjDoctosProp(8).Value = ""
            tProjDoctosProp(9).Value = " "
            tProjDoctosProp(10).Value = " "
        End If
        tProjDoctosProp.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe ProjDoctosProp - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtDocto As String) As Integer

        On Error GoTo Ealtera
        tProjDoctosProp(2).Value = dtRegistro
        tProjDoctosProp(3).Value = dtFim
        tProjDoctosProp(4).Value = IIf(txtDocto = "", " ", txtDocto)
        If doc1 = True Then
            If path1 <> "" Then
                ArmazenaDB(path1, 5)
            End If
            tProjDoctosProp(6).Value = nomeDocto1
            tProjDoctosProp(7).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tProjDoctosProp(5).Value = ""
            tProjDoctosProp(6).Value = " "
            tProjDoctosProp(7).Value = " "
        End If
        If doc2 = True Then
            If path2 <> "" Then
                ArmazenaDB(path2, 8)
            End If
            tProjDoctosProp(9).Value = nomeDocto2
            tProjDoctosProp(10).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tProjDoctosProp(8).Value = ""
            tProjDoctosProp(9).Value = " "
            tProjDoctosProp(10).Value = " "
        End If
        tProjDoctosProp.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ProjDoctosProp - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaMemos() As String

        CarregaMemos = tProjDoctosProp.Fields("txtDocto").Value

    End Function

    Public Sub ArmazenaDB(nomeTabF As String, colF As Short)

        Dim bytes = My.Computer.FileSystem.ReadAllBytes(nomeTabF)

        tProjDoctosProp(colF).AppendChunk(bytes)

    End Sub

    Public Sub GravaArq(colF As Short, nomeTabF As String)
        'Extrai o conteúdo de um campo binário, em um arquivo

        'Define as variáveis:
        Dim J As Short
        Dim ChunkSize As Integer
        Dim CurSize As Integer
        'Dim CurChunk As String

        Dim CurChunk() As Byte
        Dim fs As FileStream
        Dim writer As BinaryWriter

        On Error Resume Next
        ChunkSize = 8192    'Define o tamanho de cada pedaço
        J = FreeFile()        'Obtém um número livre de arquivo
        'Abre o arquivo como binário:
        fs = File.Open(nomeTabF, FileMode.Create)
        writer = New BinaryWriter(fs)
        Do
            CurChunk = tProjDoctosProp(colF).GetChunk(ChunkSize * 2)
            ' Write each byte
            For Each value As Byte In CurChunk
                writer.Write(value)
            Next
            If CurChunk.Length < ChunkSize Then Exit Do
        Loop
        writer.Close()
        fs.Close()

    End Sub

    Public Function elimina(proposta As String, tipo As Short, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ProjDoctosProp.                      *
        '* *                                      *
        '* * proposta   = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ProjDoctosProp WHERE "
        vSelec = vSelec & " idProposta = '" & proposta & "'"
        vSelec = vSelec & " AND tpDoctoProposta = " & tipo
        vSelec = vSelec & " AND Convert(datetime, dtRegistro, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ProjDoctosProp - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function



End Class
