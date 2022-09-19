Option Explicit On
Imports System.IO

Public Class EmprDoctos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850309"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320309"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0309"
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


    Private mvaridTpDocto As Short
    Public Property idTpDocto() As Short
        Get
            Return mvaridTpDocto
        End Get
        Set(value As Short)
            mvaridTpDocto = value
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


    Private mvartipodocto1 As String
    Public Property tipodocto1() As String
        Get
            Return mvartipodocto1
        End Get
        Set(value As String)
            mvartipodocto1 = value
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


    Private mvartipodocto2 As String
    Public Property tipodocto2() As String
        Get
            Return mvartipodocto2
        End Get
        Set(value As String)
            mvartipodocto2 = value
        End Set
    End Property



#End Region

#Region "Conexão com Banco"
    Public tEmprDoctos As ADODB.Recordset
    Public Db As ADODB.Connection
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
        tEmprDoctos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            tEmprDoctos.Open("EmprDoctos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 1 Then    'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprDoctos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            Else
                If tipo = 3 Then    'Fecha a tabela
                    tEmprDoctos.Close()
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprDoctos.Close()

                If tipo = 1 Then
                    tEmprDoctos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprDoctos.Open("EmprDoctos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                If tipo <> 3 And Err.Number = 3704 Then
                    dbConecta = Err.Number
                    If Err.Number = 3024 Then dbConecta = 1046
                    MsgBox("Classe EmprDoctos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprDoctos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprDoctos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprDoctos.EOF Then
            idEmpresa = tEmprDoctos.Fields("idEmpresa").Value
            idTpDocto = tEmprDoctos.Fields("idTpDocto").Value
            dtRegistro = tEmprDoctos.Fields("dtRegistro").Value
            dtFim = tEmprDoctos.Fields("dtFim").Value
            nomeDocto1 = tEmprDoctos.Fields("nomeDocto1").Value
            If Len(tEmprDoctos(5)) > 0 Then
                If tEmprDoctos.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tEmprDoctos.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tEmprDoctos.Fields("nomeDocto2").Value
            If Len(tEmprDoctos(8)) > 0 Then
                If tEmprDoctos.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tEmprDoctos.Fields("tipodocto2").Value
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
            MsgBox("Classe EmprDoctos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Double, tipo As Short, vData As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprDoctos.                          *
        '* *                                      *
        '* * empresa  = identif. para pesquisa +  *
        '* * tipo       = identif. para pesquisa+ *
        '* * vdata      = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprDoctos WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idTpDocto = " & tipo
        vSelect = vSelect & " AND Convert(datetime, dtRegistro, 112) = '" & Format(vData, "yyyymmdd") & "'"
        tEmprDoctos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprDoctos.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprDoctos.Fields("idEmpresa").Value
            idTpDocto = tEmprDoctos.Fields("idTpDocto").Value
            dtRegistro = tEmprDoctos.Fields("dtRegistro").Value
            dtFim = tEmprDoctos.Fields("dtFim").Value
            nomeDocto1 = tEmprDoctos.Fields("nomeDocto1").Value
            If Len(tEmprDoctos(5)) > 0 Then
                If tEmprDoctos.Fields("tipodocto1").Value = " " Then
                    tipodocto1 = "ZZZZZ"
                Else
                    tipodocto1 = tEmprDoctos.Fields("tipodocto1").Value
                End If
            Else
                tipodocto1 = " "
            End If
            nomeDocto2 = tEmprDoctos.Fields("nomeDocto2").Value
            If Len(tEmprDoctos(8)) > 0 Then
                If tEmprDoctos.Fields("tipodocto2").Value = " " Then
                    tipodocto2 = "ZZZZZ"
                Else
                    tipodocto2 = tEmprDoctos.Fields("tipodocto2").Value
                End If
            Else
                tipodocto2 = " "
            End If
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprDoctos.Close()
                tEmprDoctos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprDoctos - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtdocto As String) As Integer

        On Error GoTo Einclui
        tEmprDoctos.AddNew()
        tEmprDoctos(0).Value = idEmpresa
        tEmprDoctos(1).Value = idTpDocto
        tEmprDoctos(2).Value = dtRegistro
        tEmprDoctos(3).Value = dtFim
        tEmprDoctos(4).Value = IIf(txtdocto = "", " ", txtdocto)
        If doc1 = True Then
            ArmazenaDB(path1, 5)
            tEmprDoctos(6).Value = nomeDocto1
            tEmprDoctos(7).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tEmprDoctos(5).Value = ""
            tEmprDoctos(6).Value = " "
            tEmprDoctos(7).Value = " "
        End If
        If doc2 = True Then
            ArmazenaDB(path2, 8)
            tEmprDoctos(9).Value = nomeDocto2
            tEmprDoctos(10).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tEmprDoctos(8).Value = ""
            tEmprDoctos(9).Value = " "
            tEmprDoctos(10).Value = " "
        End If
        tEmprDoctos.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprDoctos - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(doc1 As Boolean, path1 As String, doc2 As Boolean, path2 As String, txtdocto As String) As Integer

        On Error GoTo Ealtera
        tEmprDoctos(0).Value = idEmpresa
        tEmprDoctos(1).Value = idTpDocto
        tEmprDoctos(2).Value = dtRegistro
        tEmprDoctos(3).Value = dtFim
        tEmprDoctos(4).Value = IIf(txtdocto = "", " ", txtdocto)
        If doc1 = True Then
            If path1 <> "" Then
                ArmazenaDB(path1, 5)
            End If
            tEmprDoctos(6).Value = nomeDocto1
            tEmprDoctos(7).Value = IIf(tipodocto1 = "ZZZZZ", " ", tipodocto1)
        Else
            tEmprDoctos(5).Value = ""
            tEmprDoctos(6).Value = " "
            tEmprDoctos(7).Value = " "
        End If
        If doc2 = True Then
            If path2 <> "" Then
                ArmazenaDB(path2, 8)
            End If
            tEmprDoctos(9).Value = nomeDocto2
            tEmprDoctos(10).Value = IIf(tipodocto2 = "ZZZZZ", " ", tipodocto2)
        Else
            tEmprDoctos(8).Value = ""
            tEmprDoctos(9).Value = " "
            tEmprDoctos(10).Value = " "
        End If
        tEmprDoctos.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprDoctos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String

        carregaMemos = tEmprDoctos.Fields("txtdocto").Value

    End Function

    Public Sub ArmazenaDB(nomeTabF As String, colF As Short)
        'Grava o conteúdo do arquivo no campo binário

        'Define as variáveis:
        Dim TotalSize As Integer
        Dim CurChunk As String
        Dim ChunkSize As Integer
        Dim J As Short

        ChunkSize = 8192    'Define o tamanho de cada pedaço
        J = FreeFile()        'Obtém um número livre de arquivo
        'Abre o arquivo como binário:
        Open(nomeTabF) For Binary As #J
    TotalSize = LOF(J)
        Do While Not EOF(J)
            If TotalSize - Seek(J) < ChunkSize Then
                ChunkSize = TotalSize - Seek(J) + 10
            End If
            CurChunk = String$(ChunkSize + 1, 32)
        'Lê o pedaço do arquivo:
        Get #J, , CurChunk
        'Grava o pedaço no final da coluna:
        tEmprDoctos(colF).AppendChunk(CurChunk)
        Loop
        Close #J

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
            CurChunk = tEmprDoctos(colF).GetChunk(ChunkSize)
            ' Write each byte
            For Each value As Byte In CurChunk
                writer.Write(value)
            Next
            If CurChunk.Length < ChunkSize Then Exit Do
        Loop
        writer.Close()
        fs.Close()

    End Sub

    Public Function elimina(empresa As Short, tipo As Short, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprDoctos.                          *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * tipo       = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM EmprDoctos WHERE "
        vSelec = vSelec & " idEmpresa = " & empresa
        vSelec = vSelec & " AND idTpDocto = " & tipo
        vSelec = vSelec & " AND Convert(datetime, dtRegistro, 112) = '" & Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprDoctos - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
