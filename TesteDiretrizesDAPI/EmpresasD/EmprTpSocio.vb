Option Explicit On
Public Class EmprTpSocio
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850316"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320316"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0316"
#End Region

#Region "Variaveis de Ambiente"
    Private mvaridTipoSocio As Short
    Public Property idTipoSocio() As Short
        Get
            Return mvaridTipoSocio
        End Get
        Set(value As Short)
            mvaridTipoSocio = value
        End Set
    End Property


    Private mvardescricTipoSocio As String
    Public Property descricTipoSocio() As String
        Get
            Return mvardescricTipoSocio
        End Get
        Set(value As String)
            mvardescricTipoSocio = value
        End Set
    End Property


    Private mvartipo As Short
    Public Property tipo() As Short
        Get
            Return mvartipo
        End Get
        Set(value As Short)
            mvartipo = value
        End Set
    End Property

#End Region


#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tEmprTpSocio As ADODB.Recordset
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
        tEmprTpSocio = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tEmprTpSocio.Open("EmprTpSocio", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If tipo = 2 Then    'Aberto como OpenDynaset
                vSelec = "SELECT * FROM EmprTpSocio ORDER BY descricTipoSocio"
                tEmprTpSocio.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            Else            'Aberto como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprTpSocio.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprTpSocio.Close()

                If tipo = 1 Or tipo = 2 Then
                    tEmprTpSocio.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprTpSocio.Open("EmprTpSocio", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprTpSocio - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprTpSocio.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprTpSocio.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprTpSocio.EOF Then
            idTipoSocio = tEmprTpSocio.Fields("idTipoSocio").Value
            descricTipoSocio = tEmprTpSocio.Fields("descricTipoSocio").Value
            tipo = tEmprTpSocio.Fields("tipo").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprTpSocio - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(tpSocio As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprTpSocio.                         *
        '* *                                      *
        '* * tpSocio = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprTpSocio WHERE idTipoSocio = " & tpSocio
        tEmprTpSocio.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprTpSocio.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idTipoSocio = tEmprTpSocio.Fields("idTipoSocio").Value
            descricTipoSocio = tEmprTpSocio.Fields("descricTipoSocio").Value
            tipo = tEmprTpSocio.Fields("tipo").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprTpSocio.Close()
                tEmprTpSocio.Open(vSelect, Db, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprTpSocio - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprTpSocio.                         *
        '* *                                      *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM EmprTpSocio WHERE descricTipoSocio = '" & descric & "'"
        tEmprTpSocio.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprTpSocio.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idTipoSocio = tEmprTpSocio.Fields("idTipoSocio").Value
            descricTipoSocio = tEmprTpSocio.Fields("descricTipoSocio").Value
            tipo = tEmprTpSocio.Fields("tipo").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprTpSocio.Close()
                tEmprTpSocio.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe EmprTpSocio - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(cCodigo As Short, cDescricao As String) As Integer

        On Error GoTo Einclui
        tEmprTpSocio.AddNew()
        tEmprTpSocio(0).Value = cCodigo
        tEmprTpSocio(1).Value = cDescricao
        tEmprTpSocio.Fields("tipo").Value = tipo
        tEmprTpSocio.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprTpSocio - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(cDescricao As String) As Integer

        On Error GoTo Ealtera
        tEmprTpSocio(1).Value = cDescricao
        tEmprTpSocio.Fields("tipo").Value = tipo
        tEmprTpSocio.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprTpSocio - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(codigo As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprTpSocio.                         *
        '* *                                      *
        '* * codigo     = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro ao abrir segMensagens")
            elimina = -1
        End If

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em EmprCompSoc:
        vSelec = "SELECT * FROM EmprCompSoc WHERE tpSocio = " & codigo
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            elimina = RC
            Exit Function
        End If

        vSelec = "DELETE FROM EmprTpSocio WHERE "
        vSelec = vSelec & " idTipoSocio = " & codigo
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprTpSocio - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
