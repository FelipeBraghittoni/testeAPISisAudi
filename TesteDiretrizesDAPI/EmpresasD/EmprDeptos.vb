Option Explicit On
Public Class EmprDeptos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850308"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320308"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0308"
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


    Private mvarnomeDepto As String
    Public Property nomeDepto() As String
        Get
            Return mvarnomeDepto
        End Get
        Set(value As String)
            mvarnomeDepto = value
        End Set
    End Property


    Private mvarfone1Depto As Integer
    Public Property fone1Depto() As Integer
        Get
            Return mvarfone1Depto
        End Get
        Set(value As Integer)
            mvarfone1Depto = value
        End Set
    End Property

    Private mvarramal1Depto As Short
    Public Property ramal1Depto() As Short
        Get
            Return mvarramal1Depto
        End Get
        Set(value As Short)
            mvarramal1Depto = value
        End Set
    End Property


    Private mvarfone2Depto As Integer
    Public Property fone2Depto() As Integer
        Get
            Return mvarfone2Depto
        End Get
        Set(value As Integer)
            mvarfone2Depto = value
        End Set
    End Property


    Private mvarramal2Depto As Short
    Public Property ramal2Depto() As Short
        Get
            Return mvarfone2Depto
        End Get
        Set(value As Short)
            mvarfone2Depto = value
        End Set
    End Property


    Private mvarfaxDepto As Integer
    Public Property faxDepto() As Integer
        Get
            Return mvarfaxDepto
        End Get
        Set(value As Integer)
            mvarfaxDepto = value
        End Set
    End Property


    Private mvaridDeptoSup As Short
    Public Property idDeptoSup() As Short
        Get
            Return mvaridDeptoSup
        End Get
        Set(value As Short)
            mvaridDeptoSup = value
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


    Private mvarregimeTrabDepto As Short
    Public Property regimeTrabDepto() As Short
        Get
            Return mvarregimeTrabDepto
        End Get
        Set(value As Short)
            mvarregimeTrabDepto = value
        End Set
    End Property


    Private mvardiaInicTrabDepto As Short
    Public Property diaInicTrabDepto() As Short
        Get
            Return mvardiaInicTrabDepto
        End Get
        Set(value As Short)
            mvardiaInicTrabDepto = value
        End Set
    End Property


    Private mvardiaFimTrabDepto As Short
    Public Property diaFimTrabDepto() As Short
        Get
            Return mvardiaFimTrabDepto
        End Get
        Set(value As Short)
            mvardiaFimTrabDepto = value
        End Set
    End Property


    Private mvarrefeitorioDepto As String
    Public Property refeitorioDepto() As String
        Get
            Return mvardiaFimTrabDepto
        End Get
        Set(value As String)
            mvardiaFimTrabDepto = value
        End Set
    End Property


    Private mvarhoraInicTrabDepto As String
    Public Property horaInicTrabDepto() As String
        Get
            Return mvarhoraInicTrabDepto
        End Get
        Set(value As String)
            mvarhoraInicTrabDepto = value
        End Set
    End Property


    Private mvarhoraInicDescDepto As String
    Public Property horaInicDescDepto() As String
        Get
            Return mvarhoraInicDescDepto
        End Get
        Set(value As String)
            mvarhoraInicDescDepto = value
        End Set
    End Property


    Private mvarhoraFimDescDepto As String
    Public Property horaFimDescDepto() As String
        Get
            Return mvarhoraFimDescDepto
        End Get
        Set(value As String)
            mvarhoraFimDescDepto = value
        End Set
    End Property


    Private mvarhoraFimTrabDepto As String
    Public Property horaFimTrabDepto() As String
        Get
            Return mvarhoraFimTrabDepto
        End Get
        Set(value As String)
            mvarhoraFimTrabDepto = value
        End Set
    End Property


    Private mvaridCtaContabil As String
    Public Property idCtaContabil() As String
        Get
            Return mvaridCtaContabil
        End Get
        Set(value As String)
            mvaridCtaContabil = value
        End Set
    End Property


    Private mvareMailDepto As String
    Public Property eMailDepto() As String
        Get
            Return mvareMailDepto
        End Get
        Set(value As String)
            mvareMailDepto = value
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


#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tEmprDeptos As ADODB.Recordset
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
        tEmprDeptos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            tEmprDeptos.Open("EmprDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM EmprDeptos ORDER BY idEmpresa, nomeDepto"
                tEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprDeptos.Close()

                If tipo = 1 Or tipo = 2 Then
                    tEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tEmprDeptos.Open("EmprDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe EmprDeptos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not tEmprDeptos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tEmprDeptos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tEmprDeptos.EOF Then
            idEmpresa = tEmprDeptos.Fields("idEmpresa").Value
            idDepto = tEmprDeptos.Fields("idDepto").Value
            nomeDepto = tEmprDeptos.Fields("nomeDepto").Value
            idCtaContabil = tEmprDeptos.Fields("idCtaCtbil").Value
            fone1Depto = tEmprDeptos.Fields("fone1Depto").Value
            ramal1Depto = tEmprDeptos.Fields("ramal1Depto").Value
            fone2Depto = tEmprDeptos.Fields("fone2Depto").Value
            ramal2Depto = tEmprDeptos.Fields("ramal2Depto").Value
            faxDepto = tEmprDeptos.Fields("faxDepto").Value
            eMailDepto = tEmprDeptos.Fields("eMailDepto").Value
            idDeptoSup = tEmprDeptos.Fields("idDeptoSup").Value
            idColabResp = tEmprDeptos.Fields("idColabResp").Value
            regimeTrabDepto = tEmprDeptos.Fields("regimeTrabDepto").Value
            diaInicTrabDepto = tEmprDeptos.Fields("diaInicTrabDepto").Value
            diaFimTrabDepto = tEmprDeptos.Fields("diaFimTrabDepto").Value
            refeitorioDepto = tEmprDeptos.Fields("refeitorioDepto").Value
            horaInicTrabDepto = tEmprDeptos.Fields("horaInicTrabDepto").Value
            horaInicDescDepto = tEmprDeptos.Fields("horaInicDescDepto").Value
            horaFimDescDepto = tEmprDeptos.Fields("horaFimDescDepto").Value
            horaFimTrabDepto = tEmprDeptos.Fields("horaFimTrabDepto").Value
            dtInicio = tEmprDeptos.Fields("dtInicio").Value
            dtFim = tEmprDeptos.Fields("dtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe EmprDeptos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function localiza(empresa As Short, depto As Short, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprDeptos.                          *
        '* *                                      *
        '* * empresa = identif. para pesquisa +   *
        '* * depto   = identif. para pesquisa     *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM EmprDeptos WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND idDepto = " & depto
        tEmprDeptos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprDeptos.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprDeptos.Fields("idEmpresa").Value
            idDepto = tEmprDeptos.Fields("idDepto").Value
            nomeDepto = tEmprDeptos.Fields("nomeDepto").Value
            idCtaContabil = tEmprDeptos.Fields("idCtaCtbil").Value
            fone1Depto = tEmprDeptos.Fields("fone1Depto").Value
            ramal1Depto = tEmprDeptos.Fields("ramal1Depto").Value
            fone2Depto = tEmprDeptos.Fields("fone2Depto").Value
            ramal2Depto = tEmprDeptos.Fields("ramal2Depto").Value
            faxDepto = tEmprDeptos.Fields("faxDepto").Value
            eMailDepto = tEmprDeptos.Fields("eMailDepto").Value
            idDeptoSup = tEmprDeptos.Fields("idDeptoSup").Value
            idColabResp = tEmprDeptos.Fields("idColabResp").Value
            regimeTrabDepto = tEmprDeptos.Fields("regimeTrabDepto").Value
            diaInicTrabDepto = tEmprDeptos.Fields("diaInicTrabDepto").Value
            diaFimTrabDepto = tEmprDeptos.Fields("diaFimTrabDepto").Value
            refeitorioDepto = tEmprDeptos.Fields("refeitorioDepto").Value
            horaInicTrabDepto = tEmprDeptos.Fields("horaInicTrabDepto").Value
            horaInicDescDepto = tEmprDeptos.Fields("horaInicDescDepto").Value
            horaFimDescDepto = tEmprDeptos.Fields("horaFimDescDepto").Value
            horaFimTrabDepto = tEmprDeptos.Fields("horaFimTrabDepto").Value
            dtInicio = tEmprDeptos.Fields("dtInicio").Value
            dtFim = tEmprDeptos.Fields("dtFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprDeptos.Close()
                tEmprDeptos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe EmprDeptos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function localizaNome(empresa As Short, descric As String, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * EmprDeptos.                          *
        '* *                                      *
        '* * empresa = argum. p/pesquisa +        *
        '* * descric = argum. p/pesquisa por nome *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo ElocalizaNome
        vSelect = "SELECT * FROM EmprDeptos WHERE idEmpresa = " & empresa
        vSelect = vSelect & " AND nomeDepto = '" & descric & "'"
        tEmprDeptos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localizaNome = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tEmprDeptos.EOF Then
            localizaNome = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idEmpresa = tEmprDeptos.Fields("idEmpresa").Value
            idDepto = tEmprDeptos.Fields("idDepto").Value
            nomeDepto = tEmprDeptos.Fields("nomeDepto").Value
            idCtaContabil = tEmprDeptos.Fields("idCtaCtbil").Value
            fone1Depto = tEmprDeptos.Fields("fone1Depto").Value
            ramal1Depto = tEmprDeptos.Fields("ramal1Depto").Value
            fone2Depto = tEmprDeptos.Fields("fone2Depto").Value
            ramal2Depto = tEmprDeptos.Fields("ramal2Depto").Value
            faxDepto = tEmprDeptos.Fields("faxDepto").Value
            eMailDepto = tEmprDeptos.Fields("eMailDepto").Value
            idDeptoSup = tEmprDeptos.Fields("idDeptoSup").Value
            idColabResp = tEmprDeptos.Fields("idColabResp").Value
            regimeTrabDepto = tEmprDeptos.Fields("regimeTrabDepto").Value
            diaInicTrabDepto = tEmprDeptos.Fields("diaInicTrabDepto").Value
            diaFimTrabDepto = tEmprDeptos.Fields("diaFimTrabDepto").Value
            refeitorioDepto = tEmprDeptos.Fields("refeitorioDepto").Value
            horaInicTrabDepto = tEmprDeptos.Fields("horaInicTrabDepto").Value
            horaInicDescDepto = tEmprDeptos.Fields("horaInicDescDepto").Value
            horaFimDescDepto = tEmprDeptos.Fields("horaFimDescDepto").Value
            horaFimTrabDepto = tEmprDeptos.Fields("horaFimTrabDepto").Value
            dtInicio = tEmprDeptos.Fields("dtInicio").Value
            dtFim = tEmprDeptos.Fields("dtFim").Value
        End If

ElocalizaNome:
        If Err.Number Then
            If Err.Number = 3705 Then
                tEmprDeptos.Close()
                tEmprDeptos.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localizaNome = Err.Number
                MsgBox("Classe EmprDeptos - localizaNome" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function inclui(responsab As String) As Integer

        On Error GoTo Einclui
        tEmprDeptos.AddNew()
        tEmprDeptos(0).Value = idEmpresa
        tEmprDeptos(1).Value = idDepto
        tEmprDeptos(2).Value = nomeDepto
        tEmprDeptos(3).Value = idCtaContabil
        tEmprDeptos(4).Value = fone1Depto
        tEmprDeptos(5).Value = ramal1Depto
        tEmprDeptos(6).Value = fone2Depto
        tEmprDeptos(7).Value = ramal2Depto
        tEmprDeptos(8).Value = faxDepto
        tEmprDeptos(9).Value = eMailDepto
        tEmprDeptos(10).Value = idDeptoSup
        tEmprDeptos(11).Value = idColabResp
        tEmprDeptos(12).Value = IIf(responsab = "", " ", responsab)
        tEmprDeptos(13).Value = regimeTrabDepto
        tEmprDeptos(14).Value = diaInicTrabDepto
        tEmprDeptos(15).Value = diaFimTrabDepto
        tEmprDeptos(16).Value = refeitorioDepto
        tEmprDeptos(17).Value = horaInicTrabDepto
        tEmprDeptos(18).Value = horaInicDescDepto
        tEmprDeptos(19).Value = horaFimDescDepto
        tEmprDeptos(20).Value = horaFimTrabDepto
        tEmprDeptos.Fields("dtInicio").Value = dtInicio
        tEmprDeptos.Fields("dtFim").Value = dtFim
        tEmprDeptos.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe EmprDeptos - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function altera(responsab As String, indice As Integer) As Integer

        On Error GoTo Ealtera
        tEmprDeptos(2).Value = nomeDepto
        tEmprDeptos(3).Value = idCtaContabil
        tEmprDeptos(4).Value = fone1Depto
        tEmprDeptos(5).Value = ramal1Depto
        tEmprDeptos(6).Value = fone2Depto
        tEmprDeptos(7).Value = ramal2Depto
        tEmprDeptos(8).Value = faxDepto
        tEmprDeptos(9).Value = eMailDepto
        tEmprDeptos(10).Value = idDeptoSup
        tEmprDeptos(11).Value = idColabResp
        If indice = 602 Then    'Alteração no cadastro de deptos.:
            tEmprDeptos(12).Value = IIf(responsab = "", " ", responsab)
        End If
        tEmprDeptos(13).Value = regimeTrabDepto
        tEmprDeptos(14).Value = diaInicTrabDepto
        tEmprDeptos(15).Value = diaFimTrabDepto
        tEmprDeptos(16).Value = refeitorioDepto
        tEmprDeptos(17).Value = horaInicTrabDepto
        tEmprDeptos(18).Value = horaInicDescDepto
        tEmprDeptos(19).Value = horaFimDescDepto
        tEmprDeptos(20).Value = horaFimTrabDepto
        tEmprDeptos.Fields("dtInicio").Value = dtInicio
        tEmprDeptos.Fields("dtFim").Value = dtFim
        tEmprDeptos.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe EmprDeptos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String
        'Carrega o valor do campo txtResponsab no form:

        carregaMemos = tEmprDeptos.Fields("responsabilDepto").Value

    End Function

    Public Function elimina(empresa As Short, depto As Short) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * EmprDeptos.                          *
        '* *                                      *
        '* * empresa    = identif. para pesquisa  *
        '* * depto      = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        'Para verificar integridade referencial:
        RC = fPesquisa(0, empresa, depto)

        If RC = 0 Then
            vSelec = "DELETE FROM EmprDeptos WHERE "
            vSelec = vSelec & " idEmpresa = " & empresa
            vSelec = vSelec & " AND idDepto = " & depto
            Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            Select Case RC
                Case Is < 3     'BibAlocEstr
                    msgErro("BibAlocEstr")
                Case Is < 5     'BibHAlocEstr
                    msgErro("BibHAlocEstr")
                Case Is < 7     'CandVagas
                    msgErro("CandVagas")
                Case Is < 9     'OcorAlocProjetos
                    msgErro("OcorAlocProjetos")
                Case Is < 11    'OcorDeptoColab
                    msgErro("OcorDeptoColab")
                Case Is < 13    'EstDeposito
                    msgErro("EstDeposito")
                Case Is < 15    'GerFeriados
                    msgErro("GerFeriados")
                Case Is < 17    'InvAlocEmpr
                    msgErro("InvAlocEmpr")
                Case Is < 19    'InvHAlocEmpr
                    msgErro("InvHAlocEmpr")
                Case Is < 21    'ProjProjetos
                    msgErro("ProjProjetos")
                Case Is < 23    'RDRecDesp
                    msgErro("RDRecDesp")
                Case Is < 25    'RDHRecDesp
                    msgErro("RDHRecDesp")
                Case Is < 27    'RDRecDespEliminados
                    msgErro("RDRecDespEliminados")
                Case Is < 29    'RDNotasExplicat
                    msgErro("RDNotasExplicat")
                Case Is < 31    'RDSaldoRecDesp
                    msgErro("RDSaldoRecDesp")
                Case Is < 33    'RDHSaldoRecDesp
                    msgErro("RDHSaldoRecDesp")
            End Select
        End If

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe EmprDeptos - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


    Public Function fPesquisa(tipo As Short, empresa As Short, depto As Short, Optional data As Date = Nothing) As Integer
        '* *********************************************
        '* tipo = 0 ==> não pesquisa data de término   *
        '*              utilizado para eliminar linha  *
        '* tipo = 1 ==> pesquisa data de término       *
        '*              utilizado para encerrar depto. *
        '* *********************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim PesqReg As SegurancaD.FuncoesGerais
        Dim segMensagens As SegurancaD.segMensagens

        'Define uma instância para exibir mensagens de erro:
        segMensagens = New SegurancaD.segMensagens
        RC = segMensagens.dbConecta(0, 0)

        fPesquisa = 0     'Se não houver erro

        'Para verificar integridade referencial:
        PesqReg = New SegurancaD.FuncoesGerais

        'Verifica se não fere integridade
        'relacional, quando controlada por código:
        'Pesquisa em BibAlocEstr:
        vSelec = "SELECT * FROM BibAlocEstr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC
            Exit Function
        End If
        'Pesquisa em BibHAlocEstr:
        vSelec = "SELECT * FROM BibHAlocEstr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 2
            Exit Function
        End If
        'Pesquisa em CandVagas:
        vSelec = "SELECT * FROM CandVagas WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND idSetorOrg = " & depto
        vSelec = vSelec & " AND localTrabVaga = 1"
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 4
            Exit Function
        End If
        'Pesquisa em OcorAlocProjetos:
        vSelec = "SELECT * FROM OcorAlocProjetos WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND idSetorOrg = " & depto
        vSelec = vSelec & " AND localTrabColab = 1"
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProjeto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProjeto, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 6
            Exit Function
        End If
        'Pesquisa em OcorDeptoColab:
        vSelec = "SELECT * FROM OcorDeptoColab WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalDepto, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalDepto, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 8
            Exit Function
        End If
        'Pesquisa em EstDeposito:
        vSelec = "SELECT * FROM EstDeposito WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 10
            Exit Function
        End If
        'Pesquisa em GerFeriados:
        vSelec = "SELECT * FROM GerFeriados WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND idSetorOrg = " & depto
        vSelec = vSelec & " AND tpLocal = 1"
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(4), dataFeriado, 112) > '" & Format(data, "yyyy") & "'"
            vSelec = vSelec & " AND Convert(nvarchar(4), dataFeriado, 112) >= '" & Format(Now, "yyyy") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 12
            Exit Function
        End If
        'Pesquisa em InvAlocEmpr:
        vSelec = "SELECT * FROM InvAlocEmpr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 14
            Exit Function
        End If
        'Pesquisa em InvAlocEmpr:
        vSelec = "SELECT * FROM InvAlocEmpr WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFim, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFim, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 16
            Exit Function
        End If
        'Pesquisa em ProjProjetos:
        vSelec = "SELECT * FROM ProjProjetos WHERE idOrganiz = " & empresa
        vSelec = vSelec & " AND idSetorOrg = " & depto
        vSelec = vSelec & " AND localContrProj = 1"
        If tipo = 1 Then
            vSelec = vSelec & " AND (Convert(datetime, dtFinalProj, 112) = '18991230'"
            vSelec = vSelec & " OR Convert(datetime, dtFinalProj, 112) > '" & Format(data, "yyyymmdd") & "')"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 18
            Exit Function
        End If
        'Pesquisa em RDRecDesp:
        vSelec = "SELECT * FROM RDRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), competencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 20
            Exit Function
        End If
        'Pesquisa em RDHRecDesp:
        vSelec = "SELECT * FROM RDHRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), competencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 22
            Exit Function
        End If
        'Pesquisa em RDRecDespEliminados:
        vSelec = "SELECT * FROM RDRecDespEliminados WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), competencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 24
            Exit Function
        End If
        'Pesquisa em RDNotasExplicat:
        vSelec = "SELECT * FROM RDNotasExplicat WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), dtReferencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 26
            Exit Function
        End If
        'Pesquisa em RDSaldoRecDesp:
        vSelec = "SELECT * FROM RDSaldoRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), dtReferencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 28
            Exit Function
        End If
        'Pesquisa em RDHSaldoRecDesp:
        vSelec = "SELECT * FROM RDHSaldoRecDesp WHERE idEmpresa = " & empresa
        vSelec = vSelec & " AND idDepto = " & depto
        If tipo = 1 Then
            vSelec = vSelec & " AND Convert(nvarchar(6), dtReferencia, 112) >= '" & Format(data, "yyyymm") & "'"
        End If
        RC = PesqReg.pesqRegistros(vSelec)
        If RC <> 0 Then
            If RC = 1 Then segMensagens.exibeMsg(, 1067)
            If RC = 2 Then segMensagens.exibeMsg(, 1069)
            fPesquisa = RC + 30
            Exit Function
        End If

    End Function


    Private Sub msgErro(tabela As String)

        MsgBox("Há um departamento na tabela " & tabela & " que impede sua eliminação.")

    End Sub

End Class
