
Option Explicit On

Imports Microsoft.VisualBasic.FileIO
<ComClass(segUsuario.ClassId, segUsuario.InterfaceId, segUsuario.EventsId)>
Public Class segUsuario

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859009"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329009"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9009"
#End Region

    'Carrega uma instância dos objetos a serem carregados:
    Private cdSeguranca1 As New SegurancaD.cdSeguranca1
    Private segmensagens As New SegurancaD.segMensagens
    Private FuncoesGerais As New SegurancaD.FuncoesGerais
    Private dbUtils As New Utilidades.DBUtils



    'Private SegUsuario As New SegurancaD.SegUsuario


    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection

    Public RC = dbUtils.RC


    Public tUsuario As ADODB.Recordset


    Private mvarindice As Integer
    Public Property indice() As Integer
        Get
            Return mvarindice
        End Get
        Set(ByVal value As Integer)
            mvarindice = value
        End Set
    End Property

    Private mvarDescricao3 As String
    Public Property Descricao3() As String
        Get
            Return mvarDescricao3
        End Get
        Set(ByVal value As String)
            mvarDescricao3 = value
        End Set
    End Property

    Private mvarDescricao2 As String
    Public Property Descricao2() As String
        Get
            Return mvarDescricao2
        End Get
        Set(ByVal value As String)
            mvarDescricao2 = value
        End Set
    End Property

    Private mvarDescricao1 As String
    Public Property Descricao1() As String
        Get
            Return mvarDescricao1
        End Get
        Set(ByVal value As String)
            mvarDescricao1 = value
        End Set
    End Property

    Private mvarDescricao As String
    Public Property Descricao() As String
        Get
            Return mvarDescricao
        End Get
        Set(ByVal value As String)
            mvarDescricao = value
        End Set
    End Property

    Private mvarproc As String 'local copy
    <System.ComponentModel.Description("Define se usuário tem acesso ao serviço: 0 = não tem acesso ao serviço; 1 = tem acesso ao serviço.")>
    Public Property proc() As String
        Get
            Return mvarproc
        End Get
        Set(ByVal value As String)
            mvarproc = value
        End Set
    End Property

    Private mvartpUsuario As Integer
    Public Property tpUsuario() As Integer
        Get
            Return mvartpUsuario
        End Get
        Set(ByVal value As Integer)
            mvartpUsuario = value
        End Set
    End Property

    Private mvarnome As String 'local copy
    <System.ComponentModel.Description("Nome do usuário.")>
    Public Property nome() As String
        Get
            Return mvarnome
        End Get
        Set(ByVal value As String)
            mvarnome = value
        End Set
    End Property

    Private mvarfreqAltSenha As String 'local copy
    <System.ComponentModel.Description("Número de dias máximo para alterar a senha. Se 0, senha não tem prazo para expiração.")>
    Public Property freqAltSenha() As String
        Get
            Return mvarfreqAltSenha
        End Get
        Set(ByVal value As String)
            mvarfreqAltSenha = value
        End Set
    End Property

    Private mvardtStatus As String 'local copy
    <System.ComponentModel.Description("Data da última alteração de Status.")>
    Public Property dtStatus() As String
        Get
            Return mvardtStatus
        End Get
        Set(ByVal value As String)
            mvardtStatus = value
        End Set
    End Property

    Private mvarstatus As Integer 'local copy
    <System.ComponentModel.Description("Situação do usuário para acesso:\r\n0 - acesso normal\r\n1 - acesso bloqueado\r\n2 - novo usuário\r\n")>
    Public Property status() As String
        Get
            Return mvarstatus
        End Get
        Set(ByVal value As String)
            mvarstatus = value
        End Set
    End Property

    Private mvardtSenhaAtu As String 'local copy
    <System.ComponentModel.Description("Data da última alteração da senha atual, após ter sido descriptografada.")>
    Public Property dtSenhaAtu() As String
        Get
            Return mvardtSenhaAtu
        End Get
        Set(ByVal value As String)
            mvardtSenhaAtu = value
        End Set
    End Property

    Private mvarsenha3Ant As String 'local copy
    <System.ComponentModel.Description("Senha anterior, após ter sido descriptografada.")>
    Public Property senha3Ant() As String
        Get
            Return mvarsenha3Ant
        End Get
        Set(ByVal value As String)
            mvarsenha3Ant = value
        End Set
    End Property

    'Public senha2Ant As String
    Private mvarsenha2Ant As String
    Public Property senha2Ant() As String
        Get
            Return mvarsenha2Ant
        End Get
        Set(ByVal value As String)
            mvarsenha2Ant = value
        End Set
    End Property

    Private mvarsenhaAnt As String
    Public Property senhaAnt() As String
        Get
            Return mvarsenhaAnt
        End Get
        Set(ByVal value As String)
            mvarsenhaAnt = value
        End Set
    End Property

    Private mvarsenha As String 'local copy
    <System.ComponentModel.Description("Senha do usuário, após ter sido descriptografada.")>
    Public Property senha As String
        Get
            Return mvarsenha
        End Get
        Set(ByVal value As String)
            mvarsenha = value
        End Set
    End Property


    Private mvarUsuario As String 'local copy
    <System.ComponentModel.Description("Informação do usuário após descriptografar")>
    Public Property Usuario() As String
        Get
            Return mvarUsuario
        End Get
        Set(ByVal value As String)
            mvarUsuario = value
        End Set
    End Property

    Private mvarUsuarioA As String
    Public Property UsuarioA() As String
        Get
            Return mvarUsuarioA
        End Get
        Set(ByVal value As String)
            mvarUsuarioA = value
        End Set
    End Property


    Dim vMsgBox

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    '###########################

    <System.ComponentModel.Description("Método que engloba os outros sub-métodos para validar um usuário.")>
    Public Function validaUsuario(pUsuario As String, pSenha As String, Optional pNovaSenha As String = "", Optional pConfNovaSenha As String = "") As Integer
        'Attribute validaUsuario.VB_Description = "Método que engloba os outros sub-métodos para validar um usuário."
        'Define as variáveis:

        Dim vPalavra As String      'Dados agrupados descriptografados - primeira parte
        Dim vPalavra1 As String     'Dados agrupados descriptografados - segunda parte
        Dim vPalavra2 As String     'Dados agrupados descriptografados - terceira parte
        Dim vPalavra3 As String     'Dados agrupados descriptografados - quarta parte
        Dim vNomeComputador As String



        On Error GoTo EvalidaUsuario

        'Busca o nome do computador onde está rodando o sistema.
        'Necessário para gravação de LOG de problemas:
        'vNomeComputador = cdSeguranca1.sGetComputerName()
        vNomeComputador = New String(" "c, 32)
        vNomeComputador = Environ("computername")

        validaUsuario = dbUtils.dbConectaUsuario(0, 0)
        Db = dbUtils.Db
        RC = dbUtils.RC

        'Conecta ao DB e abre as tabelas:
        'validaUsuario = dbConectaUsuario(0, 0)
        If validaUsuario <> 0 Then Exit Function

        'Faz consistências iniciais dos ítens digitados:
        If pNovaSenha = pSenha Then
            validaUsuario = 1010
            Exit Function
        End If
        If pNovaSenha <> "" And pConfNovaSenha <> pNovaSenha Then
            If pConfNovaSenha = "" Then
                validaUsuario = 1011
            Else
                validaUsuario = 1012
            End If
            Exit Function
        End If

        'Verifica a existência do usuário:
        validaUsuario = localizaUsuario(pUsuario, 0)
        If validaUsuario <> 0 Then Exit Function

        'Descriptografa a primeira parte de Usuário:
        vPalavra = DeCripto(tUsuario.Fields(1).Value, 5)
        If Len(vPalavra) < 10 Then      'Ocorreu problemas
            validaUsuario = Val(vPalavra)
            Exit Function
        End If
        'Descriptografa a segunda parte de Usuário:
        vPalavra1 = DeCripto(tUsuario.Fields(2).Value, 5)
        If Len(vPalavra1) < 10 Then      'Ocorreu problemas
            validaUsuario = Val(vPalavra1)
            Exit Function
        End If
        'Descriptografa a terceira parte de Usuário:
        vPalavra2 = DeCripto(tUsuario.Fields(3).Value, 5)
        If Len(vPalavra2) < 10 Then      'Ocorreu problemas
            validaUsuario = Val(vPalavra2)
            Exit Function
        End If
        'Descriptografa a quarta parte de Usuário:
        vPalavra3 = DeCripto(tUsuario.Fields(4).Value, 5)
        If Len(vPalavra3) < 10 Then      'Ocorreu problemas
            validaUsuario = Val(vPalavra3)
            Exit Function
        End If

        'Separa variáveis:
        validaUsuario = separaAtribUsuario(vPalavra)
        If validaUsuario <> 0 Then Exit Function

        'Concatena a segunda parte da propriedade proc (serviços autorizados):
        vPalavra = proc
        proc = New String("0"c, 560)
        proc = vPalavra & vPalavra1 & vPalavra2 & vPalavra3

        'Verifica se usuário informado é o mesmo usuário criptografado:
        validaUsuario = verificaUsuario(pUsuario, Usuario)
        If validaUsuario <> 0 Then
            Dim segLog As SegurancaD.segLog
            segLog = New SegurancaD.segLog
            vPalavra = "Segurança: VIOLAÇÃO - Usuário criptografado = " & Usuario
            RC = segLog.dbConecta(1, 0)
            RC = segLog.inclui(pUsuario, vNomeComputador, vPalavra)
            Exit Function
        End If

        'Trata o tipo de status do usuário:
        Select Case status
            Case 0      'Usuário autorizado a acessar o sistema
                validaUsuario = status0(pSenha, pNovaSenha)
                If validaUsuario <> 0 Then Exit Function
            Case 1      'Usuário bloqueado
                validaUsuario = status1(vNomeComputador, pUsuario, pSenha, pNovaSenha)
                If validaUsuario <> 0 Then Exit Function
            Case 2      'Novo usuário - precisa alterar a senha
                validaUsuario = status2(vNomeComputador, pUsuario, pSenha, pNovaSenha)
                If validaUsuario <> 0 Then Exit Function
            Case 3      'Usuário já criado e não validado pelo administrador
                validaUsuario = 1009  'Msg para contactar o administrador do sistema
                Exit Function
            Case 4      'Usuário já criado e suspenso pelo administrador
                validaUsuario = 1008  'Msg de Usuário Bloqueado - Acesso Negado
                Exit Function
            Case Else
                validaUsuario = 1025
                Exit Function
        End Select

        'Grava arquivo com dados de segurança:
        geraDadosParaLog(pUsuario, vNomeComputador, senha)

        'Se não ocorreu nenhum erro, retorna 0:
        validaUsuario = 0

EvalidaUsuario:
        If Err.Number Then
            If Err.Number = 6 Then
                Return "Erro ao Acessar o Banco de Dados, Contacte o Administrador do Sistema."
            End If
            validaUsuario = Err.Number
            Exit Function
        End If

    End Function

    Public Function separaAtribUsuario(vPalavra As String) As Integer

        'Define variáveis:
        Dim vData As String

        On Error GoTo Esepara

        Usuario = Mid(vPalavra, 1, 25)
        senha = Mid(vPalavra, 26, 25)
        senhaAnt = Mid(vPalavra, 51, 25)
        senha2Ant = Mid(vPalavra, 76, 25)
        senha3Ant = Mid(vPalavra, 101, 25)
        vData = Mid(vPalavra, 126, 10)
        If IsDate(vData) Then
            dtSenhaAtu = vData
        End If
        status = Val(Mid(vPalavra, 136, 1))
        vData = Mid(vPalavra, 137, 10)
        If IsDate(vData) Then
            dtStatus = vData
        End If
        If IsNumeric(Mid(vPalavra, 147, 3)) Then
            freqAltSenha = Val(Mid(vPalavra, 147, 3))
        Else
            freqAltSenha = 0
        End If
        nome = Mid(vPalavra, 150, 50)
        Select Case Mid(vPalavra, 200, 1)
            Case "0"
                tpUsuario = 0   'Usuário Normal - tem acesso a todos os ambientes
            Case "1"
                tpUsuario = 1   'Administrador - tem acesso a todos os ambientes
            Case "2"
                tpUsuario = 2   'Produção - só tem acesso ao ambiente de produção
            Case "3"
                tpUsuario = 3   'Desenv./Teste - só tem acesso ao ambiente de desenvolvimento/testes
            Case Else
                tpUsuario = 3   'Desenv./Teste - só tem acesso ao ambiente de desenvolvimento/testes
        End Select
        proc = Mid(vPalavra, 201, 20)

        separaAtribUsuario = 0

Esepara:
        If Err.Number Then separaAtribUsuario = Err.Number

    End Function

    Public Function agrupaAtribUsuario() As String

        'Define variáveis de trabalho:
        Dim vUsuario As String
        Dim vSenha As String
        Dim vsenhaAnt As String
        Dim vsenha2Ant As String
        Dim vsenha3Ant As String
        Dim vdtSenhaAtu As String
        Dim vstatus As String
        Dim vdtStatus As String
        Dim vfreqAltSenha As String
        Dim vnome As String
        Dim vTpUsuario As String
        Dim vproc As String
        Dim vDescricao As String

        On Error GoTo Eagrupa

        'Move atributos para as variáveis:
        vUsuario = Usuario & IIf(Len(Usuario) < 25, New String(" "c, 25 - Len(Usuario)), "")
        vSenha = senha & IIf(Len(senha) < 25, New String(" "c, 25 - Len(senha)), "")
        vsenhaAnt = senhaAnt & IIf(Len(senhaAnt) < 25, New String(" "c, 25 - Len(senhaAnt)), "")
        vsenha2Ant = senha2Ant & IIf(Len(senha2Ant) < 25, New String(" "c, 25 - Len(senha2Ant)), "")
        vsenha3Ant = senha2Ant & IIf(Len(senha3Ant) < 25, New String(" "c, 25 - Len(senha3Ant)), "")
        vdtSenhaAtu = dtSenhaAtu
        vstatus = status
        vdtStatus = dtStatus
        vfreqAltSenha = Val(freqAltSenha).ToString("000")
        If Len(nome) < 50 Then
            vnome = nome & New String(" "c, 50 - Len(nome))
        Else
            vnome = Mid(nome, 1, 50)
        End If

        If tpUsuario = 1 Then
            vTpUsuario = "1"
        Else
            If tpUsuario = 0 Then
                vTpUsuario = "0"
            Else
                If tpUsuario = 2 Then
                    vTpUsuario = "2"
                Else
                    vTpUsuario = "3"
                End If
            End If
        End If
        If Len(proc) <= 20 Then
            If Len(proc) < 20 Then
                vproc = proc & New String(" "c, 20 - Len(proc))
            Else
                vproc = proc
            End If
        Else
            vproc = Mid(proc, 1, 20)
        End If

        'Agrupa os atributos:
        vDescricao = vUsuario & vSenha & vsenhaAnt & vsenha2Ant & vsenha3Ant & vdtSenhaAtu & vstatus
        vDescricao = vDescricao & vdtStatus & vfreqAltSenha & vnome & vTpUsuario & vproc
        If Len(vDescricao) < 220 Then
            vDescricao = vDescricao & New String(" "c, 220 - Len(vDescricao))
        Else
            vDescricao = vDescricao
        End If

        agrupaAtribUsuario = vDescricao

Eagrupa:
        If Err.Number Then agrupaAtribUsuario = Str(Err.Number)

    End Function

    Public Function verificaUsuario(pUsuarioA As String, pUsuario As String) As Integer

        If RTrim(UCase(pUsuarioA)) <> RTrim(UCase(pUsuario)) Then
            verificaUsuario = 1003
        Else
            verificaUsuario = 0
        End If

    End Function

    Public Function verificaSenha(pSenhaA As String, pSenha As String) As Integer
        'Se novo usuário, transforma senha informada em maiúscula
        'Senão, compara as senhas sem transformá-la em maiúscula
        'Verifica se a senha atende aos parâmetros de uma senha válida

        'Elimina brancos à direita:
        pSenhaA = RTrim(pSenhaA)
        pSenha = RTrim(pSenha)
        'Compara as senhas:
        If (pSenhaA <> pSenha And status <> 2) Or (UCase(pSenhaA) <> UCase(pSenha) And status = 2) Then
            verificaSenha = 1004
        Else
            verificaSenha = 0
        End If

    End Function

    Public Function alteraSenha(pNovaSenha As String) As Integer

        Dim FuncoesGerais As New SegurancaD.FuncoesGerais
        Dim vNovaSenha As String
        Dim I As Integer
        Dim J As Integer
        Dim vCaracter As String
        Dim vTamanho As Integer
        Dim vSelec As String

        On Error GoTo EalteraSenha
        '* *************************************************************
        '* Verifica se a nova senha atende aos critérios de segurança: *
        '* *************************************************************
        vSelec = "A senha tem que ter pelo menos 8 caracteres, " & Chr(10)
        vSelec = vSelec & "Ser diferente das últimas 3 senhas e:" & Chr(10)
        vSelec = vSelec & ". Pelo menos 1 letra maiúscula" & Chr(10)
        vSelec = vSelec & ". Pelo menos 1 letra minúscula" & Chr(10)
        vSelec = vSelec & ". Pelo menos 1 dígito numérico" & Chr(10)
        vSelec = vSelec & ". Pelo menos 1 caracter especial válido"
        vNovaSenha = RTrim(pNovaSenha)
        vTamanho = Len(vNovaSenha)
        'Testa letra maiúscula:
        For I = 1 To vTamanho
            vCaracter = Mid(vNovaSenha, I, 1)
            If Asc(vCaracter) >= 65 And Asc(vCaracter) <= 90 Then
                J = 1
                Exit For
            End If
        Next I
        If J = 0 Then
            MsgBox(vSelec)
            alteraSenha = 1007
            Exit Function
        End If
        'Testa letra mainúscula:
        For I = 1 To vTamanho
            vCaracter = Mid(vNovaSenha, I, 1)
            If Asc(vCaracter) >= 97 And Asc(vCaracter) <= 122 Then
                J = J + 2
                Exit For
            End If
        Next I
        If J <> 3 Then
            MsgBox(vSelec)
            alteraSenha = 1007
            Exit Function
        End If
        'Testa número:
        For I = 1 To vTamanho
            vCaracter = Mid(vNovaSenha, I, 1)
            If Asc(vCaracter) >= 48 And Asc(vCaracter) <= 57 Then
                J = J + 3
                Exit For
            End If
        Next I
        If J <> 6 Then
            MsgBox(vSelec)
            alteraSenha = 1007
            Exit Function
        End If
        'Testa caracter especial:
        For I = 1 To vTamanho
            vCaracter = Mid(vNovaSenha, I, 1)
            If Asc(vCaracter) >= 33 And Asc(vCaracter) <= 47 Then
                J = J + 4
                Exit For
            End If
            If Asc(vCaracter) >= 58 And Asc(vCaracter) <= 64 Then
                J = J + 4
                Exit For
            End If
        Next I
        If J <> 10 Then
            MsgBox(vSelec)
            alteraSenha = 1007
            Exit Function
        End If
        '* ***************************************
        '* Move senha atual para senha anterior: *
        '* ***************************************
        senha3Ant = senha2Ant
        senha2Ant = senhaAnt
        senhaAnt = senha
        'Move data atual para data da última alteração de senha:
        dtSenhaAtu = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortDate)
        'Move Nova Senha para Senha Atual:
        senha = pNovaSenha

        alteraSenha = 0

EalteraSenha:
        If Err.Number Then alteraSenha = Err.Number

    End Function

    Public Function verificaNovaSenha(pSenhaAnt As String, pNovaSenha As String) As Integer
        'Verifica se a Nova Senha náo é igual à senha anterior

        On Error GoTo EverifNSenha
        'Inválido se Nova Senha for igual à Senha Anterior:
        If RTrim(pNovaSenha) = RTrim(pSenhaAnt) Then
            verificaNovaSenha = 1007  'Msg de Nova Senha Inválida
            Exit Function
        End If

        verificaNovaSenha = 0

EverifNSenha:
        If Err.Number Then verificaNovaSenha = Err.Number

    End Function

    Public Function verifica2NovaSenha(pSenha2Ant As String, pNovaSenha As String) As Integer
        'Verifica se a Nova Senha náo é igual à segunda senha anterior

        On Error GoTo Everif2NSenha
        'Inválido se Nova Senha for igual à Senha Anterior:
        If RTrim(pNovaSenha) = RTrim(pSenha2Ant) Then
            verifica2NovaSenha = 1007  'Msg de Nova Senha Inválida
            Exit Function
        End If

        verifica2NovaSenha = 0

Everif2NSenha:
        If Err.Number Then verifica2NovaSenha = Err.Number

    End Function

    Public Function verifica3NovaSenha(pSenha3Ant As String, pNovaSenha As String) As Integer
        'Verifica se a Nova Senha náo é igual à terceira senha anterior

        On Error GoTo Everif3NSenha
        'Inválido se Nova Senha for igual à Senha Anterior:
        If RTrim(pNovaSenha) = RTrim(pSenha3Ant) Then
            verifica3NovaSenha = 1007  'Msg de Nova Senha Inválida
            Exit Function
        End If

        verifica3NovaSenha = 0

Everif3NSenha:
        If Err.Number Then verifica3NovaSenha = Err.Number

    End Function

    Private Function status0(pSenha As String, Optional pNovaSenha As String = "") As Integer
        '* **********************************************************
        '* * Efetua a validação, quando o status do usuário for 0   *
        '* **********************************************************

        'Define as variáveis:
        Dim vPalavra As String      'Dados agrupados descriptografados - 1a. parte
        Dim vPalavraC As String     'Dados agrupados criptografados - 1a. parte
        Dim vPalavra1 As String     'Dados agrupados descriptografados - 2a. parte
        Dim vPalavraC1 As String    'Dados agrupados criptografados - 2a. parte
        Dim vPalavra2 As String     'Dados agrupados descriptografados - 3a. parte
        Dim vPalavraC2 As String    'Dados agrupados criptografados - 3a. parte
        Dim vPalavra3 As String     'Dados agrupados descriptografados - 4a. parte
        Dim vPalavraC3 As String    'Dados agrupados criptografados - 4a. parte

        'Verifica se senha informada está correta:
        status0 = verificaSenha(pSenha, senha)
        If status0 <> 0 Then Exit Function
        'Se nova senha foi informada:
        If Len(pNovaSenha) >= 8 Then    'Informou a nova senha
            'Inválido se Nova Senha for igual à Senha Anterior:
            status0 = verificaNovaSenha(senhaAnt, pNovaSenha)
            If status0 <> 0 Then Exit Function
            status0 = verifica2NovaSenha(senha2Ant, pNovaSenha)
            If status0 <> 0 Then Exit Function
            status0 = verifica3NovaSenha(senha3Ant, pNovaSenha)
            If status0 <> 0 Then Exit Function
            'Altera Nova Senha como Senha Atual:
            status0 = alteraSenha(pNovaSenha)   'Altera senha
            If status0 <> 0 Then Exit Function
            vPalavra1 = Mid(proc, 21, 180)      'Captura a segunda parte de proc
            vPalavra2 = Mid(proc, 201, 180)     'Captura a terceira parte de proc
            vPalavra3 = Mid(proc, 381, 180)     'Captura a quarta parte de proc
            vPalavra = agrupaAtribUsuario() 'Monta dados para serem criptografados
            If Len(vPalavra) = 220 Then      'Agrupou dados sem problemas
                'Criptografa informações do usuário:
                vPalavraC = cripto(vPalavra, 5)     '1a. parte
                If Len(vPalavraC) <> Len(vPalavra) Then  'Ocorreu erro
                    status0 = Val(vPalavraC)
                    Exit Function
                End If
                vPalavraC1 = cripto(vPalavra1, 5)   '2a. parte
                If Len(vPalavraC1) <> 180 Then  'Ocorreu erro
                    status0 = Val(vPalavraC1)
                    Exit Function
                End If
                vPalavraC2 = cripto(vPalavra2, 5)   '3a. parte
                If Len(vPalavraC2) <> 180 Then  'Ocorreu erro
                    status0 = Val(vPalavraC2)
                    Exit Function
                End If
                vPalavraC3 = cripto(vPalavra3, 5)   '4a. parte
                If Len(vPalavraC3) <> 180 Then  'Ocorreu erro
                    status0 = Val(vPalavraC3)
                    Exit Function
                End If
                'Grava alterações: nova senha e data da alteração da senha:
                status0 = alteraUsuario(vPalavraC, vPalavraC1, vPalavraC2, vPalavraC3)
                If status0 <> 0 Then Exit Function
                'Exibe mensagem de Senha Alterada com Sucesso!, e continua o processamento:
                status0 = segmensagens.dbConecta(0, 0)
                segmensagens.exibeMsg(64, 1006)
            Else        'Ocorreu erro durante o agrupamento dos dados
                status0 = Val(vPalavra)
                Exit Function
            End If
        Else                    'Não informou a nova senha
            'Calcula data de expiração da senha:
            If freqAltSenha > 0 Then    'Num. de dias corridos para alterar a senha
                'Necessário trocar a senha:
                If Val(freqAltSenha) < DateDiff("d", dtSenhaAtu, FuncoesGerais.dataServ()) Then
                    status0 = 1005  'Msg para informar nova senha
                    Exit Function
                End If
            End If
        End If
        status0 = 0

    End Function

    Private Function status1(vNomeComputador As String, pUsuario As String, pSenha As String, Optional pNovaSenha As String = "") As Integer
        '* **********************************************************
        '* * Efetua a validação, quando o status do usuário for 1   *
        '* **********************************************************

        'Define as variáveis:
        Dim vPalavra As String      'Dados agrupados descriptografados - 1a. parte
        Dim vPalavraC As String     'Dados agrupados criptografados - 1a. parte
        Dim vPalavra1 As String      'Dados agrupados descriptografados - 2a. parte
        Dim vPalavraC1 As String     'Dados agrupados criptografados - 2a. parte
        Dim vPalavra2 As String      'Dados agrupados descriptografados - 3a. parte
        Dim vPalavraC2 As String     'Dados agrupados criptografados - 3a. parte
        Dim vPalavra3 As String      'Dados agrupados descriptografados - 4a. parte
        Dim vPalavraC3 As String     'Dados agrupados criptografados - 4a. parte
        Dim RC As Integer

        'Calcula data de bloqueio do usuário:
        If DateDiff("d", dtStatus, FuncoesGerais.dataServ()) >= 1 Then 'Pelo menos 1 dia bloqueado
            'Verifica se senha informada está correta:
            status1 = verificaSenha(pSenha, senha)
            If status1 <> 0 Then Exit Function
            If Len(pNovaSenha) >= 6 Then    'Informou a nova senha
                'Inválido se Nova Senha for igual à Senha Anterior:
                status1 = verificaNovaSenha(senha, pNovaSenha)
                If status1 <> 0 Then Exit Function
                status1 = verifica2NovaSenha(senha2Ant, pNovaSenha)
                If status1 <> 0 Then Exit Function
                status1 = verifica3NovaSenha(senha3Ant, pNovaSenha)
                If status1 <> 0 Then Exit Function
                'Altera Nova Senha como Senha Atual:
                status1 = alteraSenha(pNovaSenha)     'Altera senha
                If status1 <> 0 Then Exit Function
                'Usuário autorizado a acessar o sistema:
                status = 0
                'Data da alteração do status = Hoje:
                dtStatus = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortDate)
                vPalavra1 = Mid(proc, 21, 180)  'Captura a segunda parte de proc
                vPalavra2 = Mid(proc, 201, 180) 'Captura a terceira parte de proc
                vPalavra3 = Mid(proc, 381, 180) 'Captura a quarta parte de proc
                vPalavra = agrupaAtribUsuario() 'Monta dados para serem criptografados
                If Len(vPalavra) > 10 Then      'Agrupou dados sem problemas
                    'Criptografa informações do usuário - 1a. parte
                    vPalavraC = cripto(vPalavra, 5)
                    If Len(vPalavraC) < 6 Then  'Ocorreu erro
                        status1 = Val(vPalavraC)
                        Exit Function
                    End If
                    'Criptografa informações do usuário - 2a. parte
                    vPalavraC1 = cripto(vPalavra1, 5)
                    If Len(vPalavraC1) < 6 Then  'Ocorreu erro
                        status1 = Val(vPalavraC1)
                        Exit Function
                    End If
                    'Criptografa informações do usuário - 3a. parte
                    vPalavraC2 = cripto(vPalavra2, 5)
                    If Len(vPalavraC2) < 6 Then  'Ocorreu erro
                        status1 = Val(vPalavraC2)
                        Exit Function
                    End If
                    'Criptografa informações do usuário - 4a. parte
                    vPalavraC3 = cripto(vPalavra3, 5)
                    If Len(vPalavraC3) < 6 Then  'Ocorreu erro
                        status1 = Val(vPalavraC3)
                        Exit Function
                    End If
                    'Grava alterações: status, data alt. status, nova senha e data alt. senha:
                    status1 = alteraUsuario(vPalavraC, vPalavraC1, vPalavraC2, vPalavraC3)
                    If status1 <> 0 Then Exit Function
                    'Grava LOG de mudança de status:
                    Dim segLog As SegurancaD.segLog
                    segLog = New SegurancaD.segLog
                    vPalavra = "Segurança: Status Anterior = 1 / Status Atual = 0"
                    RC = segLog.dbConecta(0, 0)
                    RC = segLog.inclui(pUsuario, vNomeComputador, vPalavra)
                    'Exibe mensagem de Senha Alterada com Sucesso!, e continua o processamento:
                    RC = segmensagens.dbConecta(0, 0)
                    segmensagens.exibeMsg(64, 1006)
                Else        'Ocorreu erro durante o agrupamento dos dados
                    status1 = Val(vPalavra)
                    Exit Function
                End If
            Else                    'Não informou a nova senha
                status1 = 1005  'Msg para informar nova senha
                Exit Function
            End If
        Else    'Usuário ainda tem que estar bloqueado
            status1 = 1008  'Msg de Usuário bloqueado. Acesso negado.
            Exit Function
        End If
        status1 = 0

    End Function

    Private Function status2(vNomeComputador As String, pUsuario As String, pSenha As String, Optional pNovaSenha As String = "") As Integer
        '* **********************************************************
        '* * Efetua a validação, quando o status do usuário for 2   *
        '* **********************************************************

        'Define as variáveis:
        Dim vPalavra As String      'Dados agrupados descriptografados - 1a. parte
        Dim vPalavraC As String     'Dados agrupados criptografados - 1a. parte
        Dim vPalavra1 As String     'Dados agrupados descriptografados - 2a. parte
        Dim vPalavraC1 As String    'Dados agrupados criptografados - 2a. parte
        Dim vPalavra2 As String     'Dados agrupados descriptografados - 3a. parte
        Dim vPalavraC2 As String    'Dados agrupados criptografados - 3a. parte
        Dim vPalavra3 As String     'Dados agrupados descriptografados - 4a. parte
        Dim vPalavraC3 As String    'Dados agrupados criptografados - 4a. parte
        Dim RC As Integer

        'Verifica se senha informada está correta.
        'Quando status = 2, colocar a senha informada em maiúscula:
        status2 = verificaSenha(UCase(pSenha), RTrim(UCase(senha)))
        If status2 <> 0 Then Exit Function
        'Se nova senha foi informada:
        If Len(pNovaSenha) >= 8 Then    'Informou a nova senha
            'Inválido se Nova Senha for igual à Senha Anterior:
            status2 = verificaNovaSenha(UCase(senha), pNovaSenha)
            If status2 <> 0 Then Exit Function
            status2 = verifica2NovaSenha(UCase(senha2Ant), pNovaSenha)
            If status2 <> 0 Then Exit Function
            status2 = verifica3NovaSenha(UCase(senha3Ant), pNovaSenha)
            If status2 <> 0 Then Exit Function
            'Altera Nova Senha como Senha Atual:
            status2 = alteraSenha(pNovaSenha)     'Altera senha
            If status2 <> 0 Then Exit Function
            'Usuário autorizado a acessar o sistema:
            status = 0
            'Data da alteração do status = Hoje:
            dtStatus = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortDate)
            vPalavra1 = Mid(proc, 21, 180)  'Captura a segunda parte de proc
            vPalavra2 = Mid(proc, 201, 180) 'Captura a terceira parte de proc
            vPalavra3 = Mid(proc, 381, 180) 'Captura a quarta parte de proc
            vPalavra = agrupaAtribUsuario() 'Monta dados para serem criptografados
            If Len(vPalavra) > 10 Then      'Agrupou dados sem problemas
                'Criptografa informações do usuário - 1a. parte:
                vPalavraC = cripto(vPalavra, 5)
                If Len(vPalavraC) < 6 Then  'Ocorreu erro
                    status2 = Val(vPalavraC)
                    Exit Function
                End If
                'Criptografa informações do usuário - 2a. parte:
                vPalavraC1 = cripto(vPalavra1, 5)
                If Len(vPalavraC1) < 6 Then  'Ocorreu erro
                    status2 = Val(vPalavraC1)
                    Exit Function
                End If
                'Criptografa informações do usuário - 3a. parte:
                vPalavraC2 = cripto(vPalavra2, 5)
                If Len(vPalavraC2) < 6 Then  'Ocorreu erro
                    status2 = Val(vPalavraC2)
                    Exit Function
                End If
                'Criptografa informações do usuário - 4a. parte:
                vPalavraC3 = cripto(vPalavra3, 5)
                If Len(vPalavraC3) < 6 Then  'Ocorreu erro
                    status2 = Val(vPalavraC3)
                    Exit Function
                End If
                'Grava alterações: status, data alt. status, nova senha e data alt. senha:
                status2 = alteraUsuario(vPalavraC, vPalavraC1, vPalavraC2, vPalavraC3)
                If status2 <> 0 Then Exit Function
                'Grava LOG de mudança de status:
                Dim segLog As SegurancaD.segLog
                segLog = New SegurancaD.segLog
                vPalavra = "Segurança: Status Anterior = 2 / Status Atual = 0"
                RC = segLog.dbConecta(0, 0)
                RC = segLog.inclui(pUsuario, vNomeComputador, vPalavra)
                'Exibe mensagem de Senha Alterada com Sucesso!, e continua o processamento:
                RC = segmensagens.dbConecta(0, 0)
                segmensagens.exibeMsg(64, 1006)
            Else        'Ocorreu erro durante o agrupamento dos dados
                status2 = Val(vPalavra)
                Exit Function
            End If
        Else                    'Não informou a nova senha
            status2 = 1005  'Msg para informar nova senha
            Exit Function
        End If
        status2 = 0

    End Function

    Public Function trataSeguranca(pNumTentativas As Integer, pUsuario As String, pSenha As String, Optional pNovaSenha As String = "", Optional pConfNovaSenha As String = "") As Integer
        '* ******************************************************
        '* * Efetua todo o processamento da Camada de Segurança *
        '* ******************************************************

        trataSeguranca = 0
        'Ativa classes:
        cdSeguranca1 = New SegurancaD.cdSeguranca1

        'Faz toda a validação do usuário:
        trataSeguranca = validaUsuario(pUsuario, pSenha, pNovaSenha, pConfNovaSenha)
        If trataSeguranca = 6 Then Exit Function

        'Faz o tratamento do Código de Retorno:
        If trataSeguranca <> 0 Then
            trataSeguranca = trataRetorno(pNumTentativas, trataSeguranca, pUsuario, pSenha, pNovaSenha)
            Exit Function
        End If

    End Function

    Private Function trataRetorno(pNumTentativas As Integer, pTrataSeguranca As Integer, pUsuario As String, pSenha As String, pPath As String, Optional pNovaSenha As String = "") As Integer

        'Define as variáveis:
        Dim vPalavra As String      'Dados agrupados descriptografados - 1a. parte
        Dim vPalavraC As String     'Dados agrupados criptografados - 1a. parte
        Dim vPalavra1 As String     'Dados agrupados descriptografados - 2a. parte
        Dim vPalavraC1 As String    'Dados agrupados criptografados - 2a. parte
        Dim vPalavra2 As String     'Dados agrupados descriptografados - 3a. parte
        Dim vPalavraC2 As String    'Dados agrupados criptografados - 3a. parte
        Dim vPalavra3 As String     'Dados agrupados descriptografados - 4a. parte
        Dim vPalavraC3 As String    'Dados agrupados criptografados - 4a. parte
        Dim vNomeComputador As String
        Dim RC As Integer
        Dim segmensagens As New SegurancaD.segMensagens

        RC = segmensagens.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro ao abrir a tabela de mensagens")
        End If
        'O DB não foi encontrado. Não adianta continuar:
        If pTrataSeguranca = 1046 Then
            MsgBox("O Banco de Dados Não Foi Encontrado. Contate o Administrador.", vbCritical, "Erro!")
            trataRetorno = pTrataSeguranca
            Exit Function
        End If

        'Ativa classe de mensagens:
        segmensagens = New SegurancaD.segMensagens

        'Abre a tabela de Mensagens:
        RC = segmensagens.dbConecta(0, 0)

        'Este é o código que deverá permanecer:
        trataRetorno = pTrataSeguranca
        'Analisa o erro que ocorreu:
        Select Case pTrataSeguranca
            Case 1001, 1002, 1003, 1007, 1008, 1009, 1010, 1011, 1012, 1025
                segmensagens.exibeMsg(16, pTrataSeguranca)
            Case 1004     'Senha inválida
                If pNumTentativas >= 2 Then
                    trataRetorno = 1008     'Usuário Bloqueado
                    'Busca o nome do computador onde está rodando o sistema.
                    'Necessário para gravação de LOG de problemas:
                    vNomeComputador = cdSeguranca1.sGetComputerName()
                    'Gera Log
                    Dim segLog As SegurancaD.segLog
                    segLog = New SegurancaD.segLog
                    vPalavra = "Segurança: Três tentativas de senha inválida"
                    RC = segLog.dbConecta(0, 0)
                    RC = segLog.inclui(pUsuario, vNomeComputador, vPalavra)
                    'Usuário bloqueado:
                    status = 1
                    'Data da alteração do status = Hoje:
                    dtStatus = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortDate)
                    vPalavra1 = Mid(proc, 21, 180)  'Extrai a 2a. parte de proc
                    vPalavra2 = Mid(proc, 281, 180) 'Extrai a 3a. parte de proc
                    vPalavra3 = Mid(proc, 381, 180) 'Extrai a 4a. parte de proc
                    vPalavra = agrupaAtribUsuario() 'Monta dados para serem criptografados
                    If Len(vPalavra) > 10 Then      'Agrupou dados sem problemas
                        'Criptografa informações do usuário - 1a. parte:
                        vPalavraC = cripto(vPalavra, 5)
                        If Len(vPalavraC) < 6 Then  'Ocorreu erro
                            segmensagens.exibeMsg(, Val(vPalavraC))
                        End If
                        'Criptografa informações do usuário - 2a. parte:
                        vPalavraC1 = cripto(vPalavra1, 5)
                        If Len(vPalavraC1) < 6 Then  'Ocorreu erro
                            segmensagens.exibeMsg(, Val(vPalavraC1))
                        End If
                        'Criptografa informações do usuário - 3a. parte:
                        vPalavraC2 = cripto(vPalavra2, 5)
                        If Len(vPalavraC2) < 6 Then  'Ocorreu erro
                            segmensagens.exibeMsg(, Val(vPalavraC2))
                        End If
                        'Criptografa informações do usuário - 4a. parte:
                        vPalavraC3 = cripto(vPalavra3, 5)
                        If Len(vPalavraC3) < 6 Then  'Ocorreu erro
                            segmensagens.exibeMsg(, Val(vPalavraC3))
                        End If
                        'Grava alterações: status e data alt. status:
                        RC = alteraUsuario(vPalavraC, vPalavraC1, vPalavraC2, vPalavraC3)
                        If RC <> 0 Then segmensagens.exibeMsg(, RC)  'Ocorreu erro
                        'Exibe mensagem de Usuário Bloqueado:
                        segmensagens.exibeMsg(16, 1008)
                    End If
                Else    'Usuário pode tentar novamente
                    'Exibe mensagem Senha Inválida. Favor informar a senha correta:
                    segmensagens.exibeMsg(16, 1004)
                End If
            Case 1005
                'Exibe mensagem Necessário Informar Nova Senha:
                segmensagens.exibeMsg(48, 1005)
            Case Else
                'Exibe a mensagem com problema:
                segmensagens.exibeMsg(, pTrataSeguranca)
        End Select

    End Function

    Public Function DeCripto(cPalavra As String, nTaxa As Integer) As String
        '* *****************************************************
        '* * DesCriptografa cPalavra segundo o algorítmo nTaxa *
        '* *****************************************************

        Dim nPosicao As Integer
        Dim cNovaPal As String = ""
        Dim cPalavr As String
        Dim I As Integer

        On Error GoTo EDeCripto
        If nTaxa = 0 Then
            DeCripto = 1002   'Taxa de Descriptografia não informada
            Exit Function
        End If

        I = 1
        Do While I <= Len(cPalavra)
            cPalavr = Mid(cPalavra, I, 10)
            For nPosicao = 1 To Len(cPalavr)
                cNovaPal = cNovaPal + Chr(Asc(Mid(cPalavr, nPosicao, 1)) - nPosicao - nTaxa)
            Next nPosicao
            I = I + 10
        Loop

        DeCripto = cNovaPal
EDeCripto:
        If Err.Number Then
            MsgBox(Err.Number & " - " & Err.Description)
            DeCripto = ""
        End If

    End Function

    Public Function cripto(cPalavra As String, nTaxa As Integer) As String
        '* **************************************************
        '* * Criptografa cPalavra segundo o algorítmo nTaxa *
        '* **************************************************

        Dim nPosicao As Integer
        Dim nInd As Integer
        Dim nI As Integer
        Dim I As Integer
        Dim cNovaPal As String = ""
        Dim vPalavr As String

        On Error GoTo Ecripto

        If nTaxa = 0 Then
            For nI = 1 To Len(cPalavra)
                nInd = Asc(Mid(cPalavra, nI, 1)) + nInd
            Next nI
            nInd = Int(nInd / Len(cPalavra))
        Else
            nInd = nTaxa
        End If

        I = 1
        Do While I <= Len(cPalavra)
            vPalavr = Mid(cPalavra, I, 10)
            For nPosicao = 1 To Len(vPalavr)
                cNovaPal = cNovaPal + Chr(Asc(Mid(vPalavr, nPosicao, 1)) + nPosicao + nInd)
            Next nPosicao
            I = I + 10
        Loop

        cripto = cNovaPal

        '* *******************************
        '* * Executa quando houver erro: *
        '* *******************************
Ecripto:
        If Err.Number Then cripto = Str(Err.Number)

    End Function

    Private Function DeCripto_Nova(cPalavra As String, nTaxa As Integer) As String
        '* **************************************************************
        '* * DesCriptografa cPalavra segundo o algorítmo nTaxa          *
        '* * ////////////////////////////////////////////////////////// *
        '* * ==> Nova Function, corrigindo erros da versão anterior <== *
        '* **************************************************************

        Dim nPosicao As Integer
        Dim cPalavr As String
        Dim I As Integer
        Dim vValor As Integer

        DeCripto_Nova = ""
        If nTaxa <= 0 Or nTaxa > 64 Then
            DeCripto_Nova = 1002    'Taxa de Descriptografia não informada
            Exit Function
        End If
        If cPalavra = "" Or cPalavra = " " Then
            DeCripto_Nova = 1057    'Descrição Tem Que Ser Preenchida
            Exit Function
        End If

        I = 1
        Do While I <= Len(cPalavra)
            cPalavr = Mid(cPalavra, I, nTaxa)
            For nPosicao = 1 To Len(cPalavr)
                vValor = Asc(Mid(cPalavr, nPosicao, 1)) - nPosicao - nTaxa
                If vValor >= 0 Then
                    DeCripto_Nova = DeCripto_Nova + Chr(vValor)
                Else
                    DeCripto_Nova = DeCripto_Nova + Chr(vValor + 255)
                End If
            Next nPosicao
            I = I + nTaxa
        Loop

    End Function

    Private Function cripto_Nova(cPalavra As String, nTaxa As Integer) As String
        '* **************************************************************
        '* * criptografa cPalavra segundo o algorítmo nTaxa             *
        '* * ////////////////////////////////////////////////////////// *
        '* * ==> Nova Function, corrigindo erros da versão anterior <== *
        '* **************************************************************

        Dim nPosicao As Integer
        Dim nInd As Integer
        Dim nI As Integer
        Dim I As Integer
        Dim vPalavr As String
        Dim vValor As Integer

        On Error GoTo Ecripto_Nova
        cripto_Nova = ""
        If nTaxa <= 0 Or nTaxa > 64 Then
            cripto_Nova = 1002       'Taxa de criptografia não informada
            Exit Function
        End If
        If cPalavra = "" Or cPalavra = " " Then
            cripto_Nova = 1057    'Descrição Tem Que Ser Preenchida
            Exit Function
        End If

        I = 1
        Do While I <= Len(cPalavra)
            vPalavr = Mid(cPalavra, I, nTaxa)
            For nPosicao = 1 To Len(vPalavr)
                vValor = Asc(Mid(vPalavr, nPosicao, 1)) + nPosicao + nTaxa
                If vValor <= 255 Then
                    cripto_Nova = cripto_Nova + Chr(vValor)
                Else
                    cripto_Nova = cripto_Nova + Chr(vValor - 255)
                End If
            Next nPosicao
            I = I + nTaxa
        Loop

        '* *******************************
        '* * Executa quando houver erro: *
        '* *******************************
Ecripto_Nova:
        If Err.Number Then cripto_Nova = Str(Err.Number)

    End Function

    Public Function trataUsuariosSeq(Palavra As String, Palavra1 As String, Palavra2 As String, Palavra3 As String) As Integer
        'Lê uma linha da tabela de usuários,
        'e faz todo o tratamento de descriptografia

        'Define variáveis
        Dim RC As Integer
        Dim vPalavra As String
        Dim vPalavra1 As String
        Dim vPalavra2 As String
        Dim vPalavra3 As String

        'Descriptografa dados - 1a. parte:
        vPalavra = DeCripto(Palavra, 5)
        If Len(vPalavra) < 10 Then
            trataUsuariosSeq = Val(vPalavra)
            'Exibe mensagem:
            RC = segmensagens.dbConecta(0, 0)
            segmensagens.exibeMsg(, Val(vPalavra))
            Exit Function
        End If
        'Descriptografa dados - 2a. parte:
        vPalavra1 = DeCripto(Palavra1, 5)
        If Len(vPalavra1) < 10 Then
            trataUsuariosSeq = Val(vPalavra1)
            'Exibe mensagem:
            RC = segmensagens.dbConecta(0, 0)
            segmensagens.exibeMsg(, Val(vPalavra1))
            Exit Function
        End If
        'Descriptografa dados - 3a. parte:
        vPalavra2 = DeCripto(Palavra2, 5)
        If Len(vPalavra2) < 10 Then
            trataUsuariosSeq = Val(vPalavra2)
            'Exibe mensagem:
            RC = segmensagens.dbConecta(0, 0)
            segmensagens.exibeMsg(, Val(vPalavra2))
            Exit Function
        End If
        'Descriptografa dados - 4a. parte:
        vPalavra3 = DeCripto(Palavra3, 5)
        If Len(vPalavra3) < 10 Then
            trataUsuariosSeq = Val(vPalavra3)
            'Exibe mensagem:
            RC = segmensagens.dbConecta(0, 0)
            segmensagens.exibeMsg(, Val(vPalavra3))
            Exit Function
        End If

        'Separa os dados nas respectivas propriedades de Usuário:
        RC = separaAtribUsuario(vPalavra)
        If RC <> 0 Then
            trataUsuariosSeq = RC
            'Exibe mensagem:
            RC = segmensagens.dbConecta(0, 0)
            segmensagens.exibeMsg(, RC)
            Exit Function
        End If
        proc = proc & vPalavra1 & vPalavra2 & vPalavra3

        trataUsuariosSeq = 0

    End Function

    Public Function preparaGravUsuario() As String
        'Efetua o agrupamento de dados do usuário, e os criptografa

        'Agrupa os valores para criptografar:
        preparaGravUsuario = agrupaAtribUsuario()
        If Len(preparaGravUsuario) < 10 Then Exit Function

        'Criptografa os dados:
        preparaGravUsuario = cripto(preparaGravUsuario, 5)

    End Function

    Public Function preparaGravUsuario2() As String
        'Efetua a criptografia da 2a. parte de proc

        preparaGravUsuario2 = cripto(Mid(proc, 21, 180), 5)

    End Function

    Public Function preparaGravUsuario3() As String
        'Efetua a criptografia da 3a. parte de proc

        preparaGravUsuario3 = cripto(Mid(proc, 201, 180), 5)

    End Function

    Public Function preparaGravUsuario4() As String
        'Efetua a criptografia da 4a. parte de proc

        preparaGravUsuario4 = cripto(Mid(proc, 381, 180), 5)

    End Function

    Public Function dbConectaUsuario(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset                      *
        '* ****************************************

        Dim strConnect As String, strSQL As String

        On Error GoTo EdbConectaUsuario

        dbConectaUsuario = 0    'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            '* ********************************************
            '* Cria uma instância de ADODB.Connection:    *
            '* ********************************************
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            'strConnect = cdSeguranca1.LeDADOSsys(1)
            strConnect = "Provider=SQLOLEDB;Data Source=192.168.1.119,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;"
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        'tUsuario = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            tUsuario.Open("Usuario", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Else
            If vSelec = "" Then
                dbConectaUsuario = 1014
            Else
                tUsuario.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConectaUsuario:
        If Err.Number Then
            If Err.Number = 3705 Then
                tUsuario.Close()
                tUsuario.Open("UsuarioBKP", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Resume Next
            Else
                dbConectaUsuario = Err.Number

                If Err.Number <> -2147467259 Then
                    If Err.Number = 3024 Then
                        dbConectaUsuario = 1046
                    Else
                        dbConectaUsuario = Err.Number
                    End If
                    MsgBox("Classe Usuario - dbConectaUsuario" & vbCrLf & Err.Number & " - " & Err.Description)
                End If
            End If
        End If

    End Function

    Public Function localizaUsuario(pUsuario As String, atualiz As Integer) As Integer
        'Busca um usuário específico (pUsuario) na tabela tUsuario:
        On Error GoTo ElocalizaUsuario
        Dim vSelect As String
        vSelect = "SELECT * FROM Usuario WHERE UsuarioA = '" & pUsuario & "'"

        RC.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic,)
        tUsuario = RC
        localizaUsuario = 0    'ReturnCode se não houver nenhum problema
        'Não encontrou:
        If tUsuario.EOF Then
            localizaUsuario = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            UsuarioA = tUsuario.Fields("UsuarioA").Value
            Descricao = tUsuario.Fields("Descricao").Value
            Descricao1 = tUsuario.Fields("Descricao1").Value
            Descricao2 = tUsuario.Fields("Descricao2").Value
            Descricao3 = tUsuario.Fields("Descricao3").Value
            indice = tUsuario.Fields("Indice").Value
        End If

ElocalizaUsuario:
        If Err.Number Then
            If Err.Number = 3705 Then
                tUsuario.Close()
                tUsuario.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
            Resume Next
        End If

    End Function

    Public Function alteraUsuario(Descricao As String, Descricao1 As String, Descricao2 As String, Descricao3 As String) As Integer

        On Error GoTo EalteraUsuario

        Dim vSelect As String

        Dim IdAtual As String = UsuarioA

        tUsuario.Fields(1).Value = Descricao
        tUsuario.Fields(2).Value = Descricao1
        tUsuario.Fields(3).Value = Descricao2
        tUsuario.Fields(4).Value = Descricao3
        tUsuario.Fields("indice").Value = 0
        tUsuario.Update()
        alteraUsuario = 0

EalteraUsuario:
        If Err.Number Then
            alteraUsuario = Err.Number

        End If

    End Function
    Public Function criaUsuario(Usuario As String, Descricao As String, Descricao1 As String, Descricao2 As String, Descricao3 As String) As Integer

        On Error GoTo EcriaUsuario
        tUsuario.AddNew()
        tUsuario.Fields(0).Value = Usuario
        tUsuario.Fields(1).Value = Descricao
        tUsuario.Fields(2).Value = Descricao1
        tUsuario.Fields(3).Value = Descricao2
        tUsuario.Fields(4).Value = Descricao3
        tUsuario.Fields("indice").Value = 0
        tUsuario.Update()
        criaUsuario = 0

EcriaUsuario:
        If Err.Number Then criaUsuario = Err.Number

    End Function

    Public Function leTabUsuarios() As Integer

        On Error GoTo EleTabUsuarios
        'Se não estiver no fim de arquivo, lê 1 linha:
        If Not tUsuario.EOF Then
            'Se encontrou, carrega propriedades:
            UsuarioA = tUsuario.Fields("UsuarioA").Value
            Descricao = tUsuario.Fields("Descricao").Value
            Descricao1 = tUsuario.Fields("Descricao1").Value
            Descricao2 = tUsuario.Fields("Descricao2").Value
            Descricao3 = tUsuario.Fields("Descricao3").Value
            indice = tUsuario.Fields("Indice").Value
            tUsuario.MoveNext()
        Else
            leTabUsuarios = 1016
        End If
        'Testa se já chegou no fim de arquivo
        If Not tUsuario.EOF Then
            leTabUsuarios = 0
        Else
            leTabUsuarios = 1016
        End If

EleTabUsuarios:
        If Err.Number Then leTabUsuarios = Err.Number

    End Function

    Public Function tabUsuariosComDados() As Integer
        'Verifica se a tabela de usuários possui dados:

        On Error GoTo EtabUsuariosComDados
        If Not tUsuario.EOF Then
            tUsuario.MoveLast()
            tUsuario.MoveFirst()
            tabUsuariosComDados = tUsuario.RecordCount
        Else
            tabUsuariosComDados = 0
        End If

EtabUsuariosComDados:
        If Err.Number Then
            MsgBox("Erro " & Err.Number & " = " & Err.Description)
            tabUsuariosComDados = -1
        End If

    End Function

    Public Sub geraDadosParaLog(pUsuario As String, pNomeComputador As String, pSenha As String)
        'Prepara dados para identificação do log do sistema a ser gravado:

        'Define variáveis:
        Dim RC As String
        Dim FNum As Integer
        Dim vCurDir As String = ""
        Dim vDados As String

        'Prepara os dados para criptografia:
        vDados = pUsuario & IIf(Len(pUsuario) < 25, New String(" "c, 25 - Len(pUsuario)), "")
        vDados = vDados & pNomeComputador & IIf(Len(pNomeComputador) < 20, New String(" "c, 20 - Len(pNomeComputador)), "")
        vDados = vDados & pSenha & IIf(Len(pSenha) < 25, New String(" "c, 25 - Len(pSenha)), "")
        'criptografa os dados:
        RC = cripto(vDados, 5)
        If Len(RC) <> Len(vDados) Then
            Exit Sub
        End If

        'Verifica se DADOS.SYS existe. Se sim, elimina-o:
        If cdSeguranca1.FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOS.SYS") Then
            Kill("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOS.SYS")
            vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
            FNum = 1
        Else
            If cdSeguranca1.FileExists("C:\Auditeste\DADOS.SYS") Then
                Kill("C:\Auditeste\DADOS.SYS")
                vCurDir = "C:\Auditeste"
                FNum = 2
            Else
                If cdSeguranca1.FileExists("C:\Arquivos de Programas\Auditeste\DADOS.SYS") Then
                    Kill("C:\Arquivos de Programas\Auditeste\DADOS.SYS")
                    vCurDir = "C:\Arquivos de Programas\Auditeste"
                    FNum = 3
                Else
                    If cdSeguranca1.FileExists("C:\Program Files\Auditeste\DADOS.SYS") Then
                        Kill("C:\Program Files\Auditeste\DADOS.SYS")
                        vCurDir = "C:\Program Files\Auditeste"
                        FNum = 4
                    Else
                        If cdSeguranca1.FileExists("C:\Windows\DADOS.SYS") Then
                            Kill("C:\Windows\DADOS.SYS")
                            vCurDir = "C:\Windows"
                            FNum = 5
                        End If
                    End If
                End If
            End If
        End If
        If FNum = 0 Then
            If cdSeguranca1.FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\AuditesteSQLe.exe") Then
                vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
                FNum = 1
            Else
                If cdSeguranca1.FileExists("C:\Auditeste\AuditesteSQLe.exe") Then
                    vCurDir = "C:\Auditeste"
                    FNum = 2
                Else
                    If cdSeguranca1.FileExists("C:\Arquivos de Programas\Auditeste\AuditesteSQLe.exe") Then
                        vCurDir = "C:\Arquivos de Programas\Auditeste"
                        FNum = 3
                    Else
                        If cdSeguranca1.FileExists("C:\Program Files\Auditeste\AuditesteSQLe.exe") Then
                            vCurDir = "C:\Program Files\Auditeste"
                            FNum = 4
                        Else
                            If cdSeguranca1.FileExists("C:\Windows\DADOS.SYS") Then
                                vCurDir = "C:\Windows"
                                FNum = 5
                            End If
                        End If
                    End If
                End If
            End If
        End If
        '* ***********************************************************
        '* ***********************************************************
        '* Se for compilar para instalação na VM, tirar o comentário *
        '* Caso contrário, deixar como comentário antes de compilar: *
        '* ***********************************************************
        '* Início
        '* ***********************************************************
        ''Verifica se DADOS.SYS existe. Se sim, elimina-o:
        'vCurDir = System.AppDomain.CurrentDomain.BaseDirectory()
        'FNum = cdSeguranca1.FileExists(vCurDir & "\DADOS.SYS")
        'If FNum <> 0 Then   'Arquivo existe
        '    Kill(vCurDir & "\DADOS.SYS")
        'End If
        '* ***********************************************************
        '* Término
        '* ***********************************************************
        'Grava um arquivo com os dados criptografados:
        Dim RCOut() As Byte = System.Text.Encoding.Default.GetBytes(RC)
        FileSystem.WriteAllBytes(vCurDir & "\DADOS.SYS", RCOut, False)
        '    FNum = FreeFile()
        '    Open vCurDir & "\DADOS.SYS" For Binary Access Write As #FNum    'Grava o arquivo
        '    Put #FNum, 1, RC
        '    Close #FNum

    End Sub

    Public Function leDadosParaLog() As String
        'Le dados para identificação do log do sistema a ser gravado:

        'Define variáveis:
        Dim FNum As Integer
        Dim vCurDir As String = ""
        Dim vSaida As String

        'Verifica se DADOS.SYS existe:
        If cdSeguranca1.FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOS.SYS") Then
            vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
            FNum = 1
        Else
            If cdSeguranca1.FileExists("C:\Auditeste\DADOS.SYS") Then
                vCurDir = "C:\Auditeste"
                FNum = 2
            Else
                If cdSeguranca1.FileExists("C:\Arquivos de Programas\Auditeste\DADOS.SYS") Then
                    vCurDir = "C:\Arquivos de Programas\Auditeste"
                    FNum = 3
                Else
                    If cdSeguranca1.FileExists("C:\Program Files\Auditeste\DADOS.SYS") Then
                        vCurDir = "C:\Program Files\Auditeste"
                        FNum = 4
                    Else
                        If cdSeguranca1.FileExists("C:\Windows\DADOS.SYS") Then
                            vCurDir = "C:\Windows"
                            FNum = 5
                        End If
                    End If
                End If
            End If
        End If
        '* ***********************************************************
        '* ***********************************************************
        '* Se for compilar para instalação na VM, tirar o comentário *
        '* Caso contrário, deixar como comentário antes de compilar: *
        '* ***********************************************************
        '* Início
        '* ***********************************************************
        'vCurDir = System.AppDomain.CurrentDomain.BaseDirectory()
        'FNum = cdSeguranca1.FileExists(vCurDir & "\DADOS.SYS")
        '* ***********************************************************
        '* Término
        '* ***********************************************************
        If FNum <> 0 Then   'Arquivo existe
            'Prepara variável para receber os dados lidos:
            vSaida = New String(" "c, 70)
            'Abre e lê o arquivo:
            Dim BvSaida() = FileSystem.ReadAllBytes(vCurDir & "\DADOS.SYS")
            vSaida = System.Text.Encoding.Default.GetString(BvSaida)
            '    FNum = FreeFile
            '    Open vCurDir & "\DADOS.SYS" For Binary Access Read As #FNum     'Le DADOS.SYS
            '    Get #FNum, 1, vSaida
            '    Close #FNum
            'Descriptografa os dados:
            leDadosParaLog = DeCripto(vSaida, 5)
            'Se ocorreu algum problema:
            If Len(leDadosParaLog) <> Len(vSaida) Then
                leDadosParaLog = ""
            End If
        Else
            leDadosParaLog = ""
        End If

    End Function

    Public Function gravaIndice(cUsuario As String, cKey As Integer) As Integer
        'Grava o índice do serviço acessado pelo usuário

        'Define variáveis:
        Dim RC As Integer

        On Error GoTo EgravaIndice
        RC = localizaUsuario(cUsuario, 0)
        If RC <> 0 Then
            gravaIndice = RC
            Exit Function
        End If
        tUsuario.Fields("indice").Value = Val(cKey.ToString("000000")) 'todo verificar se correto
        tUsuario.Update()
        gravaIndice = 0

EgravaIndice:
        If Err.Number Then
            MsgBox(Err.Number & " - " & Err.Description)
            gravaIndice = Err.Number
        End If

    End Function

    Public Function leIndice(cUsuario As String) As Integer
        'Lê o índice do serviço acessado pelo usuário

        'Define variáveis:
        Dim RC As Integer

        On Error GoTo EleIndice
        RC = dbConectaUsuario(0, 0)
        RC = localizaUsuario(cUsuario, 0)
        leIndice = tUsuario.Fields(5).Value

EleIndice:
        If Err.Number Then leIndice = 9999

    End Function

    Public Function lePerfil(cUsuario As String, vIndice As Integer) As Byte
        '* ****************************************************************
        '* Busca o perfil de acesso do usuário ao serviço selecionado(I): *
        '* ****************************************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim vPalavra As String
        Dim vPalavra1 As String
        Dim vPalavra2 As String
        Dim vPalavra3 As String
        Dim segServicos As SegurancaD.segServicos

        On Error GoTo ElePerfil
        RC = dbConectaUsuario(0, 0)
        RC = localizaUsuario(cUsuario, 0)
        segServicos = New SegurancaD.segServicos
        lePerfil = 0

        '* ********************************************************
        '* Primeira parte: Monta os perfis de acesso por serviço: *
        '* ********************************************************
        'Descriptografa a primeira parte de Usuário:
        vPalavra = DeCripto(tUsuario.Fields(1).Value, 5)
        If Len(vPalavra) <> 220 Then      'Ocorreu problemas
            lePerfil = 9
            Exit Function
        End If
        'Descriptografa a segunda parte de Usuário:
        vPalavra1 = DeCripto(tUsuario.Fields(2).Value, 5)
        If Len(vPalavra1) <> 180 Then      'Ocorreu problemas
            lePerfil = 9
            Exit Function
        End If
        'Descriptografa a terceira parte de Usuário:
        vPalavra2 = DeCripto(tUsuario.Fields(3).Value, 5)
        If Len(vPalavra2) <> 180 Then      'Ocorreu problemas
            lePerfil = 9
            Exit Function
        End If
        'Descriptografa a quarta parte de Usuário:
        vPalavra3 = DeCripto(tUsuario.Fields(4).Value, 5)
        If Len(vPalavra3) <> 180 Then      'Ocorreu problemas
            lePerfil = 9
            Exit Function
        End If

        'Separa variáveis:
        RC = separaAtribUsuario(vPalavra)
        If RC <> 0 Then
            lePerfil = 9
            Exit Function
        End If
        'Concatena a segunda parte da propriedade proc (serviços autorizados):
        vPalavra = proc
        proc = New String("0"c, 560)
        proc = vPalavra & vPalavra1 & vPalavra2 & vPalavra3
        '* ******************************************************************
        '* Segunda parte: Localiza o perfil específico do serviço acessado: *
        '* ******************************************************************
        RC = segServicos.dbConecta(0, 0)
        If RC <> 0 Then
            MsgBox("Erro no acesso a segServicos")
            lePerfil = 9
            Exit Function
        End If
        Dim straux As String = Str(vIndice)
        Dim Tamstr = straux.Length
        Select Case Tamstr
        'Select Case Str(vIndice).Length 'Len(Str(vIndice))
            Case 4
                    RC = segServicos.localizaClasServico(Mid(Str(vIndice), 1, 2), Mid(Str(vIndice), 3, 2), 0)
                Case 5
                    RC = segServicos.localizaClasServico(Mid(Str(vIndice), 2, 2), Mid(Str(vIndice), 4, 2), 0)
                Case 6
                    RC = segServicos.localizaClasServico(Mid(Str(vIndice), 1, 2), Mid(Str(vIndice), 3, 2), Mid(Str(vIndice), 5, 2))
                Case 7
                    RC = segServicos.localizaClasServico(Mid(Str(vIndice), 2, 2), Mid(Str(vIndice), 4, 2), Mid(Str(vIndice), 6, 2))
                Case Else
                    MsgBox("Indice do Serviço " & vIndice & " inválido. Processamento interrompido.")
                    lePerfil = 9
                    Exit Function
            End Select
            If RC = 0 Then
            vSelec = segServicos.leServicoSeq(0)
            lePerfil = Mid(proc, Val(Mid(vSelec, 7, 3)), 1)
        Else
            segmensagens.exibeMsg(1013)
            Exit Function
        End If

ElePerfil:
        If Err.Number Then lePerfil = 9

    End Function

    Public Function RetornaPerfis(cUsuario As String) As String
        '* *************************************************************
        '* Retorna o perfil de acesso do usuário de todos os serviços: *
        '* *************************************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim vPalavra As String
        Dim vPalavra1 As String
        Dim vPalavra2 As String
        Dim vPalavra3 As String
        Dim segServicos As SegurancaD.segServicos

        On Error GoTo ERetornaPerfis
        RC = dbConectaUsuario(0, 0)
        RC = localizaUsuario(cUsuario, 0)
        segServicos = New SegurancaD.segServicos
        RetornaPerfis = New String("0"c, 560)

        '* ********************************************************
        '* Primeira parte: Monta os perfis de acesso por serviço: *
        '* ********************************************************
        'Descriptografa a primeira parte de Usuário:
        vPalavra = DeCripto(tUsuario.Fields(1).Value, 5)
        If Len(vPalavra) < 10 Then      'Ocorreu problemas
            RetornaPerfis = 9
            Exit Function
        End If
        'Descriptografa a segunda parte de Usuário:
        vPalavra1 = DeCripto(tUsuario.Fields(2).Value, 5)
        If Len(vPalavra1) < 10 Then      'Ocorreu problemas
            RetornaPerfis = 9
            Exit Function
        End If
        'Descriptografa a terceira parte de Usuário:
        vPalavra2 = DeCripto(tUsuario.Fields(3).Value, 5)
        If Len(vPalavra2) < 10 Then      'Ocorreu problemas
            RetornaPerfis = 9
            Exit Function
        End If
        'Descriptografa a quarta parte de Usuário:
        vPalavra3 = DeCripto(tUsuario.Fields(4).Value, 5)
        If Len(vPalavra3) < 10 Then      'Ocorreu problemas
            RetornaPerfis = 9
            Exit Function
        End If

        'Separa variáveis:
        RC = separaAtribUsuario(vPalavra)
        If RC <> 0 Then
            RetornaPerfis = 9
            Exit Function
        End If
        'Concatena a segunda parte da propriedade proc (serviços autorizados):
        vPalavra = proc
        proc = New String("0"c, 560)
        proc = vPalavra & vPalavra1 & vPalavra2 & vPalavra3
        RetornaPerfis = proc

ERetornaPerfis:
        If Err.Number Then RetornaPerfis = 9

    End Function

    Public Function leSeq(vPrimVez As Integer) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String
        Dim vPalavra As String
        Dim vPalavra1 As String
        Dim vPalavra2 As String
        Dim vPalavra3 As String

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tUsuario.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tUsuario.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tUsuario.EOF Then
            UsuarioA = tUsuario.Fields("UsuarioA").Value
            Descricao = tUsuario.Fields("Descricao").Value
            Descricao1 = tUsuario.Fields("Descricao1").Value
            Descricao2 = tUsuario.Fields("Descricao2").Value
            Descricao3 = tUsuario.Fields("Descricao3").Value
            indice = tUsuario.Fields("indice").Value
            'Separa os atributos do usuário:
            vSelec = DeCripto(Descricao, 5)
            RC = separaAtribUsuario(vSelec)
            '* ****************************************
            '* Monta os perfis de acesso por serviço: *
            '* ****************************************
            'Descriptografa a segunda parte de Usuário:
            vPalavra1 = DeCripto(tUsuario.Fields(2).Value, 5)
            If Len(vPalavra1) < 10 Then      'Ocorreu problemas
                leSeq = 9
                Exit Function
            End If
            'Descriptografa a terceira parte de Usuário:
            vPalavra2 = DeCripto(tUsuario.Fields(3).Value, 5)
            If Len(vPalavra2) < 10 Then      'Ocorreu problemas
                leSeq = 9
                Exit Function
            End If
            'Descriptografa a quarta parte de Usuário:
            vPalavra3 = DeCripto(tUsuario.Fields(4).Value, 5)
            If Len(vPalavra3) < 10 Then      'Ocorreu problemas
                leSeq = 9
                Exit Function
            End If

            'Concatena a segunda parte da propriedade proc (serviços autorizados):
            vPalavra = proc
            proc = New String("0"c, 560)
            proc = vPalavra & vPalavra1 & vPalavra2 & vPalavra3
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe segUsuario - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
