Option Explicit On
Public Class repEmprDeptos

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850323"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320323"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0323"
#End Region

#Region "Variaveis de ambiente"

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
            Return mvarramal2Depto
        End Get
        Set(value As Short)
            mvarramal2Depto = value
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
            Return mvarrefeitorioDepto
        End Get
        Set(value As String)
            mvarrefeitorioDepto = value
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


#Region "Geração do excel"

    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvardtFim
        End Get
        Set(value As String)
            mvardtFim = value
        End Set
    End Property

    Private mvardFone1Depto As String
    Public Property dFone1Depto() As String
        Get
            Return mvardFone1Depto
        End Get
        Set(value As String)
            mvardFone1Depto = value
        End Set
    End Property

    Private mvardFone2Depto As String
    Public Property dFone2Depto() As String
        Get
            Return mvardFone2Depto
        End Get
        Set(value As String)
            mvardFone2Depto = value
        End Set
    End Property

    Private mvarnomeDeptoSup As String
    Public Property nomeDeptoSup() As String
        Get
            Return mvarnomeDeptoSup
        End Get
        Set(value As String)
            mvarnomeDeptoSup = value
        End Set
    End Property

    Private mvarnomeGerente As String
    Public Property nomeGerente() As String
        Get
            Return mvarnomeGerente
        End Get
        Set(value As String)
            mvarnomeGerente = value
        End Set
    End Property

    Private mvarnomeRegimeTrab As String
    Public Property nomeRegimeTrab() As String
        Get
            Return mvarnomeRegimeTrab
        End Get
        Set(value As String)
            mvarnomeRegimeTrab = value
        End Set
    End Property

    Private mvarnomeCtaCtbil As String
    Public Property nomeCtaCtbil() As String
        Get
            Return mvarnomeCtaCtbil
        End Get
        Set(value As String)
            mvarnomeCtaCtbil = value
        End Set
    End Property

    Private mvardDiaInicTrabDepto As String
    Public Property dDiaInicTrabDepto() As String
        Get
            Return mvardDiaInicTrabDepto
        End Get
        Set(value As String)
            mvardDiaInicTrabDepto = value
        End Set
    End Property

    Private mvardDiaFimTrabDepto As String
    Public Property dDiaFimTrabDepto() As String
        Get
            Return mvardDiaFimTrabDepto
        End Get
        Set(value As String)
            mvardDiaFimTrabDepto = value
        End Set
    End Property
#End Region

#Region "Conexão com banco"

    Public Db As ADODB.Connection
    Public trepEmprDeptos As ADODB.Recordset
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
        trepEmprDeptos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            trepEmprDeptos.Open("repEmprDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM repEmprDeptos ORDER BY idEmpresa, nomeDepto"
                trepEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprDeptos.Close()

                If tipo = 1 Or tipo = 2 Then
                    trepEmprDeptos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprDeptos.Open("repEmprDeptos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprDeptos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprDeptos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprDeptos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprDeptos.EOF Then
            idEmpresa = trepEmprDeptos.Fields("idEmpresa").Value
            idDepto = trepEmprDeptos.Fields("idDepto").Value
            nomeDepto = trepEmprDeptos.Fields("nomeDepto").Value
            idCtaContabil = trepEmprDeptos.Fields("idCtaCtbil").Value
            fone1Depto = trepEmprDeptos.Fields("fone1Depto").Value
            ramal1Depto = trepEmprDeptos.Fields("ramal1Depto").Value
            fone2Depto = trepEmprDeptos.Fields("fone2Depto").Value
            ramal2Depto = trepEmprDeptos.Fields("ramal2Depto").Value
            faxDepto = trepEmprDeptos.Fields("faxDepto").Value
            eMailDepto = trepEmprDeptos.Fields("eMailDepto").Value
            idDeptoSup = trepEmprDeptos.Fields("idDeptoSup").Value
            idColabResp = trepEmprDeptos.Fields("idColabResp").Value
            regimeTrabDepto = trepEmprDeptos.Fields("regimeTrabDepto").Value
            diaInicTrabDepto = trepEmprDeptos.Fields("diaInicTrabDepto").Value
            diaFimTrabDepto = trepEmprDeptos.Fields("diaFimTrabDepto").Value
            refeitorioDepto = trepEmprDeptos.Fields("refeitorioDepto").Value
            horaInicTrabDepto = trepEmprDeptos.Fields("horaInicTrabDepto").Value
            horaInicDescDepto = trepEmprDeptos.Fields("horaInicDescDepto").Value
            horaFimDescDepto = trepEmprDeptos.Fields("horaFimDescDepto").Value
            horaFimTrabDepto = trepEmprDeptos.Fields("horaFimTrabDepto").Value
            dtInicio = trepEmprDeptos.Fields("dtInicio").Value
            dtFim = trepEmprDeptos.Fields("dtFim").Value
            nomeEmpresa = trepEmprDeptos.Fields("nomeEmp").Value
            dFone1Depto = trepEmprDeptos.Fields("dFone1Depto").Value
            dFone2Depto = trepEmprDeptos.Fields("dFone2Depto").Value
            nomeDeptoSup = trepEmprDeptos.Fields("nomeDeptoSup").Value
            nomeGerente = trepEmprDeptos.Fields("nomeGerente").Value
            nomeRegimeTrab = trepEmprDeptos.Fields("nomeRegimeTrab").Value
            nomeCtaCtbil = trepEmprDeptos.Fields("nomeCtaCtbil").Value
            dDiaInicTrabDepto = trepEmprDeptos.Fields("dDiaInicTrabDepto").Value
            dDiaFimTrabDepto = trepEmprDeptos.Fields("dDiaFimTrabDepto").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprDeptos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function altera(responsab As String) As Integer

        On Error GoTo Ealtera
        trepEmprDeptos(2).Value = nomeDepto
        trepEmprDeptos(3).Value = idCtaContabil
        trepEmprDeptos(4).Value = fone1Depto
        trepEmprDeptos(5).Value = ramal1Depto
        trepEmprDeptos(6).Value = fone2Depto
        trepEmprDeptos(7).Value = ramal2Depto
        trepEmprDeptos(8).Value = faxDepto
        trepEmprDeptos(9).Value = eMailDepto
        trepEmprDeptos(10).Value = idDeptoSup
        trepEmprDeptos(11).Value = idColabResp
        'If indice = 602 Then    'Alteração no cadastro de deptos.:
        trepEmprDeptos(12).Value = IIf(responsab = "", " ", responsab)
        'End If
        trepEmprDeptos(13).Value = regimeTrabDepto
        trepEmprDeptos(14).Value = diaInicTrabDepto
        trepEmprDeptos(15).Value = diaFimTrabDepto
        trepEmprDeptos(16).Value = refeitorioDepto
        trepEmprDeptos(17).Value = horaInicTrabDepto
        trepEmprDeptos(18).Value = horaInicDescDepto
        trepEmprDeptos(19).Value = horaFimDescDepto
        trepEmprDeptos(20).Value = horaFimTrabDepto
        trepEmprDeptos.Fields("dtInicio").Value = dtInicio
        trepEmprDeptos.Fields("dtFim").Value = dtFim
        trepEmprDeptos.Fields("nomeEmpresa").Value = nomeEmpresa
        trepEmprDeptos.Fields("dFone1Depto").Value = dFone1Depto
        trepEmprDeptos.Fields("dFone2Depto").Value = dFone2Depto
        trepEmprDeptos.Fields("nomeDeptoSup").Value = nomeDeptoSup
        trepEmprDeptos.Fields("nomeGerente").Value = nomeGerente
        trepEmprDeptos.Fields("nomeRegimeTrab").Value = nomeRegimeTrab
        trepEmprDeptos.Fields("nomeCtaCtbil").Value = nomeCtaCtbil
        trepEmprDeptos.Fields("dDiaInicTrabDepto").Value = dDiaInicTrabDepto
        trepEmprDeptos.Fields("dDiaFimTrabDepto").Value = dDiaFimTrabDepto
        trepEmprDeptos.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe repEmprDeptos - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function carregaMemos() As String
        'Carrega o valor do campo txtResponsab no form:

        carregaMemos = trepEmprDeptos.Fields("responsabilDepto").Value

    End Function


End Class
