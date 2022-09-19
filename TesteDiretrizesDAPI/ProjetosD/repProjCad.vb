Option Explicit On
Public Class repProjCad
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850733"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320733"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0733"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
        End Set
    End Property

    Private mvarnomeProjeto As String
    Public Property nomeProjeto() As String
        Get
            Return mvarnomeProjeto
        End Get
        Set(value As String)
            mvarnomeProjeto = value
        End Set
    End Property

    Private mvarDlocalContr As String
    Public Property DlocalContr() As String
        Get
            Return mvarDlocalContr
        End Get
        Set(value As String)
            mvarDlocalContr = value
        End Set
    End Property

    Private mvarOrganizacao As String
    Public Property Organizacao() As String
        Get
            Return mvarOrganizacao
        End Get
        Set(value As String)
            mvarOrganizacao = value
        End Set
    End Property

    Private mvarSetorOrg As String
    Public Property SetorOrg() As String
        Get
            Return mvarSetorOrg
        End Get
        Set(value As String)
            mvarSetorOrg = value
        End Set
    End Property

    Private mvarEmpresa As String
    Public Property Empresa() As String
        Get
            Return mvarEmpresa
        End Get
        Set(value As String)
            mvarEmpresa = value
        End Set
    End Property

    Private mvarDepto As String
    Public Property Depto() As String
        Get
            Return mvarDepto
        End Get
        Set(value As String)
            mvarDepto = value
        End Set
    End Property

    Private mvarnatureza As String
    Public Property natureza() As String
        Get
            Return mvarnatureza
        End Get
        Set(value As String)
            mvarnatureza = value
        End Set
    End Property

    Private mvardataInicio As String
    Public Property dataInicio() As String
        Get
            Return mvardataInicio
        End Get
        Set(value As String)
            mvardataInicio = value
        End Set
    End Property

    Private mvardataFim As String
    Public Property dataFim() As String
        Get
            Return mvardataFim
        End Get
        Set(value As String)
            mvardataFim = value
        End Set
    End Property

    Private mvarGerComerc As String
    Public Property GerComerc() As String
        Get
            Return mvarGerComerc
        End Get
        Set(value As String)
            mvarGerComerc = value
        End Set
    End Property

    Private mvarGerTecnico As String
    Public Property GerTecnico() As String
        Get
            Return mvarGerTecnico
        End Get
        Set(value As String)
            mvarGerTecnico = value
        End Set
    End Property

    Private mvarGerProjeto As String
    Public Property GerProjeto() As String
        Get
            Return mvarGerProjeto
        End Get
        Set(value As String)
            mvarGerProjeto = value
        End Set
    End Property

    Private mvarmodalidade As String
    Public Property modalidade() As String
        Get
            Return mvarmodalidade
        End Get
        Set(value As String)
            mvarmodalidade = value
        End Set
    End Property

#End Region

#Region "conexão com banco"
    Public Db As ADODB.Connection
    Public trepProjCad As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjCad          *
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
        trepProjCad = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjCad como OpenTable
                trepProjCad.Open("repProjCad", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjCad.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjCad.Close()

                Select Case tipo
                    Case 0      'Abre repProjCad como OpenTable
                        trepProjCad.Open("repProjCad", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjCad.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjCad - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjCad.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjCad.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjCad.EOF Then
            idProjeto = trepProjCad.Fields("idProjeto").Value
            DlocalContr = trepProjCad.Fields("DlocalContr").Value
            nomeProjeto = trepProjCad.Fields("nomeProjeto").Value
            Organizacao = trepProjCad.Fields("Organizacao").Value
            SetorOrg = trepProjCad.Fields("SetorOrg").Value
            Empresa = trepProjCad.Fields("Empresa").Value
            Depto = trepProjCad.Fields("Depto").Value
            natureza = trepProjCad.Fields("natureza").Value
            dataInicio = trepProjCad.Fields("dataInicio").Value
            dataFim = trepProjCad.Fields("dataFim").Value
            GerComerc = trepProjCad.Fields("GerComerc").Value
            GerTecnico = trepProjCad.Fields("GerTecnico").Value
            GerProjeto = trepProjCad.Fields("GerProjeto").Value
            modalidade = trepProjCad.Fields("modalidade").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjCad - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
