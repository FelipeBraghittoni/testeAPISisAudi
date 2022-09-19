Option Explicit On
Public Class repEmprAlocCargos

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850317"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320317"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0317"
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


    Private mvaridCargo As Short
    Public Property idCargo() As Short
        Get
            Return mvaridCargo
        End Get
        Set(value As Short)
            mvaridCargo = value
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


    Private mvarnomeCargo As String
    Public Property nomeCargo() As String
        Get
            Return mvarnomeCargo
        End Get
        Set(value As String)
            mvarnomeCargo = value
        End Set
    End Property


    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvarnomeEmpresa
        End Get
        Set(value As String)
            mvarnomeEmpresa = value
        End Set
    End Property


    Private mvardDtFim As String
    Public Property dDtFim() As String
        Get
            Return mvardDtFim
        End Get
        Set(value As String)
            mvardDtFim = value
        End Set
    End Property


#End Region

#Region "Conexão com Banco"

    Public Db As ADODB.Connection
    Public trepEmprAlocCargos As ADODB.Recordset
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
        trepEmprAlocCargos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            trepEmprAlocCargos.Open("repEmprAlocCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                trepEmprAlocCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprAlocCargos.Close()

                If tipo = 1 Then
                    trepEmprAlocCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprAlocCargos.Open("repEmprAlocCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                'dbConecta = Err
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprAlocCargos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprAlocCargos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprAlocCargos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprAlocCargos.EOF Then
            idEmpresa = trepEmprAlocCargos.Fields("idEmpresa").Value
            idCargo = trepEmprAlocCargos.Fields("idCargo").Value
            dtInicio = trepEmprAlocCargos.Fields("dtInicio").Value
            dtFim = trepEmprAlocCargos.Fields("dtFim").Value
            nomeCargo = trepEmprAlocCargos.Fields("nomeCargo").Value
            nomeEmpresa = trepEmprAlocCargos.Fields("nomeEmpresa").Value
            dDtFim = trepEmprAlocCargos.Fields("dDtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprAlocCargos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function


End Class
