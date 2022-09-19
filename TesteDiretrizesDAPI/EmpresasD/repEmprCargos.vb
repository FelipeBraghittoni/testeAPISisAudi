Option Explicit On
Public Class repEmprCargos
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850321"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320321"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0321"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridCargo As Short
    Public Property idCargo() As Short
        Get
            Return mvaridCargo
        End Get
        Set(value As Short)
            mvaridCargo = value
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


    Private mvarDconfiancaCargo As String
    Public Property DconfiancaCargo() As String
        Get
            Return mvarDconfiancaCargo
        End Get
        Set(value As String)
            mvarDconfiancaCargo = value
        End Set
    End Property


    Private mvarpermanMinCargo As Short
    Public Property permanMinCargo() As Short
        Get
            Return mvarpermanMinCargo
        End Get
        Set(value As Short)
            mvarpermanMinCargo = value
        End Set
    End Property


    Private mvarpermanMaxCargo As Short
    Public Property permanMaxCargo() As Short
        Get
            Return mvarpermanMaxCargo
        End Get
        Set(value As Short)
            mvarpermanMaxCargo = value
        End Set
    End Property


    Private mvarsalarioCargoMin As Double
    Public Property salarioCargoMin() As Double
        Get
            Return mvarsalarioCargoMin
        End Get
        Set(value As Double)
            mvarsalarioCargoMin = value
        End Set
    End Property


    Private mvarsalarioCargo1 As Double
    Public Property salarioCargo1() As Double
        Get
            Return mvarsalarioCargo1
        End Get
        Set(value As Double)
            mvarsalarioCargo1 = value
        End Set
    End Property


    Private mvarsalarioCargoMed As Double
    Public Property salarioCargoMed() As Double
        Get
            Return mvarsalarioCargoMed
        End Get
        Set(value As Double)
            mvarsalarioCargoMed = value
        End Set
    End Property


    Private mvarsalarioCargo3 As Double
    Public Property salarioCargo3() As Double
        Get
            Return mvarsalarioCargo3
        End Get
        Set(value As Double)
            mvarsalarioCargo3 = value
        End Set
    End Property


    Private mvarsalarioCargoMax As Double
    Public Property salarioCargoMax() As Double
        Get
            Return mvarsalarioCargoMax
        End Get
        Set(value As Double)
            mvarsalarioCargoMax = value
        End Set
    End Property


    Private mvarnomeMoeda As String
    Public Property nomeMoeda() As String
        Get
            Return mvarnomeMoeda
        End Get
        Set(value As String)
            mvarnomeMoeda = value
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

#Region "Conexão com banco"
    Public Db As ADODB.Connection
    Public trepEmprCargos As ADODB.Recordset
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
        trepEmprCargos = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable
            trepEmprCargos.Open("repEmprCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        Else                'Aberto como OpenDynaset
            If tipo = 2 Then
                vSelec = "SELECT * FROM repEmprCargos ORDER BY nomeCargo"
                trepEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            ElseIf tipo = 1 Then
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprCargos.Close()

                If tipo = 1 Or tipo = 2 Then
                    trepEmprCargos.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprCargos.Open("repEmprCargos", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                'dbConecta = Err
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprCargos - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprCargos.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprCargos.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprCargos.EOF Then
            idCargo = trepEmprCargos.Fields("idCargo").Value
            nomeCargo = trepEmprCargos.Fields("nomeCargo").Value
            DconfiancaCargo = trepEmprCargos.Fields("DconfiancaCargo").Value
            permanMinCargo = trepEmprCargos.Fields("permanMinCargo").Value
            permanMaxCargo = trepEmprCargos.Fields("permanMaxCargo").Value
            salarioCargoMin = trepEmprCargos.Fields("salarioCargoMin").Value
            salarioCargo1 = trepEmprCargos.Fields("salarioCargo1").Value
            salarioCargoMed = trepEmprCargos.Fields("salarioCargoMed").Value
            salarioCargo3 = trepEmprCargos.Fields("salarioCargo3").Value
            salarioCargoMax = trepEmprCargos.Fields("salarioCargoMax").Value
            nomeMoeda = trepEmprCargos.Fields("nomeMoeda").Value
            dtInicio = trepEmprCargos.Fields("dtInicio").Value
            dDtFim = trepEmprCargos.Fields("dDtFim").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprCargos - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaPreReq() As String

        CarregaPreReq = trepEmprCargos.Fields("preRequisitosCargo").Value

    End Function

    Public Function CarregaResponsab() As String

        CarregaResponsab = trepEmprCargos.Fields("FresponsabCargo").Value

    End Function

    Public Function CarregaTarefas() As String

        CarregaTarefas = trepEmprCargos.Fields("tarefasCargo").Value

    End Function

End Class
