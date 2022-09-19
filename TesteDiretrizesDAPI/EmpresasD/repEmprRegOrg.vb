Option Explicit On
Public Class repEmprRegOrg
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850327"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320327"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0327"
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

    Private mvarnomeEmpresa As String
    Public Property nomeEmpresa() As String
        Get
            Return mvarnomeEmpresa
        End Get
        Set(value As String)
            mvarnomeEmpresa = value
        End Set
    End Property

    Private mvardescricDoctoGer As String
    Public Property descricDoctoGer() As String
        Get
            Return mvardescricDoctoGer
        End Get
        Set(value As String)
            mvardescricDoctoGer = value
        End Set
    End Property

    Private mvarnumDoctoEmpr As String
    Public Property numDoctoEmpr() As String
        Get
            Return mvarnumDoctoEmpr
        End Get
        Set(value As String)
            mvarnumDoctoEmpr = value
        End Set
    End Property

    Private mvarcomplemen1Docto As String
    Public Property complemen1Docto() As String
        Get
            Return mvarcomplemen1Docto
        End Get
        Set(value As String)
            mvarcomplemen1Docto = value
        End Set
    End Property

    Private mvarcomplemenD1 As String
    Public Property complemenD1() As String
        Get
            Return mvarcomplemenD1
        End Get
        Set(value As String)
            mvarcomplemenD1 = value
        End Set
    End Property


    Private mvarcomplemen2Docto As String
    Public Property complemen2Docto() As String
        Get
            Return mvarcomplemen2Docto
        End Get
        Set(value As String)
            mvarcomplemen2Docto = value
        End Set
    End Property


    Private mvarcomplemenD2 As String
    Public Property complemenD2() As String
        Get
            Return mvarcomplemenD2
        End Get
        Set(value As String)
            mvarcomplemenD2 = value
        End Set
    End Property

    Private mvarcomplemen3Docto As String
    Public Property complemen3Docto() As String
        Get
            Return mvarcomplemen3Docto
        End Get
        Set(value As String)
            mvarcomplemen3Docto = value
        End Set
    End Property


    Private mvarcomplemenD3 As String
    Public Property complemenD3() As String
        Get
            Return mvarcomplemenD3
        End Get
        Set(value As String)
            mvarcomplemenD3 = value
        End Set
    End Property

#End Region

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public trepEmprRegOrg As ADODB.Recordset
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
        trepEmprRegOrg = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then        'Aberto como OpenTable
            trepEmprRegOrg.Open("repEmprRegOrg", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        ElseIf tipo = 1 Then    'Aberto como OpenDynaset
            If vSelec = "" Then
                dbConecta = 1014
            Else
                trepEmprRegOrg.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepEmprRegOrg.Close()

                If tipo = 1 Then
                    trepEmprRegOrg.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    trepEmprRegOrg.Open("repEmprRegOrg", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repEmprRegOrg - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepEmprRegOrg.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepEmprRegOrg.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepEmprRegOrg.EOF Then
            idEmpresa = trepEmprRegOrg.Fields("idEmpresa").Value
            nomeEmpresa = trepEmprRegOrg.Fields("nomeEmpresa").Value
            descricDoctoGer = trepEmprRegOrg.Fields("descricDoctoGer").Value
            numDoctoEmpr = trepEmprRegOrg.Fields("numDoctoEmpr").Value
            complemen1Docto = trepEmprRegOrg.Fields("complemen1Docto").Value
            'complemenD1 = trepEmprRegOrg.Fields("complemenD1").Value
            complemen2Docto = trepEmprRegOrg.Fields("complemen2Docto").Value
            'complemenD2 = trepEmprRegOrg.Fields("complemenD2").Value
            complemen3Docto = trepEmprRegOrg.Fields("complemen3Docto").Value
            'complemenD3 = trepEmprRegOrg.Fields("complemenD3").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repEmprRegOrg - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

End Class
