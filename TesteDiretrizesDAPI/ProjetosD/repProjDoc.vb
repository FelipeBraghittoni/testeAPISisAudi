Option Explicit On
Public Class repProjDoc
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850734"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320734"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0734"
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

    Private mvardescricDocto As String
    Public Property descricDocto() As String
        Get
            Return mvardescricDocto
        End Get
        Set(value As String)
            mvardescricDocto = value
        End Set
    End Property

    Private mvarDdtRegistro As String
    Public Property DdtRegistro() As String
        Get
            Return mvarDdtRegistro
        End Get
        Set(value As String)
            mvarDdtRegistro = value
        End Set
    End Property

    Private mvarDdtFim As String
    Public Property DdtFim() As String
        Get
            Return mvarDdtFim
        End Get
        Set(value As String)
            mvarDdtFim = value
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

    Private mvardtRegistro As Date
    Public Property dtRegistro() As Date
        Get
            Return mvardtRegistro
        End Get
        Set(value As Date)
            mvardtRegistro = value
        End Set
    End Property


#End Region

#Region "conexão de banco"
    Public Db As ADODB.Connection
    Public trepProjDoc As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable repProjDoc          *
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
        trepProjDoc = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre repProjDoc como OpenTable
                trepProjDoc.Open("repProjDoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    trepProjDoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                trepProjDoc.Close()

                Select Case tipo
                    Case 0      'Abre repProjDoc como OpenTable
                        trepProjDoc.Open("repProjDoc", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                    Case 1      'Abre como OpenDynaset
                        trepProjDoc.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End Select
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe repProjDoc - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
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
        If Not trepProjDoc.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                trepProjDoc.MoveNext()    'lê 1 linha
            End If
        End If

        'Se não chegou no final do arquivo, carrega propriedades:
        If Not trepProjDoc.EOF Then
            idProjeto = trepProjDoc.Fields("idProjeto").Value
            nomeProjeto = trepProjDoc.Fields("nomeProjeto").Value
            DdtRegistro = trepProjDoc.Fields("DdtRegistro").Value
            DdtFim = trepProjDoc.Fields("DdtFim").Value
            descricDocto = trepProjDoc.Fields("descricDocto").Value
            tipodocto1 = trepProjDoc.Fields("tipodocto1").Value
            tipodocto2 = trepProjDoc.Fields("tipodocto2").Value
            nomeDocto1 = trepProjDoc.Fields("nomeDocto1").Value
            nomeDocto2 = trepProjDoc.Fields("nomeDocto2").Value
            dtRegistro = trepProjDoc.Fields("dtRegistro").Value
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe repProjDoc - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaTxt() As String

        CarregaTxt = trepProjDoc.Fields("txtDocto").Value

    End Function
End Class
