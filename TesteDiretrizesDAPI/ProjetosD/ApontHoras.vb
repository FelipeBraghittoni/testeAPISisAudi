Option Explicit On
Public Class ApontHoras
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537850701"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d320701"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c0701"
#End Region

#Region "Variaveis de ambiente"
    Private mvaridColaborador As Single
    Public Property idColaborador() As Single
        Get
            Return mvaridColaborador
        End Get
        Set(value As Single)
            mvaridColaborador = value
        End Set
    End Property

    Private mvaridProjeto As Double
    Public Property idProjeto() As Double
        Get
            Return mvaridProjeto
        End Get
        Set(value As Double)
            mvaridProjeto = value
        End Set
    End Property

    Private mvardata As Date
    Public Property data() As Date
        Get
            Return mvardata
        End Get
        Set(value As Date)
            mvardata = value
        End Set
    End Property

    Private mvaridAtividade As Short
    Public Property idAtividade() As Short
        Get
            Return mvaridAtividade
        End Get
        Set(value As Short)
            mvaridAtividade = value
        End Set
    End Property


    Private mvaridSAtividade As Short
    Public Property idSAtividade() As Short
        Get
            Return mvaridSAtividade
        End Get
        Set(value As Short)
            mvaridSAtividade = value
        End Set
    End Property

    Private mvarhoraInicio As Date
    Public Property horaInicio() As Date
        Get
            Return mvarhoraInicio
        End Get
        Set(value As Date)
            mvarhoraInicio = value
        End Set
    End Property

    Private mvarhoraFim As Date
    Public Property horaFim() As Date
        Get
            Return mvarhoraFim
        End Get
        Set(value As Date)
            mvarhoraFim = value
        End Set
    End Property

#End Region

#Region "Conexão com Banco"
    Public Db As ADODB.Connection
    Public tApontHoras As ADODB.Recordset
#End Region

    Public Function dbConecta(abreDB As Short, tipo As Short, Optional vSelec As String = "") As Integer
        '* *************************************
        '* * abreDB = se abre ou não o DB:     *
        '* * 0 = abre o DB                     *
        '* * 1 = não abre o DB                 *
        '* *                                   *
        '* * tipo = forma de abrir as tabelas: *
        '* * 0 = OpenTable - ApontHoras        *
        '* * 1 = OpenDynaset                   *
        '* * 3 = OpenTable - OcorHApontHoras   *
        '* *************************************
        Dim RC As Long

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
            'Abre o DB do AudiHoras (parâmetro 2):
            strConnect = cdSeguranca1.LeDADOSsys(2)
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tApontHoras = New ADODB.Recordset
        'Abre a tabela:
        Select Case tipo
            Case 0      'Abre ApontHoras como OpenTable
                tApontHoras.Open("ApontHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
            Case 1      'Abre como OpenDynaset
                If vSelec = "" Then
                    dbConecta = 1014
                Else
                    tApontHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                End If
            Case 3      'Abre OcorHApontHoras como OpenTable
                tApontHoras.Open("OcorHApontHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
        End Select

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tApontHoras.Close()

                If tipo = 1 Then
                    tApontHoras.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Else
                    tApontHoras.Open("ApontHoras", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic, )
                End If
                Resume Next
            Else
                If Err.Number = -2147467259 Then
                    MsgBox("Banco de Dados AppHoras não disponível." & Err.Number & "Contacte o Administrador do AppHoras.")
                Else
                    If Err.Number = 3024 Then dbConecta = 1046
                    MsgBox("Classe ApontHoras - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
                End If
                dbConecta = Err.Number
            End If
        End If

    End Function
    Public Function localiza(colab As Single, data As Date, hora As Date, atualiz As Short) As Integer
        '* ****************************************
        '* * Localiza um registro específico,     *
        '* * baseado na Chave Primária da tabela  *
        '* * ApontHoras.                          *
        '* *                                      *
        '* * colab      = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* * hora       = identif. para pesquisa  *
        '* * atualiz = se atualiza prop.da classe *
        '* ****************************************
        Dim vSelect As String

        On Error GoTo Elocaliza
        vSelect = "SELECT * FROM ApontHoras WHERE fk_Colaborador = " & colab
        vSelect = vSelect & " AND Convert(datetime, data, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        vSelect = vSelect & " AND Convert(nvarchar(8), horaInicio, 14) = '" & String.Format(hora, "hh:mm:ss") & "'"
        tApontHoras.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)

        localiza = 0    'ReturnCode se não houver nenhum problema

        'Não encontrou:
        If tApontHoras.EOF Then
            localiza = 1054
        ElseIf atualiz = 1 Then
            'Se encontrou, carrega propriedades:
            idColaborador = tApontHoras.Fields("fk_Colaborador").Value
            data = tApontHoras.Fields("data").Value
            idProjeto = tApontHoras.Fields("fk_Projeto").Value
            idAtividade = tApontHoras.Fields("fk_Atividade").Value
            idSAtividade = tApontHoras.Fields("fk_SAtividade").Value
            horaInicio = tApontHoras.Fields("horaInicio").Value
            horaFim = tApontHoras.Fields("horaFim").Value
        End If

Elocaliza:
        If Err.Number Then
            If Err.Number = 3705 Then
                tApontHoras.Close()
                tApontHoras.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                localiza = Err.Number
                MsgBox("Classe ApontHoras - localiza" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function


    Public Function inclui(descric As String) As Integer

        Dim vSelec As String

        On Error GoTo Einclui
        tApontHoras.AddNew()
        tApontHoras(0).Value = idColaborador
        tApontHoras(1).Value = data
        tApontHoras(2).Value = idProjeto
        tApontHoras(3).Value = idAtividade
        tApontHoras(4).Value = idSAtividade
        tApontHoras(5).Value = horaInicio
        tApontHoras(6).Value = horaFim
        tApontHoras(7).Value = IIf(descric = "", " ", descric)
        tApontHoras.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            If Err.Number = -2147217873 Then    'Linha já existe
                inclui = -2147217873
            Else
                MsgBox("Classe ApontHoras - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
                inclui = Err.Number
            End If
        End If

    End Function

    Public Function altera(descric As String) As Integer

        On Error GoTo Ealtera
        tApontHoras(0).Value = idColaborador
        tApontHoras(1).Value = data
        tApontHoras(2).Value = idProjeto
        tApontHoras(3).Value = idAtividade
        tApontHoras(4).Value = idSAtividade
        tApontHoras(5).Value = horaInicio
        tApontHoras(6).Value = horaFim
        tApontHoras(7).Value = IIf(descric = "", " ", descric)
        tApontHoras.Update()
        altera = 0

Ealtera:
        If Err.Number Then
            altera = Err.Number
            MsgBox("Classe ApontHoras - altera" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(colab As Single, data As Date) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * ApontHoras.                          *
        '* *                                      *
        '* * colab      = identif. para pesquisa  *
        '* * data       = identif. para pesquisa  *
        '* ****************************************

        Dim RC As Integer
        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro

        vSelec = "DELETE FROM ApontHoras WHERE "
        vSelec = vSelec & " fk_Colaborador = " & colab
        vSelec = vSelec & " AND Convert(datetime, data, 112) = '" & String.Format(data, "yyyymmdd") & "'"
        Db.Execute(vSelec, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ApontHoras - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CarregaDescric() As String

        CarregaDescric = tApontHoras.Fields("Descricao").Value

    End Function
End Class
