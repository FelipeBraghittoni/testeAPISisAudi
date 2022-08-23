Public Class DBUtils

    Private mvarDb As ADODB.Connection
    Public Property Db() As ADODB.Connection
        Get
            Return mvarDb
        End Get
        Set(ByVal value As ADODB.Connection)
            mvarDb = value
        End Set
    End Property

    Private mvarRC As ADODB.Recordset
    Public Property RC() As ADODB.Recordset
        Get
            Return mvarRC
        End Get
        Set(ByVal value As ADODB.Recordset)
            mvarRC = value
        End Set
    End Property



    Public Function dbConecta(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As String
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset             *
        '* ****************************************

        Dim strConnect As String, strSQL As String

        On Error GoTo EdbConecta

        'dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            'Cria uma instância de ADODB.Connection:
            Db = New ADODB.Connection
            'Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            'cdSeguranca1 = New SegurancaD.cdSeguranca1
            'strConnect = cdSeguranca1.LeDADOSsys(1)
            strConnect = "Provider=SQLOLEDB;Data Source=192.168.1.119,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;"
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        RC = New ADODB.Recordset
        'Abre a tabela:
        Dim vSelect As String = ""
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            RC.Open("Mensagens", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Return "Sucesso ao conectar"
        Else
            If vSelec = "" Then
                dbConecta = 1014
                Return "Erro durante conexão com banco de dados - vSelect vazio"
            Else
                RC.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Return "Sucesso ao logar 2"
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                RC.Close()
                RC.Open(vSelect, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                'MsgBox("Classe Mensagens - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
                Return "Classe Mensagens - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description
            End If
        End If

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

        'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            '* ********************************************
            '* Cria uma instância de ADODB.Connection:    *
            '* ********************************************
            Db = New ADODB.Connection
            'Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            'cdSeguranca1 = New SegurancaD.cdSeguranca1
            'strConnect = cdSeguranca1.LeDADOSsys(1)
            strConnect = "Provider=SQLOLEDB;Data Source=192.168.1.119,1433\sqlexpress;Initial Catalog=Auditeste;User Id=sa;Password=sisaudi@2022;"
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        RC = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            RC.Open("Usuario", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            Return 0
        Else
            If vSelec = "" Then
                dbConectaUsuario = 1014
                Return 1
            Else
                RC.Open(vSelec, Db, ADODB.CursorTypeEnum.adOpenDynamic)
                Return 0
            End If
        End If

EdbConectaUsuario:
        If Err.Number Then
            If Err.Number = 3705 Then
                RC.Close()
                RC.Open("UsuarioBKP", Db, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                Resume Next
            Else
                dbConectaUsuario = Err.Number

                If Err.Number <> -2147467259 Then
                    If Err.Number = 3024 Then
                        dbConectaUsuario = 1046
                    Else
                        dbConectaUsuario = Err.Number
                    End If
                    'MsgBox("Classe Usuario - dbConectaUsuario" & vbCrLf & Err.Number & " - " & Err.Description)
                    Return "Classe Usuario - dbConectaUsuario" & vbCrLf & Err.Number & " - " & Err.Description
                End If
            End If
        End If

    End Function

End Class

