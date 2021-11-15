VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMercader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%   Creado por El Yind   %%%%%
'%%%%%       23/01/2012       %%%%%
'%%%%% Build: 3               %%%%%
'%%%%% Modificado: 24/02/2012 %%%%%
'%%%%% Modificado: 16/05/2021 NeoSefirot
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit
Private MAX_MANA As Integer
Private ORO_INICIAL As Integer
Private INCREMENTO_ORO As Integer
Private MAX_ORO As Integer
Private TIEMPO_ESPERA As Integer 'Minutos
Private NUM_ROJAS As Integer
Private NUM_AZULES As Integer

'Neo Nuevas variables que setea el usuario para NPC
Private Clan As String
Private TiraRemo As Integer
Private TiraInmo As Integer
Private TiraApoca As Integer
Private TiraDescarga As Integer
Private TiraTormentaFuego As Integer
Private TiraVerInvi As Integer
Private MsjAgresorInvi As String
Private MsjNuevoAgresor As String
Private MsjAgresorRandom1 As String
Private MsjAgresorRandom2 As String
Private MsjAgresorRandom3 As String
Private MsjAvisoPosibleAgresor As String
Private MsjNoMolestes As String
Private MsjAlHacerClick As String
Private Criatura As Integer
'Tiempo de vida del trabajador
'Hora del login del trabajador
'Ultima ciudad donde se creo el trabajador
Private FinalCamino As Boolean
Private Pescador As Boolean
Private Talador As Boolean
Private Herrero As Boolean
Private Carpintero As Boolean
Private Minero As Boolean
Private GiroInicio As String
Private GiroFin As String
Private YaGire As Boolean
Private TiempoHablar As Date
Private CadaCuantoHabla As Integer
Private TiempoComer As Date
Private CadaCuantoComo As Integer
'Fin nuevas variables del npc

Private Camino() As Position
Private pCantPasos As Byte
Private pPaso As Byte
Private pNpcIndex As Integer
Private pNroNPC As Integer
Private pRuta As clsTree
Private startX As Integer, startY As Integer
Private Yendo As Boolean
Private Tiempo As Date
Private Agresores As Dictionary
Private Direccion As Byte
Private Recalcular As Boolean
Private DestinoX As Integer
Private DestinoY As Integer
Private LastAgresor As Integer
Private estado As Byte
Private Destino(1 To 2) As String
Private TiempoTrabajo As Date
Private SegundosTrabajo As Integer
'INVENTARIO
Private potasRojas As Integer
Private potasAzules As Integer
Private Oro As Long


'DEFENSA
Private CPRemo As Integer
Private CPDialogo As Integer
Private CPAtaca As Integer
Private BardeoInvis As Boolean
Private Mana As Integer
Private Estrategia As Byte 'Se usa para definir la manera en que el npc va a jugar
Private Combear As Boolean 'Se usa cuando dio un golpe para priorizar un ataque magico
Private Escapar As Boolean 'Se usa para que si le esta pegando cuerpo a cuerpo y el random dio, se corra para no quedarse siempre ahi
Private Emergencia As Boolean 'Se usa por si el npc le queda poca vida que no ataque hasta tomar una pocion
Private PasosEnemigo As Integer 'Cuenta cuantos pasos hace con alguien cerca que rompa las bolas

Private LanzoAtaque As Integer
Private Trabajador As Boolean

'####################NEO######################
'Tiempos minimos y maximos configurados en el dat de configuracion
Private TiempoTotalTrabajo As Integer
Private TiempoTrabajado As Date
'Tiempos de login
Private TiempoTotalLogin As Integer
'Tiempos de resucitar
Private TiempoTotalResucitar As Integer
'Dia y fecha que se conecto el trabajador
Private CuandoMeConecte As Date
'Variable que indica si el npc esta online o no
Private NpcOnlineNpc As Boolean
Private TIEMPO_TOTAL_TRABAJO As Integer
Private PrimeraVuelta As Boolean
'########################################################
'Hechizo de invocacion
Private NroHechizoInvocacion As Integer
'Hechizos de defensa
Private Const HECHI_REMO As Integer = 10
Private Const HECHI_REMOINVI As Integer = 38
Private Const HECHI_CELERIDAD As Integer = 18
'Hechizos de ataque
Private Const HECHI_INMO As Integer = 24
Private Const HECHI_APOCA As Integer = 25
Private Const HECHI_DESCARGA As Integer = 23
Private Const HECHI_TORMENTA_FUEGO As Integer = 15
'Hechizos oscuros
Private Const HECHI_HAMBRE_IGOR As Integer = 13
Private Const HECHI_CEGUERA As Integer = 30
Private Const HECHI_ESTUPIDEZ As Integer = 31
Private CalcularInmo As Boolean


Public Property Get NpcIndex() As Integer
NpcIndex = pNpcIndex
End Property

Public Property Let NpcIndex(ByVal lNewValue As Integer)
pNpcIndex = lNewValue
End Property

Public Property Get NpcNum() As Integer
NpcNum = pNroNPC
End Property

Public Property Let NpcNum(ByVal lNewValue As Integer)
pNroNPC = lNewValue
End Property

'Public Sub Init(ByVal NpcIndex As Integer, ByVal Mana As Integer, ByVal OroInicial As Integer, ByVal OroMaximo As Integer, ByVal IncrementoOro As Integer, ByVal TiempoEspera As Integer, ByVal PocionesRojas As Integer, ByVal PocionesAzules As Integer, ByVal Ruta As String, ByVal Destino1 As String, ByVal Destino2 As String)
Public Sub Init(ByVal NPCI As Integer, ByVal esTrabajador As Boolean)
'Antes de trabajar con la ruta, selecciono una
Dim Rutas() As String
Dim rutaSeleccionada As Integer
rutaSeleccionada = 1
PrimeraVuelta = True
If esTrabajador = False Then
    'No es un trabajador, o sea es un NPC de lucha
    Trabajador = esTrabajador
    pNroNPC = NpcInteligentes(NPCI).NumeroNPC
    'Se setean datos
    MAX_MANA = NpcInteligentes(NPCI).Mana
    ORO_INICIAL = NpcInteligentes(NPCI).OroInicial
    INCREMENTO_ORO = NpcInteligentes(NPCI).IncrementoOro
    MAX_ORO = NpcInteligentes(NPCI).OroMaximo
    TIEMPO_ESPERA = NpcInteligentes(NPCI).TiempoEsperaMinuto
    NUM_ROJAS = NpcInteligentes(NPCI).PocionesRojas
    NUM_AZULES = NpcInteligentes(NPCI).PocionesAzules
    'Si el npc tiene mas de una ruta, selecciona una al azar
    If NpcInteligentes(NPCI).NroRutas > 1 Then
        rutaSeleccionada = RandomNumber(1, NpcInteligentes(NPCI).NroRutas)
    End If
    'Parseo el camino
    Rutas = Split(NpcInteligentes(NPCI).Ruta(rutaSeleccionada).Camino, ";")
    'Se setean los destinos 1 y 2, en caso de tenerlos
    If pNroNPC = 617 Then
        Destino(1) = "Ullathorpe"
        Destino(2) = "Banderbill"
    ElseIf pNroNPC = 618 Then
        Destino(1) = "Nix"
        Destino(2) = "Arkhein"
    Else
        Destino(1) = ""
        Destino(2) = ""
    End If
    TIEMPO_TOTAL_TRABAJO = 0
    'Se setean variables de ataque
    Clan = NpcInteligentes(NPCI).Clan
    'Se ponen mensajes al mercader
    InitMensaje NPCI
    
    TiraRemo = NpcInteligentes(NPCI).TiraRemo
    TiraInmo = NpcInteligentes(NPCI).TiraInmo
    TiraApoca = NpcInteligentes(NPCI).TiraApoca
    TiraDescarga = NpcInteligentes(NPCI).TiraDescarga
    TiraTormentaFuego = NpcInteligentes(NPCI).TiraTormentaFuego
    TiraVerInvi = NpcInteligentes(NPCI).TiraVerInvi
    NroHechizoInvocacion = NpcInteligentes(NPCI).NroHechizoInvocacion
Else
    'Neo Inicia un nuevo trabajador
    'Es un npc trabajador
    pNroNPC = NpcTrabajadores(NPCI).NumeroNPC
    FinalCamino = False
    Trabajador = esTrabajador
    YaGire = False
    'Se setean los giros
    GiroInicio = ""
    GiroFin = ""
    Clan = ""
    'Se setean todas las profesiones
    Pescador = False
    Talador = False
    Minero = False
    Carpintero = False
    Herrero = False
    TIEMPO_TOTAL_TRABAJO = 0
    'En el caso de npc trabajador, selecciono una ruta al azar, depende si es pescador o talador
    If NpcTrabajadores(NPCI).Profesion = "PESCADOR" Then
        'Es Pescador
        Pescador = True
        rutaSeleccionada = RandomNumber(1, NumeroRutaPescador)
        'Parseo el camino
        Rutas = Split(RutaPescadores(rutaSeleccionada).Ruta, ";")
        GiroInicio = RutaPescadores(rutaSeleccionada).DireccionInicio
        GiroFin = RutaPescadores(rutaSeleccionada).DireccionFin
    ElseIf NpcTrabajadores(NPCI).Profesion = "TALADOR" Then
        'Es Talador
        Talador = True
        rutaSeleccionada = RandomNumber(1, NumeroRutaTalador)
        'Parseo el camino
        Rutas = Split(RutaTaladores(rutaSeleccionada).Ruta, ";")
        GiroInicio = RutaTaladores(rutaSeleccionada).DireccionInicio
        GiroFin = RutaTaladores(rutaSeleccionada).DireccionFin
     ElseIf NpcTrabajadores(NPCI).Profesion = "MINERO" Then
        'Es minero
        Minero = True
        rutaSeleccionada = RandomNumber(1, NumeroRutaMinero)
        'Parseo el camino
        Rutas = Split(RutaMineros(rutaSeleccionada).Ruta, ";")
        GiroInicio = RutaMineros(rutaSeleccionada).DireccionInicio
        GiroFin = RutaMineros(rutaSeleccionada).DireccionFin
    ElseIf NpcTrabajadores(NPCI).Profesion = "CARPINTERO" Then
        'Es carpintero
        Carpintero = True
        rutaSeleccionada = RandomNumber(1, NumeroRutaCarpintero)
        'Parseo el camino
        Rutas = Split(RutaCarpinteros(rutaSeleccionada).Ruta, ";")
        GiroInicio = RutaCarpinteros(rutaSeleccionada).DireccionInicio
        GiroFin = RutaCarpinteros(rutaSeleccionada).DireccionFin
    ElseIf NpcTrabajadores(NPCI).Profesion = "HERRERO" Then
        'Es herrero
        Herrero = True
        rutaSeleccionada = RandomNumber(1, NumeroRutaHerrero)
        'Parseo el camino
        Rutas = Split(RutaHerreros(rutaSeleccionada).Ruta, ";")
        GiroInicio = RutaHerreros(rutaSeleccionada).DireccionInicio
        GiroFin = RutaHerreros(rutaSeleccionada).DireccionFin
    End If
    'Se setean los datos
    InitDatoTrabajador
End If

'Se procesan las rutas
Dim i As Integer
pCantPasos = UBound(Rutas)
ReDim Camino(pCantPasos)
For i = 0 To pCantPasos
    Camino(i).X = val(ReadField(1, Rutas(i), 44))
    Camino(i).Y = val(ReadField(2, Rutas(i), 44))
Next i
    
Set Agresores = New Dictionary


ReSpawn

End Sub
Private Sub SetTiempoRandom()
'Neo Por cada npc, se toma un tiempo random en base a los tiempos globales
'Tiempo total que el npc trabaja
TiempoTotalTrabajo = GetTiempoTrabajo
'Cada cuando ejecuta el ruido del trabajo
TiempoTrabajo = Now
SegundosTrabajo = GetIntarvaloTrabajo
'Cada cuanto habla el trabajador
TiempoHablar = Now
CadaCuantoHabla = GetTiempoHablar
'Comer y beber
TiempoComer = Now
CadaCuantoComo = GetTiempoComerBeber
End Sub
Private Sub InitDatoTrabajador()
'Neo Se inician los datos para un trabajador
    MAX_MANA = 2000
    ORO_INICIAL = 1
    INCREMENTO_ORO = 1
    MAX_ORO = 1000
    TIEMPO_ESPERA = 0
    NUM_ROJAS = 800
    NUM_AZULES = 800
    
    MsjAgresorInvi = ""
    MsjNuevoAgresor = ""
    MsjAgresorRandom1 = ""
    MsjAgresorRandom2 = ""
    MsjAgresorRandom3 = ""
    MsjAvisoPosibleAgresor = ""
    MsjNoMolestes = ""
    MsjAlHacerClick = ""
    
    TiraRemo = 0
    TiraInmo = 0
    TiraApoca = 0
    TiraDescarga = 0
    TiraTormentaFuego = 0
    TiraVerInvi = 0
    NroHechizoInvocacion = 0
    
    'Se setean los destinos 1 y 2, en caso de tenerlos
    Destino(1) = ""
    Destino(2) = ""
End Sub
Private Sub InitMensaje(ByVal NPCI As Integer)
'Neo Setea mensajes para npc, configurados en el .dat
MsjAgresorInvi = NpcInteligentes(NPCI).MsjAgresorInvi
MsjNuevoAgresor = NpcInteligentes(NPCI).MsjNuevoAgresor
MsjAgresorRandom1 = NpcInteligentes(NPCI).MsjAgresorRandom1
MsjAgresorRandom2 = NpcInteligentes(NPCI).MsjAgresorRandom2
MsjAgresorRandom3 = NpcInteligentes(NPCI).MsjAgresorRandom3
MsjAvisoPosibleAgresor = NpcInteligentes(NPCI).MsjAvisoPosibleAgresor
MsjNoMolestes = NpcInteligentes(NPCI).MsjNoMolestes
MsjAlHacerClick = NpcInteligentes(NPCI).MsjAlHacerClick
End Sub

Private Function SiguientePaso() As Boolean

If Yendo Then
    If pPaso < pCantPasos Then
        'En camino al final de la ruta
        EstoyCaminando
        pPaso = pPaso + 1
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        BuscarCamino
        SiguientePaso = True
    Else
        'Llegue al final de la ruta
        'Neo Si es trabajador, lo giro y seteo que llego al fin
        If Trabajador Then
            LlegueAlFinal
        End If
        Yendo = False
        ResetMercader
        SiguientePaso = False
    End If
Else
    
    If pPaso > 0 Then
        'En camino al INICIO de la ruta
        EstoyCaminando
        pPaso = pPaso - 1
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        BuscarCamino
        SiguientePaso = True
    Else
        EstoyCaminando
        'Llegue al INICIO de la ruta
        If Trabajador Then
            LlegueAlInicio
        End If
        Yendo = True
        ResetMercader
        SiguientePaso = False
    End If
End If
End Function
Private Sub LlegueAlInicio()
'Cuando el trabajador llega al inicio, simulo que vende sus productos durante 1 minuto
If Trabajador Then
    TIEMPO_ESPERA = 1
    If FinalCamino Then
        FinalCamino = False
    End If
    If YaGire Then
        YaGire = False
    End If
End If
End Sub

Private Sub LlegueAlFinal()
If Trabajador Then
    'Giro al personaje para que quede de cara al agua o arbol
    FinalCamino = True
    YaGire = False
    TiempoTrabajado = Now
    TIEMPO_ESPERA = GetTiempoTrabajo
    If PrimeraVuelta Then
        PrimeraVuelta = False
    End If
End If
End Sub
Private Sub EstoyCaminando()
If Trabajador Then
    'Giro al personaje para que quede de cara al inicio
    If FinalCamino Then
        FinalCamino = False
    End If
    If YaGire Then
        YaGire = False
    End If
    If TIEMPO_ESPERA > 0 Then
        TIEMPO_ESPERA = 0
    End If
End If
End Sub

Private Sub ResetMercader()
potasRojas = NUM_ROJAS
potasAzules = NUM_AZULES

Mana = MAX_MANA
'If pNpcIndex <> 0 Then
Npclist(pNpcIndex).Stats.MinHP = Npclist(pNpcIndex).Stats.MaxHP

Oro = Oro + INCREMENTO_ORO * MultiplicadorORO
If Oro > MAX_ORO * MultiplicadorORO Then Oro = MAX_ORO * MultiplicadorORO
Npclist(pNpcIndex).GiveGLDMin = Oro
Npclist(pNpcIndex).GiveGLDMax = Oro

Npclist(pNpcIndex).Invent.Object(1).Amount = potasRojas
Npclist(pNpcIndex).Invent.Object(2).Amount = potasAzules

Emergencia = False
Recalcular = False
Combear = False
Estrategia = 0

Set Agresores = New Dictionary
Set pRuta = Nothing

LastAgresor = -1
Tiempo = Now
estado = 0
'End If
End Sub

Private Function BuscarCamino() As Boolean
Set pRuta = New clsTree
BuscarCamino = pRuta.RunSearch(1, Npclist(pNpcIndex).Pos.X, Npclist(pNpcIndex).Pos.Y, DestinoX, DestinoY)
End Function


Public Sub MoverMercader()
On Error GoTo errhandler:
With Npclist(pNpcIndex)
Dim agresorIndex As Integer
Dim tmpIndex As Integer
Dim tHeading As Byte
Dim X As Integer, Y As Integer
Dim nPos As WorldPos
Dim i As Integer

    'Si tiene poca vida que priorize curarse.
    If .Stats.MinHP <= 210 Then
        Emergencia = True
    End If

    'Si tiene potas rojas y esta herido se cura solo :D
    If .Stats.MinHP < .Stats.MaxHP And (LanzoAtaque <= 3 Or CPAtaca >= 3) Then
        If potasRojas > 0 Then
            potasRojas = potasRojas - 1
            Npclist(pNpcIndex).Invent.Object(1).Amount = potasRojas
            .Stats.MinHP = .Stats.MinHP + 30
            If .Stats.MinHP > .Stats.MaxHP Then
                .Stats.MinHP = .Stats.MaxHP
            End If
            Emergencia = False
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
    End If
    'Si tiene potas azules que restaure mana
    If Mana < MAX_MANA And LanzoAtaque <= 4 Then
        If potasAzules > 0 Then
            potasAzules = potasAzules - 1
            Npclist(pNpcIndex).Invent.Object(2).Amount = potasAzules
            Mana = Mana + MAX_MANA * 0.035
            If Mana > MAX_MANA Then
                Mana = MAX_MANA
            End If
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
    End If
    
    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
        'Neo El npc tira remo?
        If TiraRemo = 1 And Trabajador = False Then
            CPRemo = CPRemo + 1
            If CPRemo > 4 And LanzoAtaque < 2 And RandomNumber(1, 3) = 1 Then
                If NpcLanzaSpellSobreNpcComoUser(pNpcIndex, pNpcIndex, HECHI_REMO, Mana) Then
                    CPRemo = 0
                    LanzoAtaque = 9
                End If
            End If
        End If
    End If

    agresorIndex = SeleccionarAgresor
    
    If agresorIndex <> LastAgresor Then
        'Si cambia de agresor le borro la ruta que tenia
        LastAgresor = agresorIndex
        CalcularInmo = False
        Set pRuta = Nothing
    End If
    If agresorIndex > 0 Then
        If estado = 1 Then
            LanzoAtaque = 12
        End If
        estado = 2
        If LanzoAtaque > 0 Then LanzoAtaque = LanzoAtaque - 1
        'Busco users atacantes que esten al lado
        If Agresores.Exists(MapData(.Pos.map).Tile(.Pos.X - 1, .Pos.Y).UserIndex) Then
            tmpIndex = MapData(.Pos.map).Tile(.Pos.X - 1, .Pos.Y).UserIndex
            tHeading = eHeading.WEST
        ElseIf Agresores.Exists(MapData(.Pos.map).Tile(.Pos.X + 1, .Pos.Y).UserIndex) Then
            tmpIndex = MapData(.Pos.map).Tile(.Pos.X + 1, .Pos.Y).UserIndex
            tHeading = eHeading.EAST
        ElseIf Agresores.Exists(MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).UserIndex) Then
            tmpIndex = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).UserIndex
            tHeading = eHeading.NORTH
        ElseIf Agresores.Exists(MapData(.Pos.map).Tile(.Pos.X, .Pos.Y + 1).UserIndex) Then
            tmpIndex = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y + 1).UserIndex
            tHeading = eHeading.SOUTH
        End If
        If .flags.Paralizado = 1 And tHeading <> .Char.heading Then tHeading = 0
        If tHeading > 0 And LanzoAtaque = 0 And Not Emergencia Then
            'Si tengo a un boludo al lado mientras estoy inmo lo ataco.
            If tHeading <> .Char.heading Then Call ChangeNPCChar(pNpcIndex, .Char.Body, .Char.Head, tHeading)
            Call NpcAtacaUser(pNpcIndex, tmpIndex)
            LanzoAtaque = 2
            Combear = True
            If RandomNumber(1, 4) <> 1 Then
                Exit Sub
            End If
        ElseIf tHeading > 0 And LanzoAtaque > 0 And Not Emergencia Then
            'Giro pero no ataco
            Call ChangeNPCChar(pNpcIndex, .Char.Body, .Char.Head, tHeading)
            Exit Sub
        Else
            If Trabajador = False Then
            'Neo Los trabajadores no usan magia
                Call AtacaConMagia(agresorIndex)
            End If
        End If
    
        If UserList(agresorIndex).flags.Inmovilizado = 1 And Not Escapar And Not Emergencia Then
            'Si lo inmovilize voy al lado
            If Not CalcularInmo Then
                Set pRuta = New clsTree
                DestinoX = UserList(agresorIndex).Pos.X
                DestinoY = UserList(agresorIndex).Pos.Y
                If pRuta.RunSearch(1, Npclist(pNpcIndex).Pos.X, Npclist(pNpcIndex).Pos.Y, DestinoX, DestinoY) Then
                    CalcularInmo = True
                End If
            ElseIf .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                If Not SeguirRuta Then
                    Escapar = True
                End If
            End If
        Else
            If Not Emergencia Then 'No quiero que deje de hacer lo que estaba haciendo, solo que tome una pocion y siga
                CalcularInmo = False
                Set pRuta = Nothing
            End If
            If UserList(agresorIndex).flags.Inmovilizado = 0 Then Escapar = False
            
            If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                Recalcular = True
                If Direccion = 0 Then
                    Direccion = .Char.heading
                ElseIf RandomNumber(1, 2) = 1 Then
                    Direccion = RandomNumber(1, 4)
                    If .Char.heading = Direccion Then Direccion = RandomNumber(1, 4)
                End If
                nPos = .Pos
                Call HeadtoPos(Direccion, nPos)
                If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                    MoveNPCChar pNpcIndex, Direccion
                Else
                    For i = 1 To 4
                        If i <> Direccion Then
                            nPos = .Pos
                            Call HeadtoPos(i, nPos)
                            If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                                MoveNPCChar pNpcIndex, i
                                Exit For
                            End If
                        End If
                    Next i
                End If
                'Neo Si no es una criatura, entonces habla
                If Criatura = 0 Then
                    CPDialogo = CPDialogo + 1
                    If CPDialogo > 25 And RandomNumber(1, 10) = 1 Then
                        CPDialogo = 0
                        'Neo Mensajes que puede tirar el bicho al azar
                        Dim frasesRandomAgresor As Integer
                        frasesRandomAgresor = 0
                        If MsjAgresorRandom1 <> "" Then
                            frasesRandomAgresor = frasesRandomAgresor + 1
                        End If
                        If MsjAgresorRandom2 <> "" Then
                            frasesRandomAgresor = frasesRandomAgresor + 1
                        End If
                        If MsjAgresorRandom3 <> "" Then
                            frasesRandomAgresor = frasesRandomAgresor + 1
                        End If
                        
                        Dim numeroRandomFrase As Integer
                        numeroRandomFrase = 0
                        numeroRandomFrase = RandomNumber(1, frasesRandomAgresor)
                        
                        'Se tira un mensaje al azar
                        If numeroRandomFrase = 1 Then
                            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(MsjAgresorRandom1, .Char.CharIndex, vbWhite))
                        Else
                            If numeroRandomFrase = 2 Then
                                Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(MsjAgresorRandom2, .Char.CharIndex, vbWhite))
                            Else
                                If numeroRandomFrase = 3 Then
                                    Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(MsjAgresorRandom3, .Char.CharIndex, vbWhite))
                                End If
                            End If
                        End If 'de numero frase
                    End If 'de dialogo > 25
                End If ' de no es criatura
            End If
        
        
        End If
    ElseIf Recalcular Then
        Recalcular = False
        BardeoInvis = False
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        If BuscarCamino Then MoverMercader
    ElseIf Not pRuta Is Nothing Then
        If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 And agresorIndex = 0 Then
            estado = 1
            'Debug.Print ("Paso: " & startX & "  " & startY & " - " & PasoActual & " - " & Camino(PasoActual).X & "  " & Camino(PasoActual).Y)
            If Not SeguirRuta Then
                If SiguientePaso Then
                    MoverMercader
                End If
            End If
            
            'Si alguien me sigue verifico cuanto viene siguiendo y si estuvo por mas de 20 pasos lo ataco.
            Dim Enemigo As Integer
            Enemigo = SeleccionarEnemigo
            If Enemigo > 0 Then
                PasosEnemigo = PasosEnemigo + 1
                If PasosEnemigo = 20 Then
                    If .Stats.Alineacion = 0 Then
                        If MsjAvisoPosibleAgresor <> "" Then
                            Call WriteChatOverHead(Enemigo, MsjAvisoPosibleAgresor, .Char.CharIndex, vbWhite)
                        End If
                    Else
                        If MsjAvisoPosibleAgresor <> "" Then
                            Call WriteChatOverHead(Enemigo, MsjAvisoPosibleAgresor, .Char.CharIndex, vbWhite)
                        End If
                    End If
                ElseIf PasosEnemigo >= 30 Then
                    AgregarAgresor (Enemigo)
                    PasosEnemigo = 22
                End If
            ElseIf PasosEnemigo > 0 Then
                PasosEnemigo = PasosEnemigo - 1
            End If
        End If
    ElseIf DateDiff("s", Tiempo, Now) >= TIEMPO_ESPERA * 60 Then
        estado = 1
        SiguientePaso
    ElseIf Trabajador = True And DateDiff("s", Tiempo, Now) < TIEMPO_ESPERA * 60 And estado <> 1 Then
        If FinalCamino = False Then
            'Estoy en el inicio
            '#### LLEGUE AL INICIO DEL CAMINO
            If YaGire = False Then
                GiroPersonaje pNpcIndex, GiroInicio, .Char.Body, .Char.Head
                YaGire = True
            End If
        Else
            '#### LLEGUE AL FINAL DEL CAMINO
            'Llegue al final de la ruta, entonces trabajo
                If YaGire = False Then
                    GiroPersonaje pNpcIndex, GiroFin, .Char.Body, .Char.Head
                    YaGire = True
                End If
                'Veo que tipo de trabajador es
                'Es hora de trabajar?
                'Solo trabajo, si NO esta lloviendo
                If Lloviendo = False Then
                    If DateDiff("s", TiempoTrabajo, Now) >= SegundosTrabajo Then
                        'Neo Segun la profesion, ejecuta el ruido de trabajo que corresponde
                        RuidoTrabajo .Pos.X, .Pos.Y
                        SegundosTrabajo = GetIntarvaloTrabajo
                        TiempoTrabajo = Now
                    End If
                End If
                
                'El npc trabajador, cada cierto tiempo random, habla
                If DateDiff("s", TiempoHablar, Now) >= CadaCuantoHabla * 60 Then
                    'Verifico que existan frases cargadas
                    If NumeroFraseTrabajador > 0 Then
                         Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(FraseTrabajadores(RandomNumber(1, NumeroFraseTrabajador)).texto, .Char.CharIndex, vbWhite))
                    End If
                    'Actualizo el tiempo de hablar
                    TiempoHablar = Now
                    CadaCuantoHabla = GetTiempoHablar
                End If 'de si le toca hablar
                
                'El npc trabajador, cada cierto tiempo random, come y bebe
                If DateDiff("s", TiempoComer, Now) >= CadaCuantoComo * 60 Then
                    'Puedo comer o beber, ejecuto un sonido random
                    ComoBebo .Pos.X, .Pos.Y
                    'Actualizo el tiempo de hablar
                    TiempoComer = Now
                    CadaCuantoComo = GetTiempoComerBeber
                End If 'de si le toca hablar
        End If 'de llegue al final del camino
        
    End If
End With

Exit Sub
errhandler:
Exit Sub
End Sub
Private Sub ComoBebo(ByVal posX As Integer, ByVal posY As Integer)
'Neo El npc come o bebe
Dim comer As Integer
comer = RandomNumber(1, 2)
If comer = 1 Then
    'Bebo
    Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_BEBER, posX, posY))
Else
    'Como
    Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_COMER, posX, posY))
End If
End Sub
Private Sub RuidoTrabajo(ByVal posX As Integer, ByVal posY As Integer)
'Neo En base a la profesion del trabajador, ejecuta el ruido del trabajo que corresponde
    If Pescador = True Then
        Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_PESCAR, posX, posY))
    ElseIf Talador = True Then
        Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_TALAR, posX, posY))
    ElseIf Carpintero = True Then
        Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(LABUROCARPINTERO, posX, posY))
    ElseIf Minero = True Then
        Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_MINERO, posX, posY))
    ElseIf Herrero = True Then
        Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(MARTILLOHERRERO, posX, posY))
    End If
End Sub
Private Sub GiroPersonaje(ByVal index, ByVal Giro As String, ByVal Body As Integer, ByVal Head As Integer)
'Neo Posiciona un personaje, segun el orden que le demos
Dim tHeading As Byte

If Giro = "NORTE" Then
    tHeading = eHeading.NORTH
ElseIf Giro = "SUR" Then
    tHeading = eHeading.SOUTH
ElseIf Giro = "ESTE" Then
    tHeading = eHeading.EAST
ElseIf Giro = "OESTE" Then
    tHeading = eHeading.WEST
End If
    
    Call ChangeNPCChar(index, Body, Head, tHeading)
End Sub

Private Function SeguirRuta() As Boolean
Dim tHeading As Byte
startX = Npclist(pNpcIndex).Pos.X
startY = Npclist(pNpcIndex).Pos.Y
If pRuta.PathStepNext(startX, startY) > 0 Then
    Dim tmpPos As WorldPos
    tmpPos.map = Npclist(pNpcIndex).Pos.map
    tmpPos.X = startX
    tmpPos.Y = startY
    tHeading = FindDirection(Npclist(pNpcIndex).Pos, tmpPos)
            
    If LegalPosNPC(tmpPos.map, tmpPos.X, tmpPos.Y, 0, True) Then
        MoveNPCChar pNpcIndex, tHeading
    ElseIf DestinoX <> tmpPos.X Or DestinoY <> tmpPos.Y Then 'Si no puedo ir al lugar que tengo que ir no importa
        If BuscarCamino Then
            MoverMercader
        End If
    Else
        SeguirRuta = False
        Exit Function
    End If
    SeguirRuta = True
Else
    SeguirRuta = False
End If
End Function

Public Sub ReSpawn()

pPaso = 0
Yendo = True
Oro = ORO_INICIAL * MultiplicadorORO
Dim Pos As WorldPos
Pos.map = 1
Pos.X = Camino(0).X
Pos.Y = Camino(0).Y
pNpcIndex = SpawnNpc(pNroNPC, Pos, True, True, 1)
'If pNpcIndex <> 0 Then
Npclist(pNpcIndex).GiveGLDMin = Oro
Npclist(pNpcIndex).GiveGLDMax = Oro
Call CheckZonaNPC(pNpcIndex)
'End If
ResetMercader
End Sub

Public Sub AgregarAgresor(ByVal UserIndex As Integer)
If Not Agresores.Exists(UserIndex) Then
    Call Agresores.Add(UserIndex, UserIndex)
    If MsjNuevoAgresor <> "" Then
        'Si no es una criatura, ni un trabajador, entonces bardea a un agresor
        If Criatura = 0 And Trabajador = False Then
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(MsjNuevoAgresor, Npclist(pNpcIndex).Char.CharIndex, vbWhite))
        End If
    End If
End If
End Sub
Public Sub QuitarAgresor(ByVal UserIndex As Integer)
If Agresores.Exists(UserIndex) Then
    Call Agresores.Remove(UserIndex)
End If
End Sub


Private Function SeleccionarAgresor() As Integer
Dim i As Integer, UserIndex As Integer
Dim Candidato As Integer, CandidatoHP As Integer
Dim FactorMovimiento As Single
Candidato = 0
CandidatoHP = 2000
For i = 0 To Agresores.Count - 1
    UserIndex = Agresores.Items(i)
    If Abs(Npclist(pNpcIndex).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y And _
       Abs(Npclist(pNpcIndex).Pos.X - UserList(UserIndex).Pos.X) <= RANGO_VISION_X And _
       UserList(UserIndex).flags.Muerto = 0 And _
       Zonas(UserList(UserIndex).zona).Segura = 0 And _
       Zonas(Npclist(pNpcIndex).zona).Segura = 0 Then
        If UserList(UserIndex).flags.invisible = 1 And LastAgresor = UserIndex And Not BardeoInvis Then
            BardeoInvis = True
            LanzoAtaque = 7
            If MsjAgresorInvi <> "" Then
                'Si no es una criatura ni un trabajador, entonces habla
                
                    Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead(MsjAgresorInvi, Npclist(pNpcIndex).Char.CharIndex, vbWhite))
                
            End If
        End If
        'Este calculo sirve para priorizar los pjs que no se estan moviendo sobre los que tienen baja vida. (Revisar)
        FactorMovimiento = (1 + (UserList(UserIndex).flags.Movimiento + 20) / 200)
        If UserList(UserIndex).Stats.MinHP * FactorMovimiento < CandidatoHP Then
            Candidato = UserIndex
            CandidatoHP = UserList(UserIndex).Stats.MinHP * FactorMovimiento
        End If
    End If
Next i
SeleccionarAgresor = Candidato
End Function

Private Function SeleccionarEnemigo() As Integer
Dim i As Integer, UserIndex As Integer
SeleccionarEnemigo = 0
For i = 0 To Npclist(pNpcIndex).AreasInfo.Users.Count - 1
    UserIndex = Npclist(pNpcIndex).AreasInfo.Users.Items(i)
    If Abs(Npclist(pNpcIndex).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y And _
       Abs(Npclist(pNpcIndex).Pos.X - UserList(UserIndex).Pos.X) <= RANGO_VISION_X And _
       UserList(UserIndex).flags.Muerto = 0 And _
       Zonas(UserList(UserIndex).zona).Segura = 0 And _
       Zonas(Npclist(pNpcIndex).zona).Segura = 0 And _
       Not Agresores.Exists(UserIndex) And ((Npclist(pNpcIndex).Stats.Alineacion = 0 And Criminal(UserIndex)) Or (Npclist(pNpcIndex).Stats.Alineacion = 1 And UserList(UserIndex).Faccion.FuerzasCaos = 0)) Then
        SeleccionarEnemigo = UserIndex
        Exit For
    End If
Next i
End Function

Private Sub AtacaConMagia(UserIndex As Integer)
'Neo Los trabajadores no pueden atacar con magia
If Trabajador = False Then
    CPAtaca = CPAtaca + 1
    
    If Emergencia Then Exit Sub
    
    Debug.Print UserList(UserIndex).flags.Movimiento & "  " & Mana & "  " & Estrategia & "  -  " & LanzoAtaque
    
    If RandomNumber(1, 100) = 1 Then
        Estrategia = Estrategia + 1
        If Estrategia = 3 Then Estrategia = 0
    End If
    
    If LanzoAtaque = 0 Then 'Un intervalo de pasos minimo por cada ataque
    
        If UserList(UserIndex).flags.invisible = 1 Then 'Si esta invisible lo unico que hago es quitarle el invi primero
            If RandomNumber(1, UserList(UserIndex).flags.Movimiento + 3) <= 2 Then
                If Mana >= Hechizos(HECHI_REMOINVI).ManaRequerido Then
                    If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_REMOINVI, Mana) And TiraVerInvi = 1 Then
                        CPAtaca = 0
                        LanzoAtaque = 8
                    End If
                End If
            End If
        Else
            If UserList(UserIndex).flags.Inmovilizado = 0 And (Estrategia = 0 Or Estrategia = 2 Or RandomNumber(1, 5) = 1) And RandomNumber(1, 10) <> 1 And Combear = False Then
                If RandomNumber(1, UserList(UserIndex).flags.Movimiento + 3) = 1 And TiraInmo = 1 Then
                    If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_INMO, Mana) Then
                        LanzoAtaque = 6
                        CPAtaca = 0
                    End If
                End If
            Else
                If RandomNumber(1, IIf(UserList(UserIndex).flags.Inmovilizado Or Combear, 2, UserList(UserIndex).flags.Movimiento + 3)) = 1 Then
                    Combear = False 'Si pego un golpe priozo tirar un hechizo
                    If Mana >= Hechizos(HECHI_APOCA).ManaRequerido And TiraApoca = 1 Then
                        If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_APOCA, Mana) Then
                            CPAtaca = 0
                            LanzoAtaque = 6
                        End If
                    ElseIf Mana >= Hechizos(HECHI_DESCARGA).ManaRequerido And TiraDescarga = 1 Then 'Si no tiene mana para apoca prueba con descarga
                        If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_DESCARGA, Mana) Then
                            CPAtaca = 0
                            LanzoAtaque = 5
                        End If
                    ElseIf Mana >= Hechizos(HECHI_TORMENTA_FUEGO).ManaRequerido And TiraTormentaFuego = 1 Then 'Si no tiene mana para descarga prueba con tormenta
                        If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_TORMENTA_FUEGO, Mana) Then
                            CPAtaca = 0
                            LanzoAtaque = 4
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

End Sub
Public Sub Clicked(ByVal UserIndex As Integer)
Dim Status As String
Dim ft As FontTypeNames
'Cuando le hacen click, muestro el nick y si tiene clan o faccion, tambien se muestra
Dim nick As String
Dim FoundIt As Integer

FoundIt = InStr(1, Npclist(pNpcIndex).Name, "!")

If FoundIt <> 0 Then
    nick = Replace(Npclist(pNpcIndex).Name, "!", "")
Else
    nick = Npclist(pNpcIndex).Name
End If

If Npclist(pNpcIndex).Stats.Alineacion = 0 Then
    If Clan <> "" Then
        Status = "Ves a " & nick & " " & Clan
    Else
        Status = "Ves a " & nick
    End If
    ft = FontTypeNames.FONTTYPE_CONSEJOVesA
Else
    If Clan <> "" Then
        Status = "Ves a " & nick & " " & Clan
    Else
        Status = "Ves a " & nick
    End If
    ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
End If

Call WriteConsoleMsg(UserIndex, Status, ft)

'Si tiene seteado Destino1 o Destino2, informo donde voy
'Solo el mercader real y el del caos informan donde van
If pNroNPC = 617 Or pNroNPC = 618 Then
    If Destino(1) <> "" And Destino(2) <> "" Then
        If Npclist(pNpcIndex).Stats.Alineacion = 0 Then
            If estado = 0 Then
                Call WriteChatOverHead(UserIndex, "Cuando termine de negociar partiré hacia " & Destino(IIf(Yendo, 2, 1)) & ".", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
            ElseIf estado = 1 Then
                Call WriteChatOverHead(UserIndex, "No tengo tiempo de charlar debo entregar estas mercancías.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
            End If
        Else
            If estado = 0 Then
                Call WriteChatOverHead(UserIndex, "Cuando termine aquí iré hacia " & Destino(IIf(Yendo, 2, 1)) & ".", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
            ElseIf estado = 1 Then
                Call WriteChatOverHead(UserIndex, "No me estorbes patán o sufrirás mi ira.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
            End If
        End If
    End If
Else
    If estado = 0 Then
        If MsjAlHacerClick <> "" Then
            Call WriteChatOverHead(UserIndex, MsjAlHacerClick, Npclist(pNpcIndex).Char.CharIndex, vbWhite)
        End If
    ElseIf estado = 1 Then
        If MsjNoMolestes <> "" Then
            Call WriteChatOverHead(UserIndex, MsjNoMolestes, Npclist(pNpcIndex).Char.CharIndex, vbWhite)
        End If
    End If
End If


End Sub
Private Function GetTiempoTrabajo() As Integer
    'Neo Retorna un nro random de trabajo
    GetTiempoTrabajo = RandomNumber(TiempoMinimoTrabajo, TiempoMaximoTrabajo)
End Function
Private Function GetTiempoHablar() As Integer
    'Neo Retorna un nro random para hablar
    GetTiempoHablar = RandomNumber(TiempoMinimoHablar, TiempoMaximoHablar)
End Function
Private Function GetTiempoLogin() As Integer
    'Neo Retorna un nro random para el login del npc
    GetTiempoLogin = RandomNumber(TiempoMinimoLogin, TiempoMaximoLogin)
End Function
Private Function GetTiempoResucitar() As Integer
    'Neo Retorna un nro random para el login del npc
    GetTiempoResucitar = RandomNumber(TiempoMinimoResucitar, TiempoMaximoResucitar)
End Function
Private Function GetTiempoComerBeber() As Integer
    'Neo Retorna un nro random para comer, beber del npc
    GetTiempoComerBeber = RandomNumber(TiempoMinimoComerBeber, TiempoMaximoComerBeber)
End Function
Private Function GetIntarvaloTrabajo() As Integer
    'Neo Retorna un nro random para el intervalo de trabajo
    GetIntarvaloTrabajo = RandomNumber(IntervaloMinimoTrabajo, IntervaloMaximoTrabajo)
End Function
