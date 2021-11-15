Attribute VB_Name = "modMercader"
Option Explicit
Private Camino() As Position
Private Const CANT_PASOS As Byte = 23
Private PasoActual As Byte
Private NpcMercader As Integer
Private Const MERCADER_NPC As Integer = 617
Private RutaMercader As clsTree
Private MercaderLlega As Boolean
Private startX As Integer, startY As Integer
Private posicionNpcI As Integer
Private TotalNpcIA As Integer


'Neo
'Public MercaderNpc As clsMercader

'Public MercaderReal As clsMercader
'Public MercaderCaos As clsMercader

Public Sub initMercader()

'Cuando se arranca el servidor, se carga la lista de npc inteligentes, por lo tanto, si la tiene tiene datos,
'entonces la recorro y voy creando npc utilizando la clase clsMercader, en base a su alineacion y todas las variables definidas.
'Se respetan los mercaderes originales (armada real y kaos), y se agregan todos los npcs inteligentes definidos en NpcInteligente.dat
'NumeroNpcInteligente me indica el total de npc inteligentes cargados
'Me fijo si al menos hay un npc inteligente cargado
'Se cargan los npc peleadores
'Creo la dimension en base a los peleadores y trabajadores
TotalNpcIA = 0
If NumeroNpcInteligente > 0 Or NumeroNpcTrabajador > 0 Then
    TotalNpcIA = NumeroNpcInteligente + NumeroNpcTrabajador
    ReDim ListaNpcIA(1 To TotalNpcIA) As clsMercader
    
    Dim ContadorGlobal As Integer
    Dim ContadorLocal As Integer
    Dim mercaderNpc As New clsMercader
    'Seteo el contador global en 1
    ContadorGlobal = 1
    'Cargo los luchadores
    If (NumeroNpcInteligente > 0) Then
        For ContadorLocal = 1 To NumeroNpcInteligente
            'Seteo el mercader
            Set mercaderNpc = New clsMercader
            'Creo el nuevo mercader con todos sus datos
            'mercaderNpc.Init NpcInteligentes(Contador).NumeroNPC, NpcInteligentes(Contador).Mana, NpcInteligentes(Contador).OroInicial, NpcInteligentes(Contador).OroMaximo, NpcInteligentes(Contador).IncrementoOro, NpcInteligentes(Contador).TiempoEsperaMinuto, NpcInteligentes(Contador).PocionesRojas, NpcInteligentes(Contador).PocionesAzules, NpcInteligentes(Contador).Ruta, NpcInteligentes(Contador).Destino1, NpcInteligentes(Contador).Destino2
            mercaderNpc.Init ContadorLocal, False
            'Agrego el mercader a la lista
            Set ListaNpcIA(ContadorGlobal) = mercaderNpc
            'Aumento el contador global
            ContadorGlobal = ContadorGlobal + 1
        Next ContadorLocal
    End If
    'Cargo los trabajadores
    If (NumeroNpcTrabajador > 0) Then
        For ContadorLocal = 1 To NumeroNpcTrabajador
            'Seteo el mercader
            Set mercaderNpc = New clsMercader
            'Creo el nuevo mercader con todos sus datos
            mercaderNpc.Init ContadorLocal, True
            'Agrego el mercader a la lista
            Set ListaNpcIA(ContadorGlobal) = mercaderNpc
            'Aumento el contador global
            ContadorGlobal = ContadorGlobal + 1
        Next ContadorLocal
    End If
End If

'If (ListaMercader.Count > 0) Then
'   Dim thing As clsMercader '// this is the key
'
'   For Each thing In ListaMercader
 '     Call MsgBox("Se agrego el npc" + CStr(thing.NpcNum), vbCritical + vbOKOnly)
'   Next
'End If

'Call CargarBotDat
'Recorro la lista de bots, por cada bot, lo creo, en base a su alineacion y sus variables

'Set MercaderReal = New clsMercader

'Call MercaderReal.Init(617, "297,859;285,863;285,845;277,838;277,802;285,794;293,793;292,687;296,683;296,617;300,615;301,557;305,556;305,520;317,519;315,492;305,491;305,346;296,344;296,282;292,277;292,222", "Ullathorpe", "Banderbill")

'Set MercaderCaos = New clsMercader

'Call MercaderCaos.Init(618, "195,1225;195,1235;201,1235;201,1251;217,1252;220,1261;272,1261;272,1268;304,1269;305,1296;332,1297;333,1308;424,1309;425,1332;450,1343;456,1343;456,1396;459,1404;459,1430;536,1431;536,1423;540,1420;540,1403;550,1397;552,1392", "Nix", "Arkhein")

End Sub

Public Sub ReSpawnMercader(ByVal NPC As Integer)
'NPC es el numero de npc, ejemplo 617 que representa al mercader Real
'Busco en la lista, si el numero de npc que me pasan como parametro, es un mercader, en ese caso, hago un ReSpawn
posicionNpcI = GetNpcIAByNumNpc(NPC)
If posicionNpcI > 0 Then
    ListaNpcIA(posicionNpcI).ReSpawn
    posicionNpcI = 0
End If

'If NPC = MercaderReal.NpcNum Then
 '   MercaderReal.ReSpawn
'ElseIf NPC = MercaderCaos.NpcNum Then
 '   MercaderCaos.ReSpawn
'End If
End Sub
Public Sub MoverMercader(ByVal NpcIndex As Integer)
'Busco al npc inteligente en particular y lo hago mover
posicionNpcI = GetNpcIA(NpcIndex)
If posicionNpcI > 0 Then
    ListaNpcIA(posicionNpcI).MoverMercader
    posicionNpcI = 0
End If

'Call MercaderByIndex(NpcIndex).MoverMercader
End Sub
Public Sub MercaderAtacado(NpcIndex As Integer, ByVal UserIndex As Integer)
'Agrego un agresor al npc inteligente
posicionNpcI = GetNpcIA(NpcIndex)
If posicionNpcI > 0 Then
    ListaNpcIA(posicionNpcI).AgregarAgresor (UserIndex)
    posicionNpcI = 0
End If

'Call MercaderByIndex(NpcIndex).AgregarAgresor(UserIndex)
End Sub
Public Sub MercaderClicked(byvalNpcIndex As Integer, ByVal UserIndex As Integer)
'Hicieron click al mercader
posicionNpcI = GetNpcIA(byvalNpcIndex)
If posicionNpcI > 0 Then
    ListaNpcIA(posicionNpcI).Clicked (UserIndex)
    posicionNpcI = 0
End If

'Call MercaderByIndex(byvalNpcIndex).Clicked(UserIndex)
End Sub
Public Sub QuitarAgresorMercader(ByVal UserIndex As Integer)
'Neo Cuando un usuario muere o se va del juego, se quita como agreso de cualquier npc inteligente que ataco
If (TotalNpcIA > 0) Then
        Dim Contador As Integer
        For Contador = 1 To TotalNpcIA
            ListaNpcIA(Contador).QuitarAgresor (UserIndex)
        Next Contador
End If

'MercaderReal.QuitarAgresor (UserIndex)
'MercaderCaos.QuitarAgresor (UserIndex)
End Sub
Public Function MercaderByIndex(ByVal NpcIndex As Integer) As clsMercader
'If NpcIndex = MercaderReal.NpcIndex Then
'    Set MercaderByIndex = MercaderReal
'ElseIf NpcIndex = MercaderCaos.NpcIndex Then
'    Set MercaderByIndex = MercaderCaos
'Else
'    Set MercaderByIndex = Nothing
'End If
End Function

Public Function EsMercader(ByVal NpcIndex As Integer, ByVal Bueno As Boolean)
'Indica si el ID de npc pasado como parametro, corresponde a un mercader
Dim Resultado As Boolean

Resultado = False

posicionNpcI = GetNpcIA(NpcIndex)
If posicionNpcI > 0 Then
    Resultado = Npclist(NpcIndex).Stats.Alineacion = IIf(Bueno, 0, 1)
    posicionNpcI = 0
End If

EsMercader = Resultado

'If MercaderByIndex(NpcIndex) Is Nothing Then
'    EsMercader = False
'Else
    'EsMercader = Npclist(NpcIndex).Stats.Alineacion = IIf(Bueno, 0, 1)
'End If
End Function
'Neo En base al index del npc, retorna el npc inteligente, en caso de encontrarlo. Caso contrario, retorna nothing
Public Function GetNpcIA(ByVal NpcIndex As Integer) As Integer
    Dim Resultado As Integer
    Resultado = 0
    If (TotalNpcIA > 0) Then
        Dim Contador As Integer
        For Contador = 1 To TotalNpcIA
            If ListaNpcIA(Contador).NpcIndex = NpcIndex Then
                Resultado = Contador
                Contador = TotalNpcIA
            End If
        Next Contador
    End If
    GetNpcIA = Resultado
End Function
'Neo En base al numero del npc, retorna el npc inteligente, en caso de encontrarlo. Caso contrario, retorna nothing
Public Function GetNpcIAByNumNpc(ByVal NumNpc As Integer) As Integer
    Dim Resultado As Integer
    Resultado = 0
    If (TotalNpcIA > 0) Then
        Dim Contador As Integer
        For Contador = 1 To TotalNpcIA
            If ListaNpcIA(Contador).NpcNum = NumNpc Then
                Resultado = Contador
                Contador = TotalNpcIA
            End If
        Next Contador
    End If
    GetNpcIAByNumNpc = Resultado
End Function
