Attribute VB_Name = "modBarcos"
Option Explicit

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Const TIEMPO_EN_PUERTO As Long = 15

Public Const NUM_PUERTOS As Byte = 6
Public Const NUM_BARCOS As Byte = 10

Public Const PUERTO_NIX As Byte = 1
Public Const PUERTO_ULLATHORPE As Byte = 2
Public Const PUERTO_BANDER As Byte = 3
Public Const PUERTO_ARGHAL As Byte = 4
Public Const PUERTO_LINDOS As Byte = 5
Public Const PUERTO_ARKHEIN As Byte = 6

Public Const VALOR_BILLETE As Integer = 500


Public Type tPuerto
    Paso(0 To 1) As Byte
    Nombre As String
End Type

Public Type tPositions
    Ruta() As Position
End Type

Public Puertos(1 To NUM_PUERTOS) As tPuerto

Public Barcos(1 To NUM_BARCOS) As clsBarco

Dim RutaBarco(0 To 1) As tPositions

Dim InicioBarcos(0 To 1) As Byte
Dim FinBarcos(0 To 1) As Byte

Dim end_time As Currency
Dim timer_freq As Currency

Dim ttt As Long


'ANCLA_EMBARCACIONES
Public Const EGaleon As Integer = 907
Public Const EGalera As Integer = 907
Public Const EFragata As Integer = 907

Public Sub InitBarcos()
Dim sRutaHoraria As String
Dim sRutaAntihoraria As String
Dim i As Integer

Puertos(PUERTO_NIX).Nombre = "Nix"
Puertos(PUERTO_NIX).Paso(0) = 0
Puertos(PUERTO_NIX).Paso(1) = 24

Puertos(PUERTO_ULLATHORPE).Nombre = "Ullathorpe"
Puertos(PUERTO_ULLATHORPE).Paso(0) = 3
Puertos(PUERTO_ULLATHORPE).Paso(1) = 20

Puertos(PUERTO_BANDER).Nombre = "Banderbill"
Puertos(PUERTO_BANDER).Paso(0) = 7
Puertos(PUERTO_BANDER).Paso(1) = 16

Puertos(PUERTO_ARGHAL).Nombre = "Arghal"
Puertos(PUERTO_ARGHAL).Paso(0) = 14
Puertos(PUERTO_ARGHAL).Paso(1) = 9

Puertos(PUERTO_LINDOS).Nombre = "Lindos"
Puertos(PUERTO_LINDOS).Paso(0) = 17
Puertos(PUERTO_LINDOS).Paso(1) = 5

Puertos(PUERTO_ARKHEIN).Nombre = "Arkhein"
Puertos(PUERTO_ARKHEIN).Paso(0) = 23
Puertos(PUERTO_ARKHEIN).Paso(1) = 0



'sRutaHoraria = "161,1248;36,1248;36,22;302,22;302,55;303,55;566,55;566,65;635,65;635,54;800,54;800,307;801,307;870,307;870,999;887,999;870,999;870,1224;645,1224;645,1371;645,1472;196,1472;196,1266;169,1266;169,1248"
'sRutaAntihoraria = "642,1384;649,1384;649,1228;875,1228;875,995;887,995;875,995;875,303;806,303;806,313;806,50;631,50;631,61;570,61;570,51;308,51;308,62;308,18;31,18;31,1251;167,1251;167,1243;167,1270;191,1270;191,1476;648,1476;648,1384"

sRutaHoraria = "161,1248;36,1248;36,897;171,897;36,897;36,22;302,22;302,55;566,55;566,65;635,65;635,54;800,54;800,307;801,307;870,307;870,999;887,999;870,999;870,1224;644,1224;644,1361;645,1361;645,1371;645,1385;644,1385;644,1472;196,1472;196,1266;169,1266;169,1248"
sRutaAntihoraria = "642,1384;649,1384;649,1228;875,1228;875,995;887,995;875,995;875,303;806,303;806,313;806,50;631,50;631,61;570,61;570,51;308,51;308,62;308,18;31,18;31,893;171,893;31,893;31,1251;167,1251;167,1243;167,1270;191,1270;191,1476;648,1476;648,1384"


Dim Rutas() As String
Dim UPasos As Integer

Rutas = Split(sRutaHoraria, ";")
UPasos = UBound(Rutas)
ReDim RutaBarco(0).Ruta(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(0).Ruta(i).X = val(ReadField(1, Rutas(i), 44))
    RutaBarco(0).Ruta(i).Y = val(ReadField(2, Rutas(i), 44))
Next i

Rutas = Split(sRutaAntihoraria, ";")
UPasos = UBound(Rutas)
ReDim RutaBarco(1).Ruta(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(1).Ruta(i).X = val(ReadField(1, Rutas(i), 44))
    RutaBarco(1).Ruta(i).Y = val(ReadField(2, Rutas(i), 44))
Next i

For i = 1 To NUM_BARCOS
    Set Barcos(i) = New clsBarco
Next i

Call Barcos(1).Init(sRutaHoraria, 1, 161, 1248, 1, 0, 1) 'NIX
Call Barcos(2).Init(sRutaHoraria, 6, 302, 24, 0, 0, 2) 'LLEGANDO BANDER
Call Barcos(3).Init(sRutaHoraria, 12, 800, 88, 0, 0, 3) 'LLEGANDO  ARGHAL
Call Barcos(4).Init(sRutaHoraria, 16, 875, 999, 0, 0, 4) ' LLEGANDO LINDOS
Call Barcos(5).Init(sRutaHoraria, 22, 645, 1365, 0, 0, 5) ' LLEGANDO ARKHEIM

Call Barcos(6).Init(sRutaAntihoraria, 1, 642, 1384, 1, 1, 6)
Call Barcos(7).Init(sRutaAntihoraria, 4, 879, 995, 0, 1, 7)
Call Barcos(8).Init(sRutaAntihoraria, 7, 860, 303, 0, 1, 8)
Call Barcos(9).Init(sRutaAntihoraria, 16, 308, 60, 0, 1, 9)
Call Barcos(10).Init(sRutaAntihoraria, 23, 167, 1250, 0, 1, 10)

'Call Barcos(1).Init(sRutaHoraria, 1, 161, 1248, 1, 0, 1) 'NIX
'Call Barcos(2).Init(sRutaHoraria, 2, 35, 352, 0, 0, 2)
'Call Barcos(3).Init(sRutaHoraria, 8, 580, 65, 0, 0, 3)
'Call Barcos(4).Init(sRutaHoraria, 14, 870, 638, 0, 0, 4)
'Call Barcos(5).Init(sRutaHoraria, 19, 645, 1343, 0, 0, 5)

'Call Barcos(6).Init(sRutaAntihoraria, 1, 642, 1384, 1, 1, 6)
'Call Barcos(7).Init(sRutaAntihoraria, 6, 887, 995, 9703, 1, 7)
'Call Barcos(8).Init(sRutaAntihoraria, 10, 806, 313, 14742, 1, 8)
'Call Barcos(9).Init(sRutaAntihoraria, 17, 308, 62, 19671, 1, 9) 'NIX
'Call Barcos(10).Init(sRutaAntihoraria, 20, 140, 1251, 0, 1, 10)



InicioBarcos(0) = 1
InicioBarcos(1) = 6
FinBarcos(0) = 5
FinBarcos(1) = 10

Call QueryPerformanceCounter(end_time)

ttt = (GetTickCount() And &H7FFFFFFF)
End Sub


Public Sub CalcularBarcos()

Dim i As Integer
Dim ElapsedTime As Single
If Barcos(1) Is Nothing Then Exit Sub
'DoEvents
'frmMain.Show
'frmMain.pBarcos.Cls

Dim start_time As Currency

    'Get the timer frequency
If timer_freq = 0 Then
    QueryPerformanceFrequency timer_freq
End If
    
    'Get current time
Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
ElapsedTime = (start_time - end_time) / timer_freq * 1000
'Get next end time
Call QueryPerformanceCounter(end_time)

For i = 1 To NUM_BARCOS
    Call Barcos(i).Calcular(ElapsedTime)
Next i

End Sub

Private Function DistanciaPasos(ByVal Paso1 As Byte, ByVal Paso2 As Integer, ByVal CantPasos As Integer) As Integer
If Paso1 >= Paso2 Then
    DistanciaPasos = Paso1 - Paso2
Else
    DistanciaPasos = Paso1 - Paso2 + CantPasos
End If
End Function

Public Sub HablaMarinero(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal Accion As Boolean = False)


Dim i As Integer
Dim NP As Integer
Dim X As Integer
Dim Y As Integer
Dim mPaso As Integer
Dim mIdPaso As Byte
Dim Puerto As Integer
Dim Sentido As Byte

Sentido = Npclist(NpcIndex).Stats.Alineacion

For i = 1 To NUM_PUERTOS
    X = 1
    If Abs(RutaBarco(Sentido).Ruta(Puertos(i).Paso(Sentido)).X - Npclist(NpcIndex).Pos.X) < 10 And Abs(RutaBarco(Sentido).Ruta(Puertos(i).Paso(Sentido)).Y - Npclist(NpcIndex).Pos.Y) < 10 Then
    
        Exit For
    End If
Next i
Puerto = i
If Sentido = 0 Then
    NP = i + 1
    If NP > NUM_PUERTOS Then NP = 1
Else
    NP = i - 1
    If NP < 1 Then NP = NUM_PUERTOS
End If
Dim Tiempo As Integer
mPaso = 1000
For i = InicioBarcos(Sentido) To FinBarcos(Sentido)
    If Barcos(i).Paso = Puertos(Puerto).Paso(Sentido) + 1 And Barcos(i).TickPuerto > 0 Then
        Tiempo = TIEMPO_EN_PUERTO - ((GetTickCount() And &H7FFFFFFF) - Barcos(i).TickPuerto) / 1000
        Exit For
    ElseIf DistanciaPasos(Puertos(Puerto).Paso(Sentido), Barcos(i).Paso, Barcos(i).UPasos) < mPaso Then
        mPaso = DistanciaPasos(Puertos(Puerto).Paso(Sentido), Barcos(i).Paso, Barcos(i).UPasos)
        mIdPaso = i
    End If
Next i

If Not Accion Then
    If Tiempo > 0 Then
        Call WriteChatOverHead(UserIndex, "El barco zarpará hacia el puerto de " & Puertos(NP).Nombre & " en " & Tiempo & " segundos.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "El proximo barco con destino a " & Puertos(NP).Nombre & " llegará en " & Barcos(mIdPaso).EstimarTiempo & " segundos.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    End If
Else

    If i > 10 Then
        Call WriteChatOverHead(UserIndex, "Aguarda unos instantes y te atenderé!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Embarcado > 0 Then
        Call Barcos(i).QuitarPasajero(UserIndex)
    ElseIf Tiempo > 0 Then
    
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
            Call WriteConsoleMsg(UserIndex, "Debes ponerte al lado del marinero para subir al barco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Equitando = True Then
            Call WriteConsoleMsg(UserIndex, "¡Debes bajar de la montura para subirte al Barco!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Chiquito = True Then
            Call WriteConsoleMsg(UserIndex, "Tu apariencia actual no te permite subir a la embarcación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Counters.Congelado > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás congelado!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco si estás navegando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Descansar = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras estás descansando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Meditando = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras estés meditando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco estando invisible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Inmovilizado = 1 Or UserList(UserIndex).flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás paralizado!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Comerciando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras comercias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD < VALOR_BILLETE Then
            Call WriteChatOverHead(UserIndex, "Lo lamento, pero el billete vale " & VALOR_BILLETE & " monedas de oro.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub
        End If

        If Not Barcos(i).AgregarPasajero(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "Lo lamento, el barco ya está completo, deberás esperar al próximo.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        End If
    End If
End If

End Sub

Public Function BarcoEn(ByVal X As Integer, ByVal Y As Integer) As clsBarco
Dim i As Byte
For i = 1 To NUM_BARCOS
    If Barcos(i).X = X And Barcos(i).Y = Y Then
        Set BarcoEn = Barcos(i)
        Exit Function
    End If
Next i
Set BarcoEn = Nothing
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'**************************************************************************************'
'**************************************************************************************'

Public Sub Anclar_Embarcacion(ByVal UserIndex As Integer)

        With UserList(UserIndex)
            If .flags.Nadando = True Then Exit Sub
            
            If .flags.Embarcado > 0 Then
                Call Barcos(.flags.Embarcado).PasajeroJumpWater(UserIndex)
                .flags.Navegando = 1
                Call WriteNavigateToggle(UserIndex)
                Call ToggleBoatBody(UserIndex)
            End If
            
            Call DesequiparTodosLosItems(UserIndex)
                       
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .flags.Desnudo = 0
            .flags.Nadando = True
            
            Dim EmbarcacionIndex As Integer
            Dim tPos As WorldPos
            Dim tHeading As Byte
            Dim origHeading As Byte
            
            
            tPos = .Pos
            origHeading = .Char.heading
                   
            If .Char.heading = EAST Or .Char.heading = WEST Then
                tHeading = SOUTH
                'tPos.Y = tPos.Y - 1
                If Not LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False) Then
                    Call WriteConsoleMsg(UserIndex, "No hay espacio para saltar al agua!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call WarpUserChar(UserIndex, .Pos.map, .Pos.X, .Pos.Y + 1, False)
            Else
                tHeading = EAST
                'tPos.x = tPos.x + 1
                If Not LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) Then
                    Call WriteConsoleMsg(UserIndex, "No hay espacio para saltar al agua!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call WarpUserChar(UserIndex, .Pos.map, .Pos.X + 1, .Pos.Y, False)
            End If
            
            
            EmbarcacionIndex = SpawnNpc(EGaleon, tPos, False, False, 1)
            Npclist(EmbarcacionIndex).Char.Body = .Char.Body
            Npclist(EmbarcacionIndex).Char.heading = origHeading
            Npclist(EmbarcacionIndex).Name = "Embarcación de " & .Name
            .flags.EmbarcacionIndex = EmbarcacionIndex
            .flags.EmbarcacionPos = Npclist(EmbarcacionIndex).Pos
            .Char.Head = .OrigChar.Head
            
               Select Case UserList(UserIndex).genero
                   Case eGenero.Hombre
                       Select Case UserList(UserIndex).raza
                           Case eRaza.Humano
                               .Char.Body = 514
                           Case eRaza.Drow
                               .Char.Body = 539
                           Case eRaza.Elfo
                               .Char.Body = 524
                           Case eRaza.Gnomo
                               .Char.Body = 521
                           Case eRaza.Enano
                               .Char.Body = 523
                       End Select
                   Case eGenero.Mujer
                       Select Case UserList(UserIndex).raza
                           Case eRaza.Humano
                               .Char.Body = 517
                           Case eRaza.Drow
                               .Char.Body = 538
                           Case eRaza.Elfo
                               .Char.Body = 520
                           Case eRaza.Gnomo
                               .Char.Body = 519
                           Case eRaza.Enano
                               .Char.Body = 518
                       End Select
               End Select
       
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(14, .Pos.X, .Pos.Y))
        
        Call WriteConsoleMsg(UserIndex, "Hás saltado al agua!!", FontTypeNames.FONTTYPE_INFO)
        
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, tHeading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.alaIndex)
        Call ChangeNPCChar(EmbarcacionIndex, Npclist(EmbarcacionIndex).Char.Body, 0, Npclist(EmbarcacionIndex).Char.heading)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageNadando(.Char.CharIndex, True))
    
    
    End With
    
End Sub

Public Sub DesAnclar_Embarcacion(ByVal UserIndex As Integer)
    
        With UserList(UserIndex)
               
        If .flags.EmbarcacionIndex = 0 Then Exit Sub
        
        If Distancia(Npclist(.flags.EmbarcacionIndex).Pos, UserList(UserIndex).Pos) > 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás muy lejos de la embarcación para subir!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim miEmbarcacion As NPC
        miEmbarcacion = Npclist(.flags.EmbarcacionIndex)
       
        .flags.Nadando = False
        
        Dim tHeading As Byte
        tHeading = Npclist(.flags.EmbarcacionIndex).Char.heading
      
        Call QuitarNPC(.flags.EmbarcacionIndex)
        
        Call ToggleBoatBody(UserIndex)
        
        Call WarpUserChar(UserIndex, miEmbarcacion.Pos.map, miEmbarcacion.Pos.X, miEmbarcacion.Pos.Y, False)
        
        .Char.heading = tHeading
        
        .Char.Head = 0
        
        .flags.EmbarcacionIndex = 0
        
        .flags.EmbarcacionPos = .Pos
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(14, .Pos.X, .Pos.Y))
        
        Call WriteConsoleMsg(UserIndex, "Hás subido a la embarcación!!", FontTypeNames.FONTTYPE_INFO)
    
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, tHeading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.alaIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageNadando(.Char.CharIndex, False))
    
    
    End With
    
End Sub

Public Sub Muere_Embarcacion(ByVal UserIndex As Integer, Optional ByVal UserNotDie As Boolean = True)
     With UserList(UserIndex)
               
        If .flags.EmbarcacionIndex = 0 Then Exit Sub
        
        Dim miEmbarcacion As NPC
        miEmbarcacion = Npclist(.flags.EmbarcacionIndex)
       
        .flags.Nadando = False
        
        'Dim tHeading As Byte
        'tHeading = Npclist(.flags.EmbarcacionIndex).Char.heading
      
        Call QuitarNPC(.flags.EmbarcacionIndex)
        
'        If ToogleBoatBody = True Then
'            Call ToggleBoatBody(UserIndex)
'        End If
        
        'Call WarpUserChar(UserIndex, miEmbarcacion.pos.map, miEmbarcacion.pos.x, miEmbarcacion.pos.Y, False)
        
        '.Char.heading = tHeading
        
        .flags.EmbarcacionIndex = 0
        
      
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageNadando(.Char.CharIndex, False))
        
        If UserNotDie = True Then
            .Char.Head = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(14, .Pos.X, .Pos.Y))
            Call WriteConsoleMsg(UserIndex, "Te hás muerto ahogado en las profundidades del mar!!", FontTypeNames.FONTTYPE_INFO)
            Call UserDie(UserIndex)
        End If
        
    End With
End Sub


Public Sub Cerrar_Embarcacion(ByVal UserIndex As Integer)
    With UserList(UserIndex)
               
        If .flags.EmbarcacionIndex = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.EmbarcacionIndex)
         
    End With

End Sub

Public Sub Reset_Embarcacion(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
               
        If .flags.EmbarcacionIndex = 0 Then Exit Sub
        
        Dim miEmbarcacion As NPC
        miEmbarcacion = Npclist(.flags.EmbarcacionIndex)
       
        .flags.Nadando = False
        
        Dim tHeading As Byte
        tHeading = Npclist(.flags.EmbarcacionIndex).Char.heading
      
        Call QuitarNPC(.flags.EmbarcacionIndex)
    
    End With
    
End Sub


