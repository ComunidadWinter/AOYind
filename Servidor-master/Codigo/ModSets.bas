Attribute VB_Name = "ModSets"
Option Explicit
 
Public Type tSets
    ArmaduraSet As String
    ArmaSet As String
    EscudoSet As String
    CascoSet As String
    AnilloSet As String
    Efecto As Byte
End Type
 
Public Enum eEfectoSet
    PegaDoble = 1
    PegaTripe = 2
    MasAgilidad = 3
    MasFuerza = 4
    MuchaMasAgilidad = 5
    MuchaMasFuerza = 6
    MasVida = 7
End Enum
 Public SetsAura As Integer
Public Sets() As tSets
''
'Cargamos las sets al vector Sets()
Public Sub CargarSets()
    
    Dim archivoN As String
    Dim numSets As Byte
    Dim i As Byte
    
    archivoN = DatPath & "\SETS.DAT"
    
    If Not FileExist(archivoN, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar el archivo SETS.DAT.", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    numSets = val(GetVar(archivoN, "NUMEROSETS", "CantidadSets"))
 
    ReDim Sets(1 To numSets) As tSets
    For i = 1 To numSets
        Sets(i).ArmaduraSet = val(GetVar(archivoN, "SET" & i, "Armadura"))
        Sets(i).ArmaSet = val(GetVar(archivoN, "SET" & i, "Arma"))
        Sets(i).EscudoSet = val(GetVar(archivoN, "SET" & i, "Escudo"))
        Sets(i).CascoSet = val(GetVar(archivoN, "SET" & i, "Casco"))
        Sets(i).AnilloSet = val(GetVar(archivoN, "SET" & i, "Anillo"))
        Sets(i).Efecto = val(GetVar(archivoN, "SET" & i, "Efecto"))
    Next i
End Sub

''
' Devuelve el numero de set que tiene, y si no lo tiene,devuelve 0
Public Function TieneSet(ByVal UserIndex As Integer) As Byte

    Dim nSet As Byte
    'CargarSets (UserIndex)
    With UserList(UserIndex).Invent

        For nSet = 1 To UBound(Sets)
            'Tiene el casco del set?

            If .CascoEqpObjIndex = Sets(nSet).CascoSet Then
                'Tiene el escudo del set?
                If .EscudoEqpObjIndex = Sets(nSet).EscudoSet Then
                    'Tiene la armadura del set?
                    If .ArmourEqpObjIndex = Sets(nSet).ArmaduraSet Then
                        'Tiene el anillo
                        If .AnilloEqpObjIndex = Sets(nSet).AnilloSet Then
                            'Tiene el arma del set?
                            If .WeaponEqpObjIndex = Sets(nSet).ArmaSet Then
                                'Tiene el set completo :D, devolvemos el numero de set
                                TieneSet = Sets(nSet).Efecto

                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next nSet
    End With
End Function
 
''
' Aplicamos el efecto de cada set
Sub AgregarEfecto(ByVal UserIndex As Integer)
 
UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
 
  Select Case UserList(UserIndex).Stats.TengoSet
    Case eEfectoSet.PegaDoble
            Call WriteConsoleMsg(UserIndex, "Te has equipado el set de doble fuerza, ahora tu ataque se multiplicara por dos!!", FontTypeNames.FONTTYPE_FIGHT)
            UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + UserList(UserIndex).Stats.MinHIT
            UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT + UserList(UserIndex).Stats.MaxHIT
            Call WriteUpdateUserStats(UserIndex)
        Exit Sub
        
    Case eEfectoSet.PegaTripe
     Call WriteConsoleMsg(UserIndex, "Te has equipado el set de doble fuerza, ahora tu ataque se multiplicara por tres!!", FontTypeNames.FONTTYPE_FIGHT)
           UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + UserList(UserIndex).Stats.MinHIT + UserList(UserIndex).Stats.MinHIT
            UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT + UserList(UserIndex).Stats.MaxHIT + UserList(UserIndex).Stats.MaxHIT
        Call WriteUpdateUserStats(UserIndex)
        Exit Sub
        
    Case eEfectoSet.MasAgilidad
     Call WriteConsoleMsg(UserIndex, "Te has equipado el set de agilidad!!", FontTypeNames.FONTTYPE_FIGHT)
     
            UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) + 4
            Call WriteAttributes(UserIndex, True)
        Exit Sub
        
    Case eEfectoSet.MasFuerza
    Call WriteConsoleMsg(UserIndex, "Te has equipado el set de fuerza!!", FontTypeNames.FONTTYPE_FIGHT)

            UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) + 4
          Call WriteAttributes(UserIndex, True)
        Exit Sub
        
    Case eEfectoSet.MuchaMasAgilidad
    Call WriteConsoleMsg(UserIndex, "Te has equipado el set de agilidad mejorado!!", FontTypeNames.FONTTYPE_FIGHT)
            UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) + 7
        Call WriteAttributes(UserIndex, True)
        Exit Sub
        
    Case eEfectoSet.MuchaMasFuerza
   Call WriteConsoleMsg(UserIndex, "Te has equipado el set de fuerza mejorado!!", FontTypeNames.FONTTYPE_FIGHT)

            UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) + 7
        Call WriteAttributes(UserIndex, True)
        Exit Sub
    Case eEfectoSet.MasVida
   Call WriteConsoleMsg(UserIndex, "Te has equipado el set de Mas Vida!!", FontTypeNames.FONTTYPE_FIGHT)

            UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP + 200
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + 200
              Call WriteUpdateUserStats(UserIndex)
              SetsAura = 6
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, SetsAura, UpdateAuras.Sets))
        Exit Sub
        
    End Select
End Sub

''
' Quitamos los efectos de cada set
Sub QuitarEfecto(ByVal UserIndex As Integer)

'UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)

    Select Case UserList(UserIndex).Stats.TengoSet
    Case eEfectoSet.PegaDoble
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - UserList(UserIndex).Stats.MinHIT
        UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT - UserList(UserIndex).Stats.MaxHIT
        Call WriteUpdateUserStats(UserIndex)
        Exit Sub

    Case eEfectoSet.PegaTripe
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - UserList(UserIndex).Stats.MinHIT - UserList(UserIndex).Stats.MinHIT
        UserList(UserIndex).Stats.MaxHIT = UserList(UserIndex).Stats.MaxHIT - UserList(UserIndex).Stats.MaxHIT - UserList(UserIndex).Stats.MaxHIT
        Call WriteUpdateUserStats(UserIndex)
        Exit Sub

    Case eEfectoSet.MasAgilidad
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) - 4
        Call WriteAttributes(UserIndex, True)
        Exit Sub

    Case eEfectoSet.MasFuerza
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) - 4
        Call WriteAttributes(UserIndex, True)
        Exit Sub

    Case eEfectoSet.MuchaMasAgilidad
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.agilidad) - 7
        Call WriteAttributes(UserIndex, True)
        Exit Sub

    Case eEfectoSet.MuchaMasFuerza
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.fuerza) - 7
        Call WriteAttributes(UserIndex, True)
        Exit Sub
    Case eEfectoSet.MasVida
        Call WriteConsoleMsg(UserIndex, "Te has desequipado el set!!", FontTypeNames.FONTTYPE_FIGHT)
        UserList(UserIndex).Stats.TengoSet = TieneSet(UserIndex)
        UserList(UserIndex).Stats.MaxHP = UserList(UserIndex).Stats.MaxHP - 200
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - 200
         Call WriteUpdateUserStats(UserIndex)
         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSendAura(UserIndex, 0, UpdateAuras.Sets))

    End Select
End Sub

