Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Sub IniciarDeposito(ByVal UserIndex As Integer)
On Error GoTo errhandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, UserIndex, 0)
'Actualizamos el dinero
Call WriteUpdateUserStats(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
Call WriteBankInit(UserIndex)
UserList(UserIndex).flags.Comerciando = True

errhandler:

End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As UserOBJ)

'UserList(UserIndex).BancoInvent.Object(Slot) = Object


 'Execute ("INSERT INTO vault (cuenta_id, slot, item, quantity) VALUES (" & UserList(UserIndex).MySQLIdCuenta & "," & Slot & "," & Object.ObjIndex & "," & Object.Amount & ")  ON DUPLICATE KEY UPDATE item=VALUES(item), quantity=VALUES(quantity);")


Call WriteChangeBankSlot(UserIndex, Slot, Object.ObjIndex, Object.Amount)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte
Dim Cant As Integer

Dim Obj As UserOBJ
Dim Datos As clsMySQLRecordSet
 Cant = mySQL.SQLQuery("SELECT * FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta, Datos)
    


'Datos("Rep_Promedio
        




'Actualiza un solo slot
If Not UpdateAll Then
Dim i As Integer
    'Actualiza el inventario
    Dim found As Boolean
    For i = 1 To Cant
    
        If Datos("slot") = Slot Then
            Obj.ObjIndex = Datos("item")
            Obj.Amount = Datos("quantity")
            Call SendBanObj(UserIndex, Slot, Obj)
            found = True
        End If

        Datos.MoveNext
    Next i
    
    If Not found Then
     Call SendBanObj(UserIndex, Slot, NullObj)
    End If
    
    
    
    'If UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex > 0 Then
    '    Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent.Object(Slot))
    'Else
    '    Call SendBanObj(UserIndex, Slot, NullObj)
    'End If

Else

'Actualiza todos los slots

    For i = 1 To Cant
    
        If Datos("slot") >= 1 And Datos("slot") <= MAX_BANCOINVENTORY_SLOTS Then
            Obj.ObjIndex = Datos("item")
            Obj.Amount = Datos("quantity")
            Call SendBanObj(UserIndex, Datos("slot"), Obj)
        End If

        Datos.MoveNext
    Next i

    'For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        'If UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
        '    Call SendBanObj(UserIndex, LoopC, UserList(UserIndex).BancoInvent.Object(LoopC))
        'Else
            
        '    Call SendBanObj(UserIndex, LoopC, NullObj)
            
        'End If

    'Next LoopC

End If

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler


If Cantidad < 1 Then Exit Sub
Dim Datos As clsMySQLRecordSet

Call WriteUpdateUserStats(UserIndex)


    Dim Cant As Integer
    Dim Query As String
    Query = "SELECT * FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & i
            
    Cant = mySQL.SQLQuery(Query, Datos)
   
    If Cant = 1 And Datos("quantity") > 0 Then
         If Cantidad > Datos("quantity") Then Cantidad = Datos("quantity")
         'Agregamos el obj que compro al inventario
         Call UserReciveObj(UserIndex, i, Datos("item"), Cantidad)
         'Actualizamos el inventario del usuario
         Call UpdateUserInv(True, UserIndex, 0)
         'Actualizamos el banco
         Call UpdateBanUserInv(False, UserIndex, i)
    End If
    'Actualizamos la ventana de comercio
    Call UpdateVentanaBanco(UserIndex)


errhandler:

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal bSlot, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer


'If UserList(UserIndex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = ObjIndex 'UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
    
    Call QuitarBancoInvItem(UserIndex, bSlot, Cantidad)
Else
    Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
End If


End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
'ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
Dim Query As String
    'Quita un Obj
    Dim Datos As clsMySQLRecordSet
Query = "SELECT * FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & Slot
Dim Cant As Integer

Cant = mySQL.SQLQuery(Query, Datos)

If Cant > 0 Then
    If Datos("quantity") - Cantidad <= 0 Then
        Execute ("DELETE FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & Slot)
    Else
        Execute ("UPDATE vault SET quantity=" & Datos("quantity") - Cantidad & " WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & Slot)
    End If
End If

      ' UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount - Cantidad
        
       ' If UserList(UserIndex).BancoInvent.Object(Slot).Amount <= 0 Then
       '     UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
       '     UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
       '     UserList(UserIndex).BancoInvent.Object(Slot).Amount = 0
       ' End If

    
    
End Sub

Sub UpdateVentanaBanco(ByVal UserIndex As Integer)
    Call WriteBankOK(UserIndex)
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler
    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
        If Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
        'Agregamos el obj que deposita al banco
        If UserDejaObj(UserIndex, CInt(Item), Cantidad) Then
        
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, UserIndex, Item)
            
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, UserIndex, 0)
        End If
    End If
    
    'Actualizamos la ventana del banco
    Call UpdateVentanaBanco(UserIndex)
errhandler:
End Sub

Function UserDejaObj(ByVal UserIndex As Integer, ByVal InvSlot As Integer, ByVal Cantidad As Integer) As Boolean
    Dim Slot As Integer
    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Function
    
    obji = UserList(UserIndex).Invent.Object(InvSlot).ObjIndex
    
    If ObjData(obji).Newbie = 1 Then
        Call WriteConsoleMsg(UserIndex, "No puedes depositar tus items newbies.", FontTypeNames.FONTTYPE_INFO)
        UserDejaObj = False
        Exit Function
    End If
    
    '¿Ya tiene un objeto de este tipo?
    Dim j As Integer

Dim Cant As Integer
Dim Datos As clsMySQLRecordSet
Dim Query As String
Query = "SELECT * FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " ORDER BY slot ASC"
        
Cant = mySQL.SQLQuery(Query, Datos)


Dim BancoInvent(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ

For j = 1 To Cant

    BancoInvent(Datos("slot")).Amount = Datos("quantity")
    BancoInvent(Datos("slot")).ObjIndex = Datos("item")

    Datos.MoveNext
Next j

    
    
    
    
    Slot = 1
    Do Until BancoInvent(Slot).ObjIndex = obji And _
        BancoInvent(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        Slot = Slot + 1
        
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until BancoInvent(Slot).ObjIndex = 0
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                UserDejaObj = False
                Exit Function
            End If
        Loop
        
        'UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
    End If
    
    If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        If BancoInvent(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
            'Menor que MAX_INV_OBJS
            BancoInvent(Slot).ObjIndex = obji
            BancoInvent(Slot).Amount = BancoInvent(Slot).Amount + Cantidad
            
            Execute ("INSERT INTO vault (cuenta_id, slot, item, quantity) VALUES (" & UserList(UserIndex).MySQLIdCuenta & "," & Slot & "," & obji & "," & BancoInvent(Slot).Amount & ")  ON DUPLICATE KEY UPDATE item=VALUES(item), quantity=VALUES(quantity);")

            
            Call QuitarUserInvItem(UserIndex, CByte(InvSlot), Cantidad)
        Else
            Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
    UserDejaObj = True
End Function

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer

Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
'Call WriteConsoleMsg(sendIndex, " Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

'For j = 1 To MAX_BANCOINVENTORY_SLOTS
'    If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
'        Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
'    End If
'Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer)
On Error Resume Next
Dim j As Integer
Dim ObjInd As Long, ObjCant As Long

Dim Datos As clsMySQLRecordSet
Dim Query As String
Dim Tmp As String
Dim Cant As Long
    
Query = "SELECT * FROM vault WHERE cuenta_id=" & UserList(sendIndex).MySQLIdCuenta
        
Cant = mySQL.SQLQuery(Query, Datos)
    
If Cant > 0 Then
    'Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, " Tiene " & Datos("BanCantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
    For j = 1 To Cant
        'Tmp = datos("BanObj" & i)
        ObjInd = Datos("item")
        ObjCant = Datos("quantity")
        If ObjInd > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
        End If
    Next
Else
    Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & UserList(sendIndex).Name, FontTypeNames.FONTTYPE_INFO)
End If

End Sub


Public Sub IntercambiarBanco(ByVal UserIndex As Integer, ByVal Slot1 As Integer, ByVal Slot2 As Integer)
If Slot1 < 1 Or Slot1 > MAX_BANCOINVENTORY_SLOTS Or Slot2 < 1 Or Slot2 > MAX_BANCOINVENTORY_SLOTS Then Exit Sub


Dim Obj1 As UserOBJ
Dim Obj2 As UserOBJ
Dim j As Integer

Dim Cant As Integer
Dim Datos As clsMySQLRecordSet
Dim Query As String
Query = "SELECT * FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND (slot=" & Slot1 & " OR slot=" & Slot2 & ")"
        
Cant = mySQL.SQLQuery(Query, Datos)

For j = 1 To Cant
    If Datos("slot") = Slot1 Then
        Obj1.Amount = Datos("quantity")
        Obj1.ObjIndex = Datos("item")
    ElseIf Datos("slot") = Slot2 Then
        Obj2.Amount = Datos("quantity")
        Obj2.ObjIndex = Datos("item")
    End If
    Datos.MoveNext
Next j


If Obj1.ObjIndex > 0 Then
    Execute ("INSERT INTO vault (cuenta_id, slot, item, quantity) VALUES (" & UserList(UserIndex).MySQLIdCuenta & "," & Slot2 & "," & Obj1.ObjIndex & "," & Obj1.Amount & ")  ON DUPLICATE KEY UPDATE item=VALUES(item), quantity=VALUES(quantity);")
Else
    Execute ("DELETE FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & Slot2)
End If

If Obj2.ObjIndex > 0 Then
    Execute ("INSERT INTO vault (cuenta_id, slot, item, quantity) VALUES (" & UserList(UserIndex).MySQLIdCuenta & "," & Slot1 & "," & Obj2.ObjIndex & "," & Obj2.Amount & ")  ON DUPLICATE KEY UPDATE item=VALUES(item), quantity=VALUES(quantity);")
Else
    Execute ("DELETE FROM vault WHERE cuenta_id=" & UserList(UserIndex).MySQLIdCuenta & " AND slot=" & Slot1)
End If



'With UserList(UserIndex)


'    tmpObj = .BancoInvent.Object(Slot1)
'    .BancoInvent.Object(Slot1) = .BancoInvent.Object(Slot2)
'    .BancoInvent.Object(Slot2) = tmpObj
'End With
End Sub

