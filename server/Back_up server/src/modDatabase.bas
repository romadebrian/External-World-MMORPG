Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.Path & "\data\options.ini", "GUILDS", "Buy_Cost", STR(Options.Buy_Cost)
    PutVar App.Path & "\data\options.ini", "GUILDS", "Buy_Lvl", STR(Options.Buy_Lvl)
    PutVar App.Path & "\data\options.ini", "GUILDS", "Buy_Item", STR(Options.Buy_Item)
    PutVar App.Path & "\data\options.ini", "GUILDS", "Join_Cost", STR(Options.Join_Cost)
    PutVar App.Path & "\data\options.ini", "GUILDS", "Join_Lvl", STR(Options.Join_Lvl)
    PutVar App.Path & "\data\options.ini", "GUILDS", "Join_Item", STR(Options.Join_Item)
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    Options.Buy_Cost = GetVar(App.Path & "\data\options.ini", "GUILDS", "Buy_Cost")
    Options.Buy_Lvl = GetVar(App.Path & "\data\options.ini", "GUILDS", "Buy_Lvl")
    Options.Buy_Item = GetVar(App.Path & "\data\options.ini", "GUILDS", "Buy_Item")
    Options.Join_Cost = GetVar(App.Path & "\data\options.ini", "GUILDS", "Join_Cost")
    Options.Join_Lvl = GetVar(App.Path & "\data\options.ini", "GUILDS", "Join_Lvl")
    Options.Join_Item = GetVar(App.Path & "\data\options.ini", "GUILDS", "Join_Item")
    
    frmServer.txtGBuyCost.Text = Options.Buy_Cost
    frmServer.txtGBuyLvl.Text = Options.Buy_Lvl
    frmServer.txtGBuyItem.Text = Options.Buy_Item
    frmServer.txtGJoinCost.Text = Options.Join_Cost
    frmServer.txtGJoinLvl.Text = Options.Join_Lvl
    frmServer.txtGJoinItem.Text = Options.Join_Item
    
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    filename = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & "Server"
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).Name = Name
        Player(Index).Sex = Sex
        Player(Index).Class = ClassNum
        
        If Player(Index).Sex = SEX_MALE Then
            Player(Index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(Index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(Index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).stat(n) = Class(ClassNum).stat(n)
        Next n

        Player(Index).Dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).x = START_X
        Player(Index).y = START_Y
        Player(Index).Spawn.Map = START_MAP
        Player(Index).Spawn.x = START_X
        Player(Index).Spawn.y = START_Y
        Player(Index).Dir = DIR_DOWN
        Player(Index).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        
        For n = 1 To MAX_COMBAT
            Player(Index).Combat(n).Level = 1
            Player(Index).Combat(n).EXP = 0
        Next
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(Index).Inv(n).num = Class(ClassNum).StartItem(n)
                        Player(Index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).Name)) > 0 Then
                        Player(Index).Spell(n) = Class(ClassNum).StartSpell(n)
                    End If
                End If
            Next
        End If
        
        ' Relate the player's skills
        Call CheckSkills(Index, False)
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If isPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
    Call ClearPlayer(Index)
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Class = 1

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemnum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.Path & "\data\items\item" & itemnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(itemnum)
    Close #F
End Sub

Sub SaveCombo(ByVal comboNum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.Path & "\data\combos\combo" & comboNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Combo(comboNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

End Sub

Sub LoadCombos()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckCombos

    For i = 1 To MAX_COMBOS
        filename = App.Path & "\data\Combos\Combo" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Combo(i)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub CheckCombos()
    Dim i As Long

    For i = 1 To MAX_COMBOS

        If Not FileExist("\Data\Combos\Combo" & i & ".dat") Then
            Call SaveCombo(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = "None."
End Sub

Sub ClearCombo(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Combo(Index)), LenB(Combo(Index)))
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Sub ClearCombos()
    Dim i As Long

    For i = 1 To MAX_COMBOS
        Call ClearCombo(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellnum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\spells\spells" & spellnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(spellnum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal npcNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\npcs\npc" & npcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , NPC(npcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , NPC(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim y As Long, i As Long, z As Long, w As Long
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(mapnum).Name
    Put #F, , Map(mapnum).Music
    Put #F, , Map(mapnum).BGS
    Put #F, , Map(mapnum).Revision
    Put #F, , Map(mapnum).Moral
    Put #F, , Map(mapnum).Up
    Put #F, , Map(mapnum).Down
    Put #F, , Map(mapnum).Left
    Put #F, , Map(mapnum).Right
    Put #F, , Map(mapnum).BootMap
    Put #F, , Map(mapnum).BootX
    Put #F, , Map(mapnum).BootY
    
    Put #F, , Map(mapnum).Weather
    Put #F, , Map(mapnum).WeatherIntensity
    
    Put #F, , Map(mapnum).Fog
    Put #F, , Map(mapnum).FogSpeed
    Put #F, , Map(mapnum).FogOpacity
    
    Put #F, , Map(mapnum).Red
    Put #F, , Map(mapnum).Green
    Put #F, , Map(mapnum).Blue
    Put #F, , Map(mapnum).Alpha
    
    Put #F, , Map(mapnum).MaxX
    Put #F, , Map(mapnum).MaxY

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            Put #F, , Map(mapnum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(mapnum).NPC(x)
        Put #F, , Map(mapnum).NpcSpawnType(x)
    Next
    Close #F
    
    'This is for event saving, it is in .ini files becuase there are non-limited values (strings) that cannot easily be loaded/saved in the normal manner.
    filename = App.Path & "\data\maps\map" & mapnum & "_eventdata.dat"
    PutVar filename, "Events", "EventCount", Val(Map(mapnum).EventCount)
    
    If Map(mapnum).EventCount > 0 Then
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "Global", Val(.Global)
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkVariable", Val(.chkVariable)
                        PutVar filename, "Event" & i & "Page" & x, "VariableIndex", Val(.VariableIndex)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCondition", Val(.VariableCondition)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSwitch", Val(.chkSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchIndex", Val(.SwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemIndex", Val(.HasItemIndex)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemAmount", Val(.HasItemAmount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX2", Val(.GraphicX2)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar filename, "Event" & i & "Page" & x, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar filename, "Event" & i & "Page" & x, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index", Val(.MoveRoute(y).Index)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1", Val(.MoveRoute(y).Data1)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2", Val(.MoveRoute(y).Data2)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3", Val(.MoveRoute(y).Data3)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4", Val(.MoveRoute(y).Data4)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5", Val(.MoveRoute(y).Data5)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6", Val(.MoveRoute(y).Data6)
                            Next
                        End If
                        
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "ShowName", Val(.ShowName)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "Position", Val(.Position)
                    End With
                    
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount", Val(Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount)
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList", Val(Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList)
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Index", Val(.Index)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text1", .Text1
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text2", .Text2
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text3", .Text3
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text4", .Text4
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Text5", .Text5
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data5", Val(.Data5)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "Data6", Val(.Data6)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).Index)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).Data5)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).Data6)
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).BGS
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        
        Get #F, , Map(i).Weather
        Get #F, , Map(i).WeatherIntensity
        
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).NPC(x)
            Get #F, , Map(i).NpcSpawnType(x)
            MapNpc(i).NPC(x).num = Map(i).NPC(x)
        Next

        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(filename, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                With Map(z).Events(i)
                    .Name = GetVar(filename, "Event" & i, "Name")
                    .Global = Val(GetVar(filename, "Event" & i, "Global"))
                    .x = Val(GetVar(filename, "Event" & i, "x"))
                    .y = Val(GetVar(filename, "Event" & i, "y"))
                    .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
                End With
                If Map(z).Events(i).PageCount > 0 Then
                    ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                    For x = 1 To Map(z).Events(i).PageCount
                        With Map(z).Events(i).Pages(x)
                            .chkVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "chkVariable"))
                            .VariableIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableIndex"))
                            .VariableCondition = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCondition"))
                            .VariableCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCompare"))
                            
                            .chkSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSwitch"))
                            .SwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchIndex"))
                            .SwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchCompare"))
                            
                            .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                            .HasItemIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemIndex"))
                            .HasItemAmount = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemAmount"))
                            
                            .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                            .SelfSwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchIndex"))
                            .SelfSwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchCompare"))
                            
                            .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                            .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                            .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                            .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                            .GraphicX2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX2"))
                            .GraphicY2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY2"))
                            
                            .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                            .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                            .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                            
                            .IgnoreMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "IgnoreMoveRoute"))
                            .RepeatMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "RepeatMoveRoute"))
                            
                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRouteCount"))
                            
                            If .MoveRouteCount > 0 Then
                                ReDim Map(z).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                                For y = 1 To .MoveRouteCount
                                    .MoveRoute(y).Index = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Index"))
                                    .MoveRoute(y).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data1"))
                                    .MoveRoute(y).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data2"))
                                    .MoveRoute(y).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data3"))
                                    .MoveRoute(y).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data4"))
                                    .MoveRoute(y).Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data5"))
                                    .MoveRoute(y).Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & y & "Data6"))
                                Next
                            End If
                            
                            .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                            .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                            .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                            .ShowName = Val(GetVar(filename, "Event" & i & "Page" & x, "ShowName"))
                            .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                            .CommandListCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandListCount"))
                         
                            .Position = Val(GetVar(filename, "Event" & i & "Page" & x, "Position"))
                        End With
                            
                        If Map(z).Events(i).Pages(x).CommandListCount > 0 Then
                            ReDim Map(z).Events(i).Pages(x).CommandList(0 To Map(z).Events(i).Pages(x).CommandListCount)
                            For y = 1 To Map(z).Events(i).Pages(x).CommandListCount
                                Map(z).Events(i).Pages(x).CommandList(y).CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "CommandCount"))
                                Map(z).Events(i).Pages(x).CommandList(y).ParentList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "ParentList"))
                                If Map(z).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(x).CommandList(y).Commands(Map(z).Events(i).Pages(x).CommandList(y).CommandCount)
                                    For p = 1 To Map(z).Events(i).Pages(x).CommandList(y).CommandCount
                                        With Map(z).Events(i).Pages(x).CommandList(y).Commands(p)
                                            .Index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Index"))
                                            .Text1 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text1")
                                            .Text2 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text2")
                                            .Text3 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text3")
                                            .Text4 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text4")
                                            .Text5 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Text5")
                                            .Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data1"))
                                            .Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data2"))
                                            .Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data3"))
                                            .Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data4"))
                                            .Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data5"))
                                            .Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "Data6"))
                                            .ConditionalBranch.CommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCommandList"))
                                            .ConditionalBranch.Condition = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchCondition"))
                                            .ConditionalBranch.Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData1"))
                                            .ConditionalBranch.Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData2"))
                                            .ConditionalBranch.Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchData3"))
                                            .ConditionalBranch.ElseCommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "ConditionalBranchElseCommandList"))
                                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRouteCount"))
                                            If .MoveRouteCount > 0 Then
                                                ReDim .MoveRoute(1 To .MoveRouteCount)
                                                For w = 1 To .MoveRouteCount
                                                    .MoveRoute(w).Index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Index"))
                                                    .MoveRoute(w).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                    .MoveRoute(w).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                    .MoveRoute(w).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                    .MoveRoute(w).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                    .MoveRoute(w).Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                    .MoveRoute(w).Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & y & "Command" & p & "MoveRoute" & w & "Data6"))
                                                Next
                                            End If
                                        End With
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, Index)), LenB(MapItem(mapnum, Index)))
    MapItem(mapnum, Index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal mapnum As Long)
    ReDim MapNpc(mapnum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).NPC(Index)), LenB(MapNpc(mapnum).NPC(Index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(mapnum)), LenB(Map(mapnum)))
    Map(mapnum).Name = vbNullString
    Map(mapnum).MaxX = MAX_MAPX
    Map(mapnum).MaxY = MAX_MAPY
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    On Error GoTo ErrHandler
    GetClassName = Trim$(Class(ClassNum).Name)
Exit Function
ErrHandler:
    HandleError "GetClassName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal stat As Stats) As Long
    GetClassStat = Class(ClassNum).stat(stat)
End Function

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(Index)
    Close #F
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    Call ClearBank(Index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(Index)
    Close #F

End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub ClearParty(ByVal partyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))
End Sub

Sub SaveSwitches()
Dim i As Long, filename As String
filename = App.Path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

End Sub

Sub SaveVariables()
Dim i As Long, filename As String
filename = App.Path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

End Sub

Sub LoadSwitches()
Dim i As Long, filename As String
filename = App.Path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
Next
End Sub

Sub LoadVariables()
Dim i As Long, filename As String
filename = App.Path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
Next
End Sub

Sub ClearProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
    ' clear the projectile
    With TempPlayer(Index).ProjecTile(PlayerProjectile)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .x = 0
        .y = 0
        .Range = 0
        .Damage = 0
        .Speed = 0
    End With
End Sub
