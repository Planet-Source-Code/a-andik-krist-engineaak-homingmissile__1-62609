Attribute VB_Name = "MoveUnit"
Option Explicit

' Get Mouse PointAPI
Type PointAPI
   x As Long
   y As Long
End Type
Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public MousePoint As PointAPI

' this to add create more ship
Global Const UnitCount As Byte = 5 ' 6 ship can create

' Unit Properties for unit (ship)
Type UnitProperties
    Active      As Boolean    ' Active Ship
    Type        As Byte       ' Only for Direct3D Object
                              ' Set Type, 0 = Fighter
                              '           1 = Missile
                              '           2 = .........
    x           As Single     ' Position x, y
    y           As Single     '
    Angle       As Single     ' Direction unit
    AngleTurn   As Single     ' Angle to spin body left or Right if unit turn
    Speed       As Single     ' Speed Unit
    Turn        As Single     ' For turn unit if Object close to destiny turn value will big
    '--------------------------------------------------------------
    Direct3D    As Boolean    ' Code using Direct 3D Object or Image Sequence
End Type
Public Unit(UnitCount)      As UnitProperties

'----------------------------------------------------------------
' Direct 3D Object
'----------------------------------------------------------------
' Frames
Public UnitFrame(UnitCount) As Direct3DRMFrame3

' Meshes (loaded 3D objects from a *.x file)
'
'                 +-> Use () because we don't how many
'                 |   Direct 3D Object want loading
'                 |
Public UnitObject()         As Direct3DRMMeshBuilder3

Sub CreateUnit(x As Single, y As Single, Speed As Single, Turn As Single, Optional UseObject3D As Boolean, Optional TypeObject As Byte)
    Dim i As Byte
    
    For i = 0 To UnitCount
        If Unit(i).Active = False Then
            Unit(i).Active = True
            Unit(i).x = x
            Unit(i).y = y
            Unit(i).Angle = 0
            Unit(i).AngleTurn = 0               ' if use Direct 3D Object files (*.x)
            Unit(i).Speed = Speed
            Unit(i).Turn = Turn
            Unit(i).Direct3D = False
            
            ' For Direct 3D Object
            If UseObject3D = True Then
                Unit(i).x = x
                Unit(i).y = y
                Unit(i).Direct3D = True
                Unit(i).Type = TypeObject
                UnitFrame(i).SetPosition Nothing, Unit(i).x, Unit(i).y, 0
                UnitFrame(i).AddVisual UnitObject(TypeObject)
            End If
            ' After Create get out sub
            Exit Sub
        End If
    Next i
End Sub

Sub MoveUnit()
    Dim i           As Byte
    Dim RectFighter As RECT
    Dim GetMouseX   As Single
    Dim GetMouseY   As Single
    Dim GetRange    As Single
    Dim GetPicNumber As Integer
    
    For i = 0 To UnitCount
        If Unit(i).Active = True Then
        
            GetCursorPos MousePoint
            If Unit(i).Direct3D = False Then
                GetMouseX = MousePoint.x
                GetMouseY = MousePoint.y    ' Different use DirectDraw and Direct3D object
            Else
                GetMouseX = MousePoint.x
                GetMouseY = -MousePoint.y   ' Different use DirectDraw and Direct3D object
                                            ' Axis Y must add (-)
            End If
            
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving unit             ]
         '[-------------------------------------------------]
            Engine Unit(i).Angle, Unit(i).x, Unit(i).y, GetMouseX, GetMouseY, Unit(i).Speed, Unit(i).Turn
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            Unit(i).x = EngineResult.x
            Unit(i).y = EngineResult.y
            Unit(i).Angle = EngineResult.Angle
         '[-------------------------------------------------]
            
            GetRange = Trigonometri(Unit(i).x, Unit(i).y, GetMouseX, GetMouseY, RESULT_RADIUS)

            If Unit(i).Direct3D = False Then
                ' Use Image Sequence
                With RectFighter
                    .Top = 0 + 5
                    .Left = 0
                    .Right = 239
                    .Bottom = 159 + 5
                End With
                GetPicNumber = ((Unit(i).Angle / 360) * 23) ' 23 Pic 0-23 = 24 Pic
                PutImageSequence Unit(i).x, Unit(i).y, GetPicNumber, PicBuffer, RectFighter, 40, 40
                
                ' Only Text, show info
                BackBuffer.DrawText Unit(i).x - 50, Unit(i).y + 40, "DirectDraw, Image Seq", False
                BackBuffer.DrawText Unit(i).x - 50, Unit(i).y + 55, "Speed:" & Unit(i).Speed & " Turn:" & Unit(i).Turn, False
                BackBuffer.DrawText Unit(i).x - 50, Unit(i).y + 70, "Range:" & GetRange, False
            Else
                ' If a ship turn left/right then body ship will make spin
                SpinBodyUnit i, GetMouseX, GetMouseY
                
                ' if ship spin after turn, then make body back to normal position
                SpinBodyUnitToNormal i
            
                ' Set Rotation, Zoom (Scale) and Position Direct 3D
                UnitFrame(i).AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
                UnitFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, DegreeToRadian(-Unit(i).Angle)
                UnitFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -DegreeToRadian(90 + Unit(i).AngleTurn)  ' 90=Position ship up
                UnitFrame(i).AddScale D3DRMCOMBINE_AFTER, 1, 1, 1
                UnitFrame(i).SetPosition Nothing, Unit(i).x, Unit(i).y, 0
                
                ' Only Text, show info, add (-) minus in y Axis, if use draw.....
                BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 20, "Direct3D Object", False
                BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 35, "Speed:" & Unit(i).Speed & " Turn:" & Unit(i).Turn, False
                BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 55, "Range:" & GetRange, False
            End If
        End If
    Next i
End Sub

' Spin body if ship turn left or right
Sub SpinBodyUnit(j As Byte, xDest As Single, yDest As Single)
    Dim PosDegrees As Single
    Dim AddSpin    As Byte
    
    PosDegrees = Int(Trigonometri(Unit(j).x, Unit(j).y, xDest, yDest))
    '---------------------------------------------------------------
    ' If a ship turn left/right then body ship will make little spin
    '---------------------------------------------------------------
    AddSpin = 3     ' is 5 = Standart for Fighter if use DelayGame
    
    If PosDegrees + AddSpin > Unit(j).Angle And PosDegrees - AddSpin < Unit(j).Angle Then
    Else
        ' to now spin ship Left/Right i am use -> DirectTurn in module * EngineAAK *
        If DirectTurn(PosDegrees, Unit(j).Angle, 1) < 0 Then     ' Value minus turn Left/Kiri
            Unit(j).AngleTurn = Unit(j).AngleTurn + Unit(j).Turn * AddSpin
            If Unit(j).AngleTurn > 70 Then Unit(j).AngleTurn = 70
        Else                                                    ' Value plus turn Right/Kanan
           Unit(j).AngleTurn = Unit(j).AngleTurn - Unit(j).Turn * AddSpin
           If Unit(j).AngleTurn < -70 Then Unit(j).AngleTurn = -70
        End If
    End If
End Sub

' After ship spin then make back to normal
Sub SpinBodyUnitToNormal(j As Byte)
    If Unit(j).AngleTurn <> 0 Then
        If Unit(j).AngleTurn < 0 Then
            Unit(j).AngleTurn = Unit(j).AngleTurn + 2
        Else
            Unit(j).AngleTurn = Unit(j).AngleTurn - 2
        End If
    Else
        Unit(j).AngleTurn = 0
    End If
End Sub

Sub PutImageSequence(x As Single, y As Single, _
    GetImageNumber As Integer, surface As DirectDrawSurface4, _
    RECTvar As RECT, WidthPic As Byte, HeightPic As Byte)
    
    Dim StoreNumberSeq      As Integer
    Dim CalcPicPerRow       As Integer
    Dim CalcPicPerColum     As Integer
    '--------------------------------------------------------------
    Dim xGetRECT            As Integer
    Dim yGetRECT            As Integer
    
    StoreNumberSeq = GetImageNumber
    
    yGetRECT = 0
    xGetRECT = WidthPic
    
    CalcPicPerRow = (RECTvar.Right + 1) / WidthPic
    
    If StoreNumberSeq > (CalcPicPerRow - 1) And StoreNumberSeq < (CalcPicPerRow * 2) Then
        yGetRECT = HeightPic
        StoreNumberSeq = StoreNumberSeq - CalcPicPerRow
    End If
    If StoreNumberSeq > ((CalcPicPerRow - 1) * 2) + 1 And StoreNumberSeq < (CalcPicPerRow * 3) Then
        yGetRECT = HeightPic * 2
        StoreNumberSeq = StoreNumberSeq - (CalcPicPerRow * 2)
    End If
    If StoreNumberSeq > ((CalcPicPerRow - 1) * 2) + 2 Then
        yGetRECT = HeightPic * 3
        StoreNumberSeq = StoreNumberSeq - (CalcPicPerRow * 3)
    End If

    With RECTvar
        .Left = StoreNumberSeq * xGetRECT
        .Right = (StoreNumberSeq + 1) * xGetRECT
        .Top = yGetRECT
        .Bottom = HeightPic + yGetRECT
    End With
    
    BackBuffer.BltFast x, y, surface, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub
