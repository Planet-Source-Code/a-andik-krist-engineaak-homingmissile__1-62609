VERSION 5.00
Begin VB.Form EngineAAK_HomingMiss 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "EngineAAK_HomingMiss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' [923342960150006  S T T D T T S  600051069243329]
' [===============================================]
' [        Introduction: EngineAAK Ver 1.1        ]
' [        -------------------------------        ]
' [         Modify from my First Program:         ]
' [             ".. HomingMissile .."             ]
' [===============================================]
' [ This introduction " EngineAAK " how make unit ]
' [         or missile moving chase mouse         ]
' [-----------------------------------------------]
' [ Not for sale or commercial without permission ]
' [-----------------------------------------------]
' [              By: A. Andik Krist.              ]
' [              -------------------              ]
' [              JAKARTA - INDONESIA              ]
' [-----------------------------------------------]
' [                                               ]
' [        for Comments, Suggestions & Ideas      ]
' [          E-mails me: aakchat@yahoo.com        ]
' [               Date: 15-Sep-2005               ]
' [                                               ]
' [===============91923=29873=30006===============]
Option Explicit

Dim ProgramFinish As Boolean

Private Sub Form_Click()
    ProgramFinish = True
End Sub

Private Sub Form_Load()
    Dim RectFighter As RECT
    Dim i           As Byte
    Dim ScrWidth    As Long
    Dim ScrHeight   As Long
    Dim TxtIntro    As String
    Dim TxtMid      As Integer
        
    '--------------------------------------------------------------
    ScrWidth = 640   ' 640
    ScrHeight = 480  ' 480
    '--------------------------------------------------------------
    ' Init Direct3D
    D3DInit EngineAAK_HomingMiss, ScrWidth, ScrHeight, 16
    '--------------------------------------------------------------
    ' Initialize Frame Direct3D like Root, Camera, Light
    FrameD3DInit ScrWidth, ScrHeight
    '--------------------------------------------------------------
    ' Load Direct 3D Object Fighter (File Extension *.x)
    '--------------------------------------------------------------
    ReDim UnitObject(1)     ' 0=Fighter, 1=Missile
    Set UnitObject(0) = D3D.CreateMeshBuilder()
    With UnitObject(0)
        .LoadFromFile App.Path & "\PlayerFighter.x", 0, 0, Nothing, Nothing
        .ScaleMesh 1.65, 1.65, 1.65
    End With
    ' Load Direct 3D Object Missile
    Set UnitObject(1) = D3D.CreateMeshBuilder()
    With UnitObject(1)
        .LoadFromFile App.Path & "\WeaponEnemyMiss01.x", 0, 0, Nothing, Nothing
        .ScaleMesh 4, 4, 4
    End With
    ' Set Frame for Direct 3D Object
    For i = 0 To UnitCount
        Set UnitFrame(i) = D3D.CreateFrame(FrameRoot)
    Next i
    '--------------------------------------------------------------
    ' Create unit 3D Object: ship
    CreateUnit 400, -100, 4, 2.5, True, 0
    ' Create unit 3D Object: Missile, different speed and turn
    CreateUnit 100, -200, 7, 2, True, 1
    '==============================================================
    ' Load image sequence Fighter
    LoadBMPandSurface PicBuffer, App.Path & "\Fighter_40x40.bmp", PicBufferRECT, 0
    ' Create unit with Image Sequence
    CreateUnit 100, 100, 5, 2, False
    '--------------------------------------------------------------
    Do                                      ' Loop main until ProgramFinish=True
        On Local Error Resume Next
        DoEvents
        D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'ClS Viewport.
        D3D_Device.Update                   ' Update the Direct3D Device.
        
        MoveUnit
        
        'DelayGame 25                        ' Set FPS
        
        D3D_ViewPort.Render FrameRoot       ' Render the 3D Objects must place after Direct3D Render
        
        ' Just Text
        '--------------------------------------------------------------
        TxtIntro = "Introduction EngineAAK Ver 1.1 (from my first program HomingMissile)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 0, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "For : ARCADE / RTS / RPG / RACE (..in test is working..)"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 15, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "Fighter (DirectDraw, Direct3DObject) and Missile (Direct3DObject) will chase mouse cursor"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, 30, TxtIntro, False
        '--------------------------------------------------------------
        TxtIntro = "Click form to EXIT"
        TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
        BackBuffer.DrawText TxtMid, ScrHeight - 20, TxtIntro, False
        
        Primary.Flip Nothing, DDFLIP_WAIT   ' Flip the BackBuffer with the FrontBuffer.
    Loop Until ProgramFinish = True
    End
    
End Sub

