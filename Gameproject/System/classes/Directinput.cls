VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''��� ���� ������ ������
''''''''''''''''''''''''''''''''
'===================================== DirectInput8
'##############################
'#Written by : jaber al_ani   #
'#Date       : 10-Jan-2006    #
'#Time       : 10.48 PM       #
'##############################


' ������� �� ���� ������� �������
Private DI         As DirectInput8

'������� �� ���� ������
Private Mouse As DirectInputDevice8


'������� �� ���� ��������
Private Keyboard As DirectInputDevice8
'������� �� ���� ���� �������� ��������
Private DIState As DIKEYBOARDSTATE

Private Sub Class_Initialize()
'����� ���� ������� �������
Set DI = DX.DirectInputCreate()

'����� ���� ����� ������
Set Mouse = DI.CreateDevice("GUID_SysMouse")
'��� ����� ������
With Mouse
.SetCommonDataFormat DIFORMAT_MOUSE
.SetCooperativeLevel form3.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
.Acquire
End With


'����� ���� ����� ��������
Set Keyboard = DI.CreateDevice("GUID_SysKeyboard")
'��� ����� ��������
With Keyboard
.SetCommonDataFormat DIFORMAT_KEYBOARD
.SetCooperativeLevel form3.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
.Acquire
End With

End Sub

Public Function CheckKey()
'######################################## Mousw
mouseinp = 5
'��� ���� ������ �� ���� ������
Mouse.GetDeviceStateMouse MouseState

If MouseState.Buttons(0) <> 0 Then mouseinp = 0
If MouseState.Buttons(1) <> 0 Then mouseinp = 1
If MouseState.Buttons(2) <> 0 Then mouseinp = 2
If MouseState.Buttons(3) <> 0 Then mouseinp = 3

'��� ����� ������ ����� �� �����
If MouseState.lX > 1 Or MouseState.lX < -1 Or MouseState.lY > 1 Or MouseState.lY < -1 Then moving = 1

'########################################## KeyBoard

Keyboard.GetDeviceStateKeyboard DIState
If DIState.Key(DIK_ESCAPE) Then End
End Function
