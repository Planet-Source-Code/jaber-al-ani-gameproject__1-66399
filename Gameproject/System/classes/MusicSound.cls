VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicSOUND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================== DirectSound8
'##############################
'#Written by : jaber al_ani   #
'#Date       : 22-Nov-2005    #
'#Time       : 01.30 PM       #
'##############################


'������� �� ���� �����
Private DS As DirectSound8

'������� �� ��� �����
Private DSBuf As DirectSoundSecondaryBuffer8

'������� �� ���� ���� �����
Private DSBDesc As DSBUFFERDESC
'���� ����� ��� ����� ��� ��� ������ ��� ��


Public Function create(name As String)
ntime = 0
'����� ���� ����� �������
Set DS = DX.DirectSoundCreate("")

'��� ����� ������ ������� ��� �� �������� �� �������
DS.SetCooperativeLevel form3.hWnd, DSSCL_PRIORITY

'��� ���� ���� ����� ����� �� ������ ������� ����� ���
DSBDesc.lFlags = DSBCAPS_CTRLVOLUME

'����� ����� ���� �����
Set DSBuf = DS.CreateSoundBufferFromFile(App.Path & "\Data\Sounds\" + name + ".wav", DSBDesc)


End Function
Public Function PLAY()

DSBuf.PLAY DSBPLAY_DEFAULT

End Function

Public Function PLAYnow()

DSBuf.SetCurrentPosition 0

DSBuf.PLAY DSBPLAY_DEFAULT

End Function

Public Function stopplay()
DSBuf.Stop
DSBuf.SetCurrentPosition 0
End Function



