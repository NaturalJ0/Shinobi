Attribute VB_Name = "iniload"
Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) _
As Long

Public Type ShinobiData
    Name As String
    Hp As Double
    Shuriken As Double
    Flame As Double
    Side As Double
    Back As Double
    Tate As Double
End Type
Public Data(16, 7) As ShinobiData
Public Tatename(9) As String
Public Stage(17) As String
Private i As Long
Private Buff As String * 100
Private State(9) As String
Private Sname(16, 7) As String
Private Stagename(17) As String
Private sx As String
Private sy As String
Private p As Integer
Private x As Integer
Private y As Integer
Private iniFileName As String
Public Sub pre()
Rem 1A
Data(1, 1).Name = "����"
Data(1, 1).Hp = 10
Data(1, 1).Shuriken = 2
Data(1, 1).Flame = 50
Data(1, 1).Side = 1.5
Data(1, 1).Back = 3
Data(1, 1).Tate = 7

Data(1, 2).Name = "��"
Data(1, 2).Hp = 5
Data(1, 2).Shuriken = 3
Data(1, 2).Flame = 75
Data(1, 2).Side = 2
Data(1, 2).Back = 3
Data(1, 2).Tate = 7

Data(1, 3).Name = "̹��"
Data(1, 3).Hp = 30
Data(1, 3).Shuriken = 0
Data(1, 3).Flame = 50
Data(1, 3).Side = 1
Data(1, 3).Back = 1
Data(1, 3).Tate = 4

Data(1, 4).Name = "������BOSS��"
Data(1, 4).Hp = 2.5
Data(1, 4).Shuriken = 2
Data(1, 4).Flame = 50
Data(1, 4).Side = 1
Data(1, 4).Back = 1
Data(1, 4).Tate = 5

Data(1, 5).Name = "���̣�BOSS��"
Data(1, 5).Hp = 10
Data(1, 5).Shuriken = 2
Data(1, 5).Flame = 50
Data(1, 5).Side = 1.5
Data(1, 5).Back = 3
Data(1, 5).Tate = 4

Data(1, 6).Name = "ֱ������BOSS��"
Data(1, 6).Hp = 100
Data(1, 6).Shuriken = 0
Data(1, 6).Flame = 50
Data(1, 6).Side = 1
Data(1, 6).Back = 1
Data(1, 6).Tate = 5

Rem 1B
Data(2, 1).Name = "����"
Data(2, 1).Hp = 10
Data(2, 1).Shuriken = 2
Data(2, 1).Flame = 50
Data(2, 1).Side = 1.5
Data(2, 1).Back = 3
Data(2, 1).Tate = 9

Data(2, 2).Name = "�Ʒ�ͷ"
Data(2, 2).Hp = 2.5
Data(2, 2).Shuriken = 2
Data(2, 2).Flame = 50
Data(2, 2).Side = 1.5
Data(2, 2).Back = 3
Data(2, 2).Tate = 9

Data(2, 3).Name = "��Ů��"
Data(2, 3).Hp = 7.5
Data(2, 3).Shuriken = 3
Data(2, 3).Flame = 75
Data(2, 3).Side = 2
Data(2, 3).Back = 3
Data(2, 3).Tate = 9

Data(2, 5).Name = "���̣�BOSS��"
Data(2, 5).Hp = 10
Data(2, 5).Shuriken = 2
Data(2, 5).Flame = 50
Data(2, 5).Side = 1.5
Data(2, 5).Back = 3
Data(2, 5).Tate = 4

Data(2, 6).Name = "�غ㣨BOSS��"
Data(2, 6).Hp = 187.5
Data(2, 6).Shuriken = 0
Data(2, 6).Flame = 50
Data(2, 6).Side = 1.5
Data(2, 6).Back = 3
Data(2, 6).Tate = 5

Rem 2A
Data(3, 1).Name = "����"
Data(3, 1).Hp = 10
Data(3, 1).Shuriken = 2
Data(3, 1).Flame = 50
Data(3, 1).Side = 1.5
Data(3, 1).Back = 3
Data(3, 1).Tate = 9

Data(3, 2).Name = "��"
Data(3, 2).Hp = 10
Data(3, 2).Shuriken = 3
Data(3, 2).Flame = 75
Data(3, 2).Side = 2
Data(3, 2).Back = 3
Data(3, 2).Tate = 6

Data(3, 3).Name = "ѻ�칷"
Data(3, 3).Hp = 60
Data(3, 3).Shuriken = 1.5
Data(3, 3).Flame = 37.5
Data(3, 3).Side = 1.5
Data(3, 3).Back = 3
Data(3, 3).Tate = 9

Data(3, 4).Name = "����BOSS��"
Data(3, 4).Hp = 10
Data(3, 4).Shuriken = 3
Data(3, 4).Flame = 75
Data(3, 4).Side = 2
Data(3, 4).Back = 3
Data(3, 4).Tate = 5

Data(3, 5).Name = "���У�BOSS��"
Data(3, 5).Hp = 150
Data(3, 5).Shuriken = 2
Data(3, 5).Flame = 50
Data(3, 5).Side = 1.5
Data(3, 5).Back = 3
Data(3, 5).Tate = 6

Data(3, 6).Name = "ͭŮ��BOSS��"
Data(3, 6).Hp = 112.5
Data(3, 6).Shuriken = 2
Data(3, 6).Flame = 50
Data(3, 6).Side = 2
Data(3, 6).Back = 3
Data(3, 6).Tate = 6

Rem 2B
Data(4, 1).Name = "����"
Data(4, 1).Hp = 25
Data(4, 1).Shuriken = 2
Data(4, 1).Flame = 50
Data(4, 1).Side = 1.5
Data(4, 1).Back = 3
Data(4, 1).Tate = 9

Data(4, 2).Name = "�Ʒ�ͷ"
Data(4, 2).Hp = 2.5
Data(4, 2).Shuriken = 2
Data(4, 2).Flame = 50
Data(4, 2).Side = 1.5
Data(4, 2).Back = 3
Data(4, 2).Tate = 9

Data(4, 3).Name = "ѻ�칷"
Data(4, 3).Hp = 60
Data(4, 3).Shuriken = 1.5
Data(4, 3).Flame = 37.5
Data(4, 3).Side = 1.5
Data(4, 3).Back = 3
Data(4, 3).Tate = 5

Data(4, 5).Name = "���̣�BOSS��"
Data(4, 5).Hp = 25
Data(4, 5).Shuriken = 2
Data(4, 5).Flame = 50
Data(4, 5).Side = 1.5
Data(4, 5).Back = 3
Data(4, 5).Tate = 4

Data(4, 6).Name = "ǿ��ֱ������BOSS��"
Data(4, 6).Hp = 200
Data(4, 6).Shuriken = 0
Data(4, 6).Flame = 50
Data(4, 6).Side = 1
Data(4, 6).Back = 1
Data(4, 6).Tate = 5

Rem 3A
Data(5, 1).Name = "����"
Data(5, 1).Hp = 25
Data(5, 1).Shuriken = 2
Data(5, 1).Flame = 50
Data(5, 1).Side = 1.5
Data(5, 1).Back = 3
Data(5, 1).Tate = 8

Data(5, 2).Name = "��"
Data(5, 2).Hp = 12.5
Data(5, 2).Shuriken = 3
Data(5, 2).Flame = 75
Data(5, 2).Side = 2
Data(5, 2).Back = 3
Data(5, 2).Tate = 7

Data(5, 3).Name = "ǿ��̹��"
Data(5, 3).Hp = 75
Data(5, 3).Shuriken = 0
Data(5, 3).Flame = 50
Data(5, 3).Side = 1
Data(5, 3).Back = 1
Data(5, 3).Tate = 4

Data(5, 5).Name = "����BOSS��"
Data(5, 5).Hp = 12.5
Data(5, 5).Shuriken = 3
Data(5, 5).Flame = 75
Data(5, 5).Side = 2
Data(5, 5).Back = 3
Data(5, 5).Tate = 6

Data(5, 6).Name = "���֣�BOSS��"
Data(5, 6).Hp = 187.5
Data(5, 6).Shuriken = 2
Data(5, 6).Flame = 50
Data(5, 6).Side = 3
Data(5, 6).Back = 0
Data(5, 6).Tate = 7

Rem 3B
Data(6, 1).Name = "����"
Data(6, 1).Hp = 25
Data(6, 1).Shuriken = 2
Data(6, 1).Flame = 50
Data(6, 1).Side = 1.5
Data(6, 1).Back = 3
Data(6, 1).Tate = 7

Data(6, 2).Name = "֩��"
Data(6, 2).Hp = 6.5
Data(6, 2).Shuriken = 2
Data(6, 2).Flame = 50
Data(6, 2).Side = 2
Data(6, 2).Back = 3
Data(6, 2).Tate = 7

Data(6, 3).Name = "��Ů��"
Data(6, 3).Hp = 22.5
Data(6, 3).Shuriken = 3
Data(6, 3).Flame = 75
Data(6, 3).Side = 2
Data(6, 3).Back = 3
Data(6, 3).Tate = 7

Data(6, 4).Name = "֩�볲"
Data(6, 4).Hp = 20
Data(6, 4).Shuriken = 2
Data(6, 4).Flame = 50
Data(6, 4).Side = 1
Data(6, 4).Back = 1
Data(6, 4).Tate = 7

Data(6, 5).Name = "֩�루BOSS��"
Data(6, 5).Hp = 6.5
Data(6, 5).Shuriken = 2
Data(6, 5).Flame = 50
Data(6, 5).Side = 2
Data(6, 5).Back = 3
Data(6, 5).Tate = 5

Data(6, 6).Name = "���ƣ�BOSS��"
Data(6, 6).Hp = 225
Data(6, 6).Shuriken = 0
Data(6, 6).Flame = 37.5
Data(6, 6).Side = 1
Data(6, 6).Back = 1
Data(6, 6).Tate = 6

Data(6, 0).Name = "֩�볲��BOSS��"
Data(6, 0).Hp = 20
Data(6, 0).Shuriken = 2
Data(6, 0).Flame = 50
Data(6, 0).Side = 1
Data(6, 0).Back = 1
Data(6, 0).Tate = 6

Rem 4A
Data(7, 1).Name = "����"
Data(7, 1).Hp = 30
Data(7, 1).Shuriken = 2
Data(7, 1).Flame = 50
Data(7, 1).Side = 1.5
Data(7, 1).Back = 3
Data(7, 1).Tate = 7

Data(7, 2).Name = "����"
Data(7, 2).Hp = 7.5
Data(7, 2).Shuriken = 3
Data(7, 2).Flame = 75
Data(7, 2).Side = 2
Data(7, 2).Back = 3
Data(7, 2).Tate = 7

Data(7, 3).Name = "ǿ��̹��"
Data(7, 3).Hp = 75
Data(7, 3).Shuriken = 0
Data(7, 3).Flame = 50
Data(7, 3).Side = 1
Data(7, 3).Back = 1
Data(7, 3).Tate = 7

Data(7, 5).Name = "��ͷ��BOSS��"
Data(7, 5).Hp = 7.5
Data(7, 5).Shuriken = 3
Data(7, 5).Flame = 75
Data(7, 5).Side = 2
Data(7, 5).Back = 3
Data(7, 5).Tate = 6

Data(7, 6).Name = "�棨BOSS��"
Data(7, 6).Hp = 225
Data(7, 6).Shuriken = 2
Data(7, 6).Flame = 0
Data(7, 6).Side = 1.5
Data(7, 6).Back = 3
Data(7, 6).Tate = 7

Rem 4B
Data(8, 1).Name = "����"
Data(8, 1).Hp = 30
Data(8, 1).Shuriken = 2
Data(8, 1).Flame = 50
Data(8, 1).Side = 1.5
Data(8, 1).Back = 3
Data(8, 1).Tate = 9

Data(8, 2).Name = "����"
Data(8, 2).Hp = 7.5
Data(8, 2).Shuriken = 3
Data(8, 2).Flame = 75
Data(8, 2).Side = 2
Data(8, 2).Back = 3
Data(8, 2).Tate = 9

Data(8, 3).Name = "ǿ��̹��"
Data(8, 3).Hp = 75
Data(8, 3).Shuriken = 0
Data(8, 3).Flame = 50
Data(8, 3).Side = 1
Data(8, 3).Back = 1
Data(8, 3).Tate = 9

Data(8, 5).Name = "���ӣ�BOSS��"
Data(8, 5).Hp = 7.5
Data(8, 5).Shuriken = 3
Data(8, 5).Flame = 75
Data(8, 5).Side = 2
Data(8, 5).Back = 3
Data(8, 5).Tate = 6

Data(8, 6).Name = "����꣨BOSS��"
Data(8, 6).Hp = 225
Data(8, 6).Shuriken = 0
Data(8, 6).Flame = 50
Data(8, 6).Side = 1
Data(8, 6).Back = 0
Data(8, 6).Tate = 7

Rem 5A
Data(9, 1).Name = "����"
Data(9, 1).Hp = 20
Data(9, 1).Shuriken = 2
Data(9, 1).Flame = 50
Data(9, 1).Side = 1.5
Data(9, 1).Back = 3
Data(9, 1).Tate = 9

Data(9, 2).Name = "�ڷ�ͷ"
Data(9, 2).Hp = 8.5
Data(9, 2).Shuriken = 2
Data(9, 2).Flame = 50
Data(9, 2).Side = 1.5
Data(9, 2).Back = 3
Data(9, 2).Tate = 9

Data(9, 3).Name = "��Ů��"
Data(9, 3).Hp = 22.5
Data(9, 3).Shuriken = 3
Data(9, 3).Flame = 75
Data(9, 3).Side = 2
Data(9, 3).Back = 3
Data(9, 3).Tate = 9

Data(9, 5).Name = "��ͷ��BOSS��"
Data(9, 5).Hp = 9
Data(9, 5).Shuriken = 3
Data(9, 5).Flame = 75
Data(9, 5).Side = 2
Data(9, 5).Back = 3
Data(9, 5).Tate = 7

Data(9, 6).Name = "��գ�BOSS��"
Data(9, 6).Hp = 375
Data(9, 6).Shuriken = 1.5
Data(9, 6).Flame = 37.5
Data(9, 6).Side = 1.5
Data(9, 6).Back = 3
Data(9, 6).Tate = 8

Rem 5B
Data(10, 5).Name = "�ߣ�BOSS��"
Data(10, 5).Hp = 9
Data(10, 5).Shuriken = 3
Data(10, 5).Flame = 75
Data(10, 5).Side = 1.5
Data(10, 5).Back = 3
Data(10, 5).Tate = 8

Data(10, 6).Name = "�����ߣ�BOSS��"
Data(10, 6).Hp = 300
Data(10, 6).Shuriken = 0
Data(10, 6).Flame = 75
Data(10, 6).Side = 1
Data(10, 6).Back = 1
Data(10, 6).Tate = 9

Rem 6A
Data(11, 1).Name = "��Ů��"
Data(11, 1).Hp = 22.5
Data(11, 1).Shuriken = 3
Data(11, 1).Flame = 75
Data(11, 1).Side = 2
Data(11, 1).Back = 3
Data(11, 1).Tate = 9

Data(11, 2).Name = "�̷�ͷ"
Data(11, 2).Hp = 10
Data(11, 2).Shuriken = 4
Data(11, 2).Flame = 100
Data(11, 2).Side = 1.5
Data(11, 2).Back = 3
Data(11, 2).Tate = 9

Data(11, 3).Name = "��ʽ��"
Data(11, 3).Hp = 30
Data(11, 3).Shuriken = 1.5
Data(11, 3).Flame = 37.5
Data(11, 3).Side = 1.5
Data(11, 3).Back = 1.5
Data(11, 3).Tate = 8

Data(11, 5).Name = "�̷�ͷ��BOSS��"
Data(11, 5).Hp = 10
Data(11, 5).Shuriken = 4
Data(11, 5).Flame = 100
Data(11, 5).Side = 1.5
Data(11, 5).Back = 3
Data(11, 5).Tate = 7

Data(11, 6).Name = "�̣�BOSS��"
Data(11, 6).Hp = 450
Data(11, 6).Shuriken = 2
Data(11, 6).Flame = 50
Data(11, 6).Side = 1.5
Data(11, 6).Back = 3
Data(11, 6).Tate = 8

Rem 6B
Data(12, 5).Name = "��ͷ��BOSS��"
Data(12, 5).Hp = 10
Data(12, 5).Shuriken = 3
Data(12, 5).Flame = 75
Data(12, 5).Side = 2
Data(12, 5).Back = 3
Data(12, 5).Tate = 8

Data(12, 6).Name = "��������BOSS��"
Data(12, 6).Hp = 675
Data(12, 6).Shuriken = 0
Data(12, 6).Flame = 37.5
Data(12, 6).Side = 1
Data(12, 6).Back = 1
Data(12, 6).Tate = 9

Rem 7A
Data(13, 1).Name = "��Ů��"
Data(13, 1).Hp = 37.5
Data(13, 1).Shuriken = 3
Data(13, 1).Flame = 75
Data(13, 1).Side = 2
Data(13, 1).Back = 3
Data(13, 1).Tate = 9

Data(13, 2).Name = "��ͷ"
Data(13, 2).Hp = 15
Data(13, 2).Shuriken = 3
Data(13, 2).Flame = 75
Data(13, 2).Side = 2
Data(13, 2).Back = 3
Data(13, 2).Tate = 9

Data(13, 3).Name = "ǿ��ѻ�칷"
Data(13, 3).Hp = 150
Data(13, 3).Shuriken = 1.5
Data(13, 3).Flame = 37.5
Data(13, 3).Side = 1.5
Data(13, 3).Back = 3
Data(13, 3).Tate = 8

Data(13, 4).Name = "��¯��BOSS��"
Data(13, 4).Hp = 50
Data(13, 4).Shuriken = 2
Data(13, 4).Flame = 50
Data(13, 4).Side = 1
Data(13, 4).Back = 1
Data(13, 4).Tate = 9

Data(13, 5).Name = "����BOSS��"
Data(13, 5).Hp = 12.5
Data(13, 5).Shuriken = 3
Data(13, 5).Flame = 75
Data(13, 5).Side = 2
Data(13, 5).Back = 3
Data(13, 5).Tate = 8

Data(13, 6).Name = "���У�BOSS��"
Data(13, 6).Hp = 525
Data(13, 6).Shuriken = 2
Data(13, 6).Flame = 50
Data(13, 6).Side = 2
Data(13, 6).Back = 3
Data(13, 6).Tate = 9

Rem 7B
Data(14, 5).Name = "С����BOSS��"
Data(14, 5).Hp = 12.5
Data(14, 5).Shuriken = 1
Data(14, 5).Flame = 25
Data(14, 5).Side = 1.5
Data(14, 5).Back = 3
Data(14, 5).Tate = 8

Data(14, 6).Name = "��������BOSS��"
Data(14, 6).Hp = 750
Data(14, 6).Shuriken = 0
Data(14, 6).Flame = 50
Data(14, 6).Side = 1.5
Data(14, 6).Back = 3
Data(14, 6).Tate = 9

Rem 8A
Data(15, 1).Name = "����"
Data(15, 1).Hp = 20
Data(15, 1).Shuriken = 2
Data(15, 1).Flame = 50
Data(15, 1).Side = 1.5
Data(15, 1).Back = 3
Data(15, 1).Tate = 9

Data(15, 2).Name = "��ͷ"
Data(15, 2).Hp = 15
Data(15, 2).Shuriken = 3
Data(15, 2).Flame = 75
Data(15, 2).Side = 2
Data(15, 2).Back = 3
Data(15, 2).Tate = 9

Data(15, 3).Name = "ǿ����ʽ��"
Data(15, 3).Hp = 45
Data(15, 3).Shuriken = 1.5
Data(15, 3).Flame = 37.5
Data(15, 3).Side = 1.5
Data(15, 3).Back = 1.5
Data(15, 3).Tate = 9

Data(15, 4).Name = "��ͷ��BOSS��"
Data(15, 4).Hp = 15
Data(15, 4).Shuriken = 3
Data(15, 4).Flame = 75
Data(15, 4).Side = 2
Data(15, 4).Back = 3
Data(15, 4).Tate = 9

Data(15, 5).Name = "ǿ����ʽ��BOSS��"
Data(15, 5).Hp = 45
Data(15, 5).Shuriken = 1.5
Data(15, 5).Flame = 37.5
Data(15, 5).Side = 1.5
Data(15, 5).Back = 1.5
Data(15, 5).Tate = 9

Data(15, 6).Name = "�ţ�BOSS��"
Data(15, 6).Hp = 300
Data(15, 6).Shuriken = 0
Data(15, 6).Flame = 50
Data(15, 6).Side = 0
Data(15, 6).Back = 0
Data(15, 6).Tate = 9

Rem 8B
Data(16, 1).Name = "����"
Data(16, 1).Hp = 60
Data(16, 1).Shuriken = 2
Data(16, 1).Flame = 50
Data(16, 1).Side = 1.5
Data(16, 1).Back = 3
Data(16, 1).Tate = 9

Data(16, 2).Name = "���ͷ"
Data(16, 2).Hp = 15
Data(16, 2).Shuriken = 2
Data(16, 2).Flame = 50
Data(16, 2).Side = 1.5
Data(16, 2).Back = 3
Data(16, 2).Tate = 9

Data(16, 3).Name = "ǿ����ʽ��"
Data(16, 3).Hp = 45
Data(16, 3).Shuriken = 1.5
Data(16, 3).Flame = 37.5
Data(16, 3).Side = 1.5
Data(16, 3).Back = 1.5
Data(16, 3).Tate = 9

Data(16, 5).Name = "����BOSS��"
Data(16, 5).Hp = 15
Data(16, 5).Shuriken = 3
Data(16, 5).Flame = 75
Data(16, 5).Side = 2
Data(16, 5).Back = 5
Data(16, 5).Tate = 8

Data(16, 6).Name = "�b���λ���BOSS��"
Data(16, 6).Hp = 900
Data(16, 6).Shuriken = 0
Data(16, 6).Flame = 50
Data(16, 6).Side = 1.5
Data(16, 6).Back = 3
Data(16, 6).Tate = 9

Rem EX
Data(0, 0).Name = "���"
Data(0, 0).Hp = 105
Data(0, 0).Shuriken = 1.5
Data(0, 0).Flame = 37.5
Data(0, 0).Side = 2
Data(0, 0).Back = 3
Data(0, 0).Tate = 7

Rem Tate
Tatename(1) = "�R"
Tatename(2) = "��"
Tatename(3) = "�L"
Tatename(4) = "��"
Tatename(5) = "��"
Tatename(6) = "�"
Tatename(7) = "��"
Tatename(8) = "��"
Tatename(9) = "ǰ"

Rem Stage
Stage(1) = "STAGE 1-A ҡ��"
Stage(2) = "STAGE 1-B ��܊"
Stage(3) = "STAGE 2-A �_�"
Stage(4) = "STAGE 2-B ����"
Stage(5) = "STAGE 3-A ���"
Stage(6) = "STAGE 3-B ��ؑ"
Stage(7) = "STAGE 4-A ���"
Stage(8) = "STAGE 4-B ����"
Stage(9) = "STAGE 5-A ��C"
Stage(10) = "STAGE 5-B »��"
Stage(11) = "STAGE 6-A ���"
Stage(12) = "STAGE 6-B ����"
Stage(13) = "STAGE 7-A ؝��"
Stage(14) = "STAGE 7-B ���"
Stage(15) = "STAGE 8-A ����"
Stage(16) = "STAGE 8-B ̫һ"
Stage(17) = "STAGE EX"



iniFileName = App.Path + "\Shinobi.ini"
For x = 0 To 16
    For y = 0 To 6
        sx = x
        sy = y
        Sname(x, y) = "Data(" + sx + ", " + sy + ").Name"
        i = GetPrivateProfileString("Shinobi", Sname(x, y), "", Buff, Len(Buff), iniFileName)
        p = InStr(Buff, Chr(0))
        If p <> 1 Then Data(x, y).Name = Trim(Left(Buff, p - 1))
    Next y
Next x
    
For x = 1 To 9
    sx = x
    State(x) = "Tate(" + sx + ")"
    i = GetPrivateProfileString("Shinobi", State(x), "", Buff, Len(Buff), iniFileName)
    p = InStr(Buff, Chr(0))
    If p <> 1 Then Tatename(x) = Trim(Left(Buff, p - 1))
Next x

For y = 1 To 17
    sy = y
    Stagename(y) = "Stage(" + sy + ")"
    i = GetPrivateProfileString("Shinobi", Stagename(y), "", Buff, Len(Buff), iniFileName)
    p = InStr(Buff, Chr(0))
    If p <> 1 Then Stage(y) = Trim(Left(Buff, p - 1))
Next y
End Sub
