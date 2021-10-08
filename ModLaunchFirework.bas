Attribute VB_Name = "ModLaunchFirework"
Attribute VB_Name = "ModLaunchFirework"
Option Explicit

'LaunchFirework                �E�E�E���ꏊ�FVBAProject.�Ζ�S�̗���
'�Ζ�S�̗����O�ՃA�j���[�V�����E�E�E���ꏊ�FVBAProject.�Ζ�S�̗���
'�����O�Ռv�Z2                 �E�E�E���ꏊ�FVBAProject.�����O��
'�Ζ���W�擾                  �E�E�E���ꏊ�FVBAProject.�Ζ���W
'ExtractRowArray               �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2D                  �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1            �E�E�E���ꏊ�FFukamiAddins3.ModArray
'ExtractArray                  �E�E�E���ꏊ�FFukamiAddins3.ModArray
'DrawPolyLine                  �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'DrawPolyLineAddPoint          �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'GetXYDocumentFromCursor       �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetXYCellScreenUpperLeft      �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetPaneOfCell                 �E�E�E���ꏊ�FFukamiAddins3.ModCursor
'GetXYCellScreenLowerRight     �E�E�E���ꏊ�FFukamiAddins3.ModCursor

'------------------------------


'�~���b�P�ʂŎ��Ԓ�~
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'------------------------------

Public Const G  As Single = 9.80665 '�d�͉����x
Public Const Pi As Single = 3.14159265358979 '�~����
'------------------------------
'------------------------------
'�z��̏����֌W�̃v���V�[�W��
'------------------------------
'�V�F�C�v��}�֘A���W���[��
'20210914�쐬
'------------------------------


'������������������������������������������������������
'�J�[�\���̃X�N���[�����W�擾�p
#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#Else
Private Declare Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#End If

Private Type PointAPI
    X As Long
    Y As Long
End Type

'������������������������������������������������������
'�X�N���[���̃T�C�Y�擾�p
#If VBA7 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXScreen As Long = 0
Private Const SM_CYScreen As Long = 1

'������������������������������������������������������
'DPI�Ƃ��擾�p
#If VBA7 Then
' ��GetDC(API)
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
' ��ReleaseDC(API)
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" _
    (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
' ��GetDeviceCaps(API)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" _
    (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
' ��GetDC(API)
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
' ��ReleaseDC(API)
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
' ��GetDeviceCaps(API)
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

'������������������������������������������������������

'------------------------------

Sub LaunchFirework(TargetSheet As Worksheet)
'�V�[�g�̃J�[�\���ʒu�ɉԉ΂�ł��グ��
'���[�N�V�[�g��(Worksheet_SelectionChange)�C�x���g�œ��삳����
'20211008

'����
'TargetSheet�E�E�E�ԉ΂�ł��グ��Ώۂ̃V�[�g

    '�}�E�X�J�[�\���̃h�L�������g���W�擾
    Dim CenterX#, CenterZ#
    Dim Dummy
    Dummy = GetXYDocumentFromCursor '�}�E�X�J�[�\���̃h�L�������g���W�擾
    
    CenterX = Dummy(1) 'X����(���E����)
    CenterZ = Dummy(2) 'Z����(��������)
    
    Dim Core_R#, Kayaku_R#, V0#, InputRGB&, Wind#
    Core_R = 0.1 '�ԉ΂̊j���a[m]
    Kayaku_R = 0.015 + 0.02 * Rnd() '�Ζ��̔��a[m]
    V0 = 50 + 10 * Rnd() '�ԉ΂���������Ƃ��̉Ζ�̏���[m/s]
    InputRGB = RGB(83 + (255 - 83) * Rnd(), 83 + (255 - 83) * Rnd(), 83 + (255 - 83) * Rnd()) '�ԉ΂̐F
    Wind = 10 * Rnd() '���̑���[m/s]
    
    Application.EnableEvents = False
    Call �Ζ�S�̗����O�ՃA�j���[�V����(CenterX, CenterZ, Core_R, Kayaku_R, V0, InputRGB, Wind, TargetSheet)
    Application.EnableEvents = True

End Sub

Private Sub �Ζ�S�̗����O�ՃA�j���[�V����(CenterX#, CenterZ#, Core_R#, Kayaku_R#, V0#, InputRGB&, Wind#, TargetSheet As Worksheet)
'�ԉ΂��Č������A�j���[�V�������s
'20211008

'����
'CenterX    �E�E�E�ԉ΂�����������WX[m]/Double�^
'CenterZ    �E�E�E�ԉ΂�����������WZ[m]/Double�^
'Core_R     �E�E�E�ԉ΂̊e���a[m]/Double�^
'Kayaku_R   �E�E�E�Ζ�̔��a[m]/Double�^
'V0         �E�E�E�Ζ�̏���[m/s]/Double�^
'InputRGB   �E�E�E�ԉ΂̐F/Long�^
'Wind       �E�E�E����[m/s]/Double�^
'TargetSheet�E�E�E�`��Ώۂ̃V�[�g/Worksheet�^
    
    Dim I&, J&, II&, JJ&, M&, K&
    
    '��{���l�ݒ�
    Dim N&, dt#, Ox#, Oy#, Oz#, X0#, Y0#, Z0#, PointCount&
    N = 30         '�O�Ս�}�̓_��/�����قǉԉ΂̋O�Ղ������Ȃ�
    dt = 0.5       '�O�ՊԊu�̎��ԕω�/�傫���قǕ`��O�Ղ̒����������Ȃ�
    Ox = 0         '�ԉ΂̍ŏ��̍��WX
    Oy = 0         '�ԉ΂̍ŏ��̍��WY
    Oz = 0         '�ԉ΂̍ŏ��̍��WZ
    PointCount = 5 '�O�Օ`��̃|�C���g��/�����قǂȂ߂炩�ɂȂ邪�v�Z���x���Ȃ�
    
    '�ŏ��̉Ζ�̍��W���v�Z����B
    Dim KayakuZahyoList
    KayakuZahyoList = �Ζ���W�擾(Core_R, Kayaku_R)
    
    Dim KayakuCount&
    KayakuCount = UBound(KayakuZahyoList, 1)
    
    '�S�ẲΖ�̗����O�Ղ��v�Z����
    Dim AllKisekiList()
    ReDim AllKisekiList(1 To KayakuCount)
    
    For I = 1 To KayakuCount
        X0 = KayakuZahyoList(I, 1) + Ox
        Y0 = KayakuZahyoList(I, 2) + Oy
        Z0 = KayakuZahyoList(I, 3) + Oz
        AllKisekiList(I) = �����O�Ռv�Z2(N, dt, V0, Ox, Oy, Oz, X0, Y0, Z0, Wind, PointCount)
    Next I
    
    '�e�Ζ�̋O�Ղ̃V�F�C�v���ݒ�
    Dim ShapeNameList() As String
    ReDim ShapeNameList(1 To KayakuCount)
    
    Dim IdeNum&, IdeStr$
    IdeNum = WorksheetFunction.RandBetween(1, 9999)
    IdeStr = "�Ζ�" & Format(IdeNum, "0000")
    
    For I = 1 To KayakuCount
        ShapeNameList(I) = IdeStr & Format(I, "0000")
    Next I
    
    '��}�A�j���[�V����
    Dim TmpKisekiList
    Dim TmpShape As Shape
    Dim TmpShapeName As String
    Dim TmpSakuzuKiseki
        
    Dim TmpTimer#, MaxSleepTime#, TmpSleepTime#
    
    MaxSleepTime = 0.2 '�ő��~����(�A�j���[�V�������x�����ɂ��邽��)������������������������������������������������
    
    TmpTimer = Timer '�v�Z���x�v���p
    For I = 1 To N
        For J = 1 To KayakuCount
            TmpKisekiList = AllKisekiList(J) '���̉Ζ�̑S�O��
            TmpShapeName = ShapeNameList(J) '���̉Ζ�̋O�Ղ̃V�F�C�v��
            If I = 1 Then '��}�̍ŏ��̏ꍇ
                TmpSakuzuKiseki = ExtractArray(TmpKisekiList, 1, 1, PointCount, 2) '�ŏ��̋O�Ր��_�𔲂��o��
                
                For II = 1 To PointCount '�J�[�\���ʒu�Ɉړ�����
                    TmpSakuzuKiseki(II, 1) = TmpSakuzuKiseki(II, 1) + CenterX
                    TmpSakuzuKiseki(II, 2) = TmpSakuzuKiseki(II, 2) + CenterZ
                Next
                
                Set TmpShape = DrawPolyLine(TmpSakuzuKiseki, TargetSheet) '�|�����C����}
                TmpShape.Name = TmpShapeName '�V�F�C�v���ݒ�
                TmpShape.Line.ForeColor.RGB = InputRGB '�ԉ΂̐F�ݒ�
            Else
                '��}2��ڈȍ~
                TmpSakuzuKiseki = ExtractRowArray(TmpKisekiList, I + 2) '���̍�}�_���o
                Set TmpShape = TargetSheet.Shapes(TmpShapeName) '���̉Ζ�̋O�Ղ̃V�F�C�v�擾
                Call DrawPolyLineAddPoint(TmpShape, CenterX + TmpSakuzuKiseki(1), CenterZ + TmpSakuzuKiseki(2)) '�_��ǉ����ċO�Ղ�����
                
            End If
        
        Next J
        
        '�A�j���[�V�����p����
        TmpSleepTime = MaxSleepTime - (Timer - TmpTimer)
        TmpSleepTime = WorksheetFunction.Max(0, TmpSleepTime)
        Debug.Print Format(Timer - TmpTimer, "0.00000�b"), "��~����" & TmpSleepTime '�v�Z���x�m�F�o��
        TmpTimer = Timer
        Sleep TmpSleepTime * 100
        Application.Calculate
        DoEvents
        
    Next I
    
    '�I�����ɋO�Ղ̐���S������
    For I = 1 To KayakuCount
        TmpShapeName = ShapeNameList(I)
        TargetSheet.Shapes(TmpShapeName).Delete
    Next
    
End Sub

Private Function �����O�Ռv�Z2(N&, dt#, V0#, Ox#, Oy#, Oz#, X0#, Y0#, Z0#, Wind#, PointCount&)
'�w��Ζ�̗����O�Ղ̑S���W���v�Z����B
'20211008

'����
'N         �E�E�E�O�Ղ̓_�̌�/Long�^
'dt        �E�E�E�O�Ղ̓_�̊Ԃ̎��ԕω�/Double�^
'V0        �E�E�E�ԉΔ����̏���/Double�^
'Ox        �E�E�E�ԉΔ����ʒu�̍ŏ��̈ʒu���WX/Double�^
'Oy        �E�E�E�ԉΔ����ʒu�̍ŏ��̈ʒu���WY/Double�^
'Oz        �E�E�E�ԉΔ����ʒu�̍ŏ��̈ʒu���WZ/Double�^
'X0        �E�E�E�Ζ�̉ԉΒ��S�ɑ΂��Ă̑��΍��WX/Double�^
'Y0        �E�E�E�Ζ�̉ԉΒ��S�ɑ΂��Ă̑��΍��WY/Double�^
'Z0        �E�E�E�Ζ�̉ԉΒ��S�ɑ΂��Ă̑��΍��WZ/Double�^
'Wind      �E�E�E����/Double�^
'PointCount�E�E�E�`��O�Ղ̓_��/Long�^

    Dim Theta#, Fai#, R# '��,��,r
    Dim r_xy# 'XY���e�̔��a
    R = ((X0 - Ox) ^ 2 + (Y0 - Oy) ^ 2 + (Z0 - Oz) ^ 2) ^ (1 / 2)
    Fai = WorksheetFunction.Asin((Z0 - Oz) / R)
    r_xy = R * Cos(Fai)
    If r_xy < 0.000001 Then
        Theta = 0
    Else
        Theta = WorksheetFunction.Atan2((X0 - Ox) / r_xy, (Y0 - Oy) / r_xy)
    End If
    
    Dim Output()
    ReDim Output(1 To N + PointCount, 1 To 2)
    Dim TmpTime#, TmpX#, TmpY#, TmpZ#
    Dim I&, K&
        
    K = 0
    For I = 1 To N + PointCount
        
        If I <= PointCount Then  '�O�Ղ̊J�n�œ_�̐����[���I�Ɍ��炷
            '�������Ȃ�
        ElseIf I >= N + 1 Then
            '�������Ȃ�
        Else
            K = K + 1
        End If
        TmpTime = dt * K
        
        TmpX = X0 + V0 * Cos(Fai) * Cos(Theta) * TmpTime + Wind * TmpTime
        TmpY = Y0 + V0 * Cos(Fai) * Sin(Theta) * TmpTime
        TmpZ = Z0 + V0 * Sin(Fai) * TmpTime - (G / 2) * TmpTime * TmpTime
        
        Output(I, 1) = TmpX
        Output(I, 2) = -TmpZ
        
    Next I
        
    �����O�Ռv�Z2 = Output
        
End Function

Private Function �Ζ���W�擾(R_core#, R_fire#, Optional HaimenNasiNaraTrue = True)
    
    Dim Theta# '���S�p��
    Dim KayakuCountList
    Dim DanCount As Byte
    Dim RkList
    Dim SkList
    Dim S1 As Integer
    Dim KayakuCount As Integer
    
    Theta = 2 * WorksheetFunction.Asin(R_fire / (R_core + R_fire))
    S1 = WorksheetFunction.RoundDown(2 * Pi / Theta, 0) '1�i�ڂ̉Ζ��
    DanCount = WorksheetFunction.RoundUp(S1 / 4, 0) '�i��
    
    Dim I&, K&, J& '�����グ�p(Integer�^)
    
    ReDim RkList(1 To DanCount)
    ReDim SkList(1 To DanCount)
    
    Dim ThetaK#
    
    For K = 1 To DanCount
        RkList(K) = (R_core + R_fire) * Cos(Theta * (K - 1))
        
        If RkList(K) < R_fire Then
            SkList(K) = SkList(K - 1)
        Else
            ThetaK = 2 * WorksheetFunction.Asin(R_fire / RkList(K))
            SkList(K) = WorksheetFunction.RoundDown(2 * Pi / ThetaK, 0)
        End If
    Next K
    
    KayakuCount = WorksheetFunction.Sum(SkList) * 2 - SkList(1)
    
    Dim KayakuZahyoList
    ReDim KayakuZahyoList(1 To KayakuCount, 1 To 3)
    
    Dim TmpRk#, TmpSk#
    
    J = 0
    For K = 1 To DanCount
        TmpSk = SkList(K)
        For I = 1 To TmpSk
            J = J + 1
            KayakuZahyoList(J, 1) = (R_core + R_fire) * Cos(Theta * (K - 1)) * Cos(2 * Pi / TmpSk * (I - 1))
            KayakuZahyoList(J, 2) = (R_core + R_fire) * Cos(Theta * (K - 1)) * Sin(2 * Pi / TmpSk * (I - 1))
            KayakuZahyoList(J, 3) = (R_core + R_fire) * Sin(Theta * (K - 1))
        Next I
        
        If K > 1 Then '�������̕�
            For I = 1 To TmpSk
                J = J + 1
                KayakuZahyoList(J, 1) = (R_core + R_fire) * Cos(Theta * (K - 1)) * Cos(2 * Pi / TmpSk * (I - 1))
                KayakuZahyoList(J, 2) = (R_core + R_fire) * Cos(Theta * (K - 1)) * Sin(2 * Pi / TmpSk * (I - 1))
                KayakuZahyoList(J, 3) = -(R_core + R_fire) * Sin(Theta * (K - 1))
            Next I
        End If
        
    Next K
    
    Dim ZenmenKayakuZahyoList
    ReDim ZenmenKayakuZahyoList(1 To KayakuCount, 1 To 3)
    
    K = 0
    If HaimenNasiNaraTrue Then
        For I = 1 To KayakuCount
            If KayakuZahyoList(I, 2) > 0 Then
                K = K + 1
                ZenmenKayakuZahyoList(K, 1) = KayakuZahyoList(I, 1)
                ZenmenKayakuZahyoList(K, 2) = KayakuZahyoList(I, 2)
                ZenmenKayakuZahyoList(K, 3) = KayakuZahyoList(I, 3)
            End If
        Next I
        
        ReDim KayakuZahyoList(1 To K, 1 To 3)
        For I = 1 To K
            KayakuZahyoList(I, 1) = ZenmenKayakuZahyoList(I, 1)
            KayakuZahyoList(I, 2) = ZenmenKayakuZahyoList(I, 2)
            KayakuZahyoList(I, 3) = ZenmenKayakuZahyoList(I, 3)
        Next I
    End If
    �Ζ���W�擾 = KayakuZahyoList
    
End Function

Private Function ExtractRowArray(Array2D, TargetRow&)
'�񎟌��z��̎w��s���ꎟ���z��Œ��o����
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'TargetRow�E�E�E���o����Ώۂ̍s�ԍ�


    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��

    If TargetRow < 1 Then
        MsgBox ("���o����s�ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf TargetRow > N Then
        MsgBox ("���o����s�ԍ��͌��̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If

    '����
    Dim Output
    ReDim Output(1 To M)
    
    For I = 1 To M
        Output(I) = Array2D(TargetRow, I)
    Next I
    
    '�o��
    ExtractRowArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function ExtractArray(Array2D, StartRow&, StartCol&, EndRow&, EndCol&)
'�񎟌��z��̎w��͈͂�z��Ƃ��Ē��o����
'20210917

'����
'Array2D �E�E�E�񎟌��z��
'StartRow�E�E�E���o�͈͂̊J�n�s�ԍ�
'StartCol�E�E�E���o�͈͂̊J�n��ԍ�
'EndRow  �E�E�E���o�͈͂̏I���s�ԍ�
'EndCol  �E�E�E���o�͈͂̏I����ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If StartRow > EndRow Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v�́A�I���s�uEndRow�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v�́A�I����uEndCol�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("���o�͈͂̏I���s�uStartRow�v�͒��o���̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("���o�͈͂̏I����uStartCol�v�͒��o���̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '�o��
    ExtractArray = Output
    
End Function

Private Function DrawPolyLine(XYList, TargetSheet As Worksheet) As Shape
'XY���W����|�����C����`��
'�V�F�C�v���I�u�W�F�N�g�ϐ��Ƃ��ĕԂ�
'20210921

'����
'XYList         �E�E�EXY���W���������񎟌��z�� X�������E���� Y������������
'TargetSheet    �E�E�E��}�Ώۂ̃V�[�g

    Dim I%, Count%
    Count = UBound(XYList, 1)
    
    With TargetSheet.Shapes.BuildFreeform(msoEditingCorner, XYList(1, 1), XYList(1, 2))
        
        For I = 2 To Count
            .AddNodes msoSegmentLine, msoEditingAuto, XYList(I, 1), XYList(I, 2)
        Next I
        Set DrawPolyLine = .ConvertToShape
    End With
    
End Function

Private Sub DrawPolyLineAddPoint(InputShape As Shape, AddX#, AddY#, Optional DeleteFirstPoint As Boolean = True)
'�|�����C���ɓ_��ǉ����ĉ�������
'20211008

'����
'InputShape         �E�E�E�Ώۂ̃|�����C��
'AddX               �E�E�E�ǉ�����_��X���W�i�E�����j
'AddY               �E�E�E�ǉ�����_��Y���W�i�������j
'[DeleteFirstPoint] �E�E�E�Ώۂ̋Ȑ��̍ŏ��̓_���폜���邩�ǂ���


    Dim TmpNode As ShapeNodes
    Set TmpNode = InputShape.Nodes
    
    With TmpNode
        .Insert .Count, msoSegmentLine, msoEditingCorner, AddX, AddY
    End With
    
    If DeleteFirstPoint Then
        TmpNode.Delete 1
    End If
    
End Sub

Private Function GetXYDocumentFromCursor(Optional ImmidiateShow As Boolean = True)
'���݃J�[�\���ʒu�̃h�L�������g���W�擾
'�J�[�\���ʒu�̃X�N���[�����W���A
'�J�[�\��������Ă���Z���̎l���̃X�N���[�����W�̊֌W������A
'�J�[�\��������Ă���Z���̎l���̃h�L�������g���W�����Ƃɕ�Ԃ��āA
'�J�[�\���ʒu�̃h�L�������g���W�����߂�B
'20211005

'����
'[ImmidiateShow]�E�E�E�C�~�f�B�G�C�g�E�B���h�E�Ɍv�Z���ʂȂǂ�\�����邩(�f�t�H���g��True)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�J�[�\���ʒu�̃h�L�������g���WX,2:�J�[�\���ʒu�̃h�L�������g���WY(Double�^)
'�J�[�\�����V�[�g���ɂȂ��ꍇ��Empty��Ԃ��B

'�Q�l�Fhttps://gist.github.com/furyutei/f0668f33d62ccac95d1643f15f19d99a?s=09#to-footnote-1

    Dim Win As Window
    Set Win = ActiveWindow
    
    '�J�[�\���̃X�N���[�����W�擾
    Dim Cursor As PointAPI, CursorScreenX#, CursorScreenY#
    Call GetCursorPos(Cursor)
    CursorScreenX = Cursor.X
    CursorScreenY = Cursor.Y
    
    '�J�[�\��������Ă���Z�����擾
    Dim CursorCell As Range, Dummy
    Set Dummy = Win.RangeFromPoint(CursorScreenX, CursorScreenY)
    If TypeName(Dummy) = "Range" Then
        Set CursorCell = Dummy
    Else
        '�J�[�\�����Z���ɏ���ĂȂ��̂ŏI��
        Exit Function
    End If
    
    '�l���̃X�N���[�����W���擾
    Dim X1Screen#, X2Screen#, Y1Screen#, Y2Screen# '�l���̃X�N���[�����W
    Dummy = GetXYCellScreenUpperLeft(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X1Screen = Dummy(1)
    Y1Screen = Dummy(2)
    
    Dummy = GetXYCellScreenLowerRight(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X2Screen = Dummy(1)
    Y2Screen = Dummy(2)
    
    '�l���̃h�L�������g���W�擾
    Dim X1Document#, X2Document#, Y1Document#, Y2Document# '�l���̃h�L�������g���W
    X1Document = CursorCell.Left
    X2Document = CursorCell.Left + CursorCell.Width
    Y1Document = CursorCell.Top
    Y2Document = CursorCell.Top + CursorCell.Height
    
    '�}�E�X�J�[�\���̃h�L�������g���W���ԂŌv�Z
    Dim CursorDocumentX#, CursorDocumentY#
    CursorDocumentX = X1Document + (X2Document - X1Document) * (CursorScreenX - X1Screen) / (X2Screen - X1Screen)
    CursorDocumentY = Y1Document + (Y2Document - Y1Document) * (CursorScreenY - Y1Screen) / (Y2Screen - Y1Screen)
        
    '�o��
    Dim Output#(1 To 2)
    Output(1) = CursorDocumentX
    Output(2) = CursorDocumentY
    
    GetXYDocumentFromCursor = Output
    
    '�m�F�\��
    If ImmidiateShow Then
        Debug.Print "�J�[�\���̏�����Z��", CursorCell.Address(False, False)
        Debug.Print "�J�[�\���X�N���[�����W", "CursorScreenX:" & CursorScreenX, "CursorScreenY:" & CursorScreenY
        Debug.Print "�J�[�\���h�L�������g���W", "CursorDocumentX:" & WorksheetFunction.Round(CursorDocumentX, 1), "CursorDocumentY:" & WorksheetFunction.Round(CursorDocumentY, 1)
        Debug.Print "�Z������X�N���[�����W", "X1Screen:" & X1Screen, , "Y1Screen:" & Y1Screen
        Debug.Print "�Z������h�L�������g���W", "X1Document:" & X1Document, "Y1Document:" & Y1Document
        Debug.Print "�Z���E���X�N���[�����W", "X2Screen:" & X2Screen, , "Y2Screen:" & Y2Screen
        Debug.Print "�Z���E���h�L�������g���W", "X2Document:" & X2Document, "Y2Document:" & Y2Document
    End If

End Function

Private Function GetXYCellScreenUpperLeft(TargetCell As Range)
'�w��Z���̍���̃X�N���[�����WXY���擾����B
'20211005

'����
'TargetCell�E�E�E�Ώۂ̃Z��(Range�^)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�Z������̃X�N���[�����WX,2;�Z������̃X�N���[�����WY(Double�^)

    '�Z�����\������Ă���Pane(�E�B���h�E�g�̌Œ���l�������\���G���A)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
       
    '�yPointsToScreenPixels�̒��ӎ����z
    '�y���z�ΏۃZ�����V�[�g��ŕ\������Ă��Ȃ��Ǝ擾�s�B�ꕔ�ł��\������Ă�����擾�\�B
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top)
    
    GetXYCellScreenUpperLeft = Output
    
End Function

Private Function GetPaneOfCell(TargetCell As Range) As Pane
'�w��Z����Pane���擾����
'�E�B���h�E�g�Œ�A�E�B���h�E�����̐ݒ�ł��擾�ł���B
'�Q�l�Fhttp://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_100_120.html
'20211006

'����
'TargetCell�E�E�E�Ώۂ̃Z��/Range�^

'�Ԃ�l
'�w��Z�����܂܂��Pane/Pane�^
'�w��Z�����\���͈͊O�Ȃ�Nothing
    
    Dim Win As Window
    Set Win = ActiveWindow
    
    Dim Output As Pane
    Dim I& '�����グ�p(Long�^)
    
    ' �E�B���h�E����������
    If Not Win.FreezePanes And Not Win.Split Then
        '�E�B���h�E�g�Œ�ł��E�B���h�E�����ł��Ȃ��ꍇ
        ' �\���ȊO�ɃZ��������ꍇ�͖���
        If Intersect(Win.VisibleRange, TargetCell) Is Nothing Then Exit Function
        Set Output = Win.Panes(1)
    Else ' ��������
        If Win.FreezePanes Then
            ' �E�B���h�E�g�Œ�̏ꍇ
            ' �ǂ̃E�B���h�E�ɑ����邩����
            For I = 1 To Win.Panes.Count
                If Not Intersect(Win.Panes(I).VisibleRange, TargetCell) Is Nothing Then
                    'Pane�̕\���͈͂Ɋ܂܂��ꍇ�͂���Pane���擾
                    Set Output = Win.Panes(I)
                    Exit For
                End If
            Next I
            
            '������Ȃ������ꍇ
            If Output Is Nothing Then Exit Function
        Else
            '�E�B���h�E�����̏ꍇ
            ' �E�B���h�E�����̓A�N�e�B�u�y�C���̂ݔ���
            If Not Intersect(Win.ActivePane.VisibleRange, TargetCell) Is Nothing Then
                Set Output = Win.ActivePane
            Else
                Exit Function
            End If
        End If
    End If
    
    '�o��
    Set GetPaneOfCell = Output
    
End Function

Private Function GetXYCellScreenLowerRight(TargetCell As Range)
'�w��Z���̉E���̃X�N���[�����WXY���擾����B
'20211005

'����
'TargetCell�E�E�E�Ώۂ̃Z��(Range�^)

'�Ԃ�l
'Output(1 to 2)�E�E�E1:�Z���E���̃X�N���[�����WX,2;�Z���E���̃X�N���[�����WY(Double�^)

    '�Z�����\������Ă���Pane(�E�B���h�E�g�̌Œ���l�������\���G���A)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
    
    '�yPointsToScreenPixels�̒��ӎ����z
    '�y���z�ΏۃZ�����V�[�g��ŕ\������Ă��Ȃ��Ǝ擾�s�B�ꕔ�ł��\������Ă�����擾�\�B
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left + TargetCell.Width)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top + TargetCell.Height)
    
    GetXYCellScreenLowerRight = Output
    
End Function


