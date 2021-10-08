Attribute VB_Name = "ModLaunchFirework"
Attribute VB_Name = "ModLaunchFirework"
Option Explicit

'LaunchFirework                ・・・元場所：VBAProject.火薬全体落下
'火薬全体落下軌跡アニメーション・・・元場所：VBAProject.火薬全体落下
'落下軌跡計算2                 ・・・元場所：VBAProject.落下軌跡
'火薬座標取得                  ・・・元場所：VBAProject.火薬座標
'ExtractRowArray               ・・・元場所：FukamiAddins3.ModArray
'CheckArray2D                  ・・・元場所：FukamiAddins3.ModArray
'CheckArray2DStart1            ・・・元場所：FukamiAddins3.ModArray
'ExtractArray                  ・・・元場所：FukamiAddins3.ModArray
'DrawPolyLine                  ・・・元場所：FukamiAddins3.ModDrawShape
'DrawPolyLineAddPoint          ・・・元場所：FukamiAddins3.ModDrawShape
'GetXYDocumentFromCursor       ・・・元場所：FukamiAddins3.ModCursor
'GetXYCellScreenUpperLeft      ・・・元場所：FukamiAddins3.ModCursor
'GetPaneOfCell                 ・・・元場所：FukamiAddins3.ModCursor
'GetXYCellScreenLowerRight     ・・・元場所：FukamiAddins3.ModCursor

'------------------------------


'ミリ秒単位で時間停止
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'------------------------------

Public Const G  As Single = 9.80665 '重力加速度
Public Const Pi As Single = 3.14159265358979 '円周率
'------------------------------
'------------------------------
'配列の処理関係のプロシージャ
'------------------------------
'シェイプ作図関連モジュール
'20210914作成
'------------------------------


'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'カーソルのスクリーン座標取得用
#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#Else
Private Declare Function GetCursorPos Lib "user32" (IpPoint As PointAPI) As Long
#End If

Private Type PointAPI
    X As Long
    Y As Long
End Type

'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'スクリーンのサイズ取得用
#If VBA7 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXScreen As Long = 0
Private Const SM_CYScreen As Long = 1

'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'DPIとか取得用
#If VBA7 Then
' ■GetDC(API)
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
' ■ReleaseDC(API)
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" _
    (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
' ■GetDeviceCaps(API)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" _
    (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
' ■GetDC(API)
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
' ■ReleaseDC(API)
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
' ■GetDeviceCaps(API)
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

'※※※※※※※※※※※※※※※※※※※※※※※※※※※

'------------------------------

Sub LaunchFirework(TargetSheet As Worksheet)
'シートのカーソル位置に花火を打ち上げる
'ワークシートの(Worksheet_SelectionChange)イベントで動作させる
'20211008

'引数
'TargetSheet・・・花火を打ち上げる対象のシート

    'マウスカーソルのドキュメント座標取得
    Dim CenterX#, CenterZ#
    Dim Dummy
    Dummy = GetXYDocumentFromCursor 'マウスカーソルのドキュメント座標取得
    
    CenterX = Dummy(1) 'X方向(左右方向)
    CenterZ = Dummy(2) 'Z方向(高さ方向)
    
    Dim Core_R#, Kayaku_R#, V0#, InputRGB&, Wind#
    Core_R = 0.1 '花火の核半径[m]
    Kayaku_R = 0.015 + 0.02 * Rnd() '火薬一個の半径[m]
    V0 = 50 + 10 * Rnd() '花火が爆発するときの火薬の初速[m/s]
    InputRGB = RGB(83 + (255 - 83) * Rnd(), 83 + (255 - 83) * Rnd(), 83 + (255 - 83) * Rnd()) '花火の色
    Wind = 10 * Rnd() '風の速さ[m/s]
    
    Application.EnableEvents = False
    Call 火薬全体落下軌跡アニメーション(CenterX, CenterZ, Core_R, Kayaku_R, V0, InputRGB, Wind, TargetSheet)
    Application.EnableEvents = True

End Sub

Private Sub 火薬全体落下軌跡アニメーション(CenterX#, CenterZ#, Core_R#, Kayaku_R#, V0#, InputRGB&, Wind#, TargetSheet As Worksheet)
'花火を再現したアニメーション実行
'20211008

'引数
'CenterX    ・・・花火が爆発する座標X[m]/Double型
'CenterZ    ・・・花火が爆発する座標Z[m]/Double型
'Core_R     ・・・花火の各半径[m]/Double型
'Kayaku_R   ・・・火薬の半径[m]/Double型
'V0         ・・・火薬の初速[m/s]/Double型
'InputRGB   ・・・花火の色/Long型
'Wind       ・・・風速[m/s]/Double型
'TargetSheet・・・描画対象のシート/Worksheet型
    
    Dim I&, J&, II&, JJ&, M&, K&
    
    '基本数値設定
    Dim N&, dt#, Ox#, Oy#, Oz#, X0#, Y0#, Z0#, PointCount&
    N = 30         '軌跡作図の点個数/多いほど花火の軌跡が長くなる
    dt = 0.5       '軌跡間隔の時間変化/大きいほど描画軌跡の長さが長くなる
    Ox = 0         '花火の最初の座標X
    Oy = 0         '花火の最初の座標Y
    Oz = 0         '花火の最初の座標Z
    PointCount = 5 '軌跡描画のポイント数/多いほどなめらかになるが計算が遅くなる
    
    '最初の火薬の座標を計算する。
    Dim KayakuZahyoList
    KayakuZahyoList = 火薬座標取得(Core_R, Kayaku_R)
    
    Dim KayakuCount&
    KayakuCount = UBound(KayakuZahyoList, 1)
    
    '全ての火薬の落下軌跡を計算する
    Dim AllKisekiList()
    ReDim AllKisekiList(1 To KayakuCount)
    
    For I = 1 To KayakuCount
        X0 = KayakuZahyoList(I, 1) + Ox
        Y0 = KayakuZahyoList(I, 2) + Oy
        Z0 = KayakuZahyoList(I, 3) + Oz
        AllKisekiList(I) = 落下軌跡計算2(N, dt, V0, Ox, Oy, Oz, X0, Y0, Z0, Wind, PointCount)
    Next I
    
    '各火薬の軌跡のシェイプ名設定
    Dim ShapeNameList() As String
    ReDim ShapeNameList(1 To KayakuCount)
    
    Dim IdeNum&, IdeStr$
    IdeNum = WorksheetFunction.RandBetween(1, 9999)
    IdeStr = "火薬" & Format(IdeNum, "0000")
    
    For I = 1 To KayakuCount
        ShapeNameList(I) = IdeStr & Format(I, "0000")
    Next I
    
    '作図アニメーション
    Dim TmpKisekiList
    Dim TmpShape As Shape
    Dim TmpShapeName As String
    Dim TmpSakuzuKiseki
        
    Dim TmpTimer#, MaxSleepTime#, TmpSleepTime#
    
    MaxSleepTime = 0.2 '最大停止時間(アニメーション速度を一定にするため)←←←←←←←←←←←←←←←←←←←←←←←←
    
    TmpTimer = Timer '計算速度計測用
    For I = 1 To N
        For J = 1 To KayakuCount
            TmpKisekiList = AllKisekiList(J) 'その火薬の全軌跡
            TmpShapeName = ShapeNameList(J) 'その火薬の軌跡のシェイプ名
            If I = 1 Then '作図の最初の場合
                TmpSakuzuKiseki = ExtractArray(TmpKisekiList, 1, 1, PointCount, 2) '最初の軌跡数点を抜き出す
                
                For II = 1 To PointCount 'カーソル位置に移動する
                    TmpSakuzuKiseki(II, 1) = TmpSakuzuKiseki(II, 1) + CenterX
                    TmpSakuzuKiseki(II, 2) = TmpSakuzuKiseki(II, 2) + CenterZ
                Next
                
                Set TmpShape = DrawPolyLine(TmpSakuzuKiseki, TargetSheet) 'ポリライン作図
                TmpShape.Name = TmpShapeName 'シェイプ名設定
                TmpShape.Line.ForeColor.RGB = InputRGB '花火の色設定
            Else
                '作図2回目以降
                TmpSakuzuKiseki = ExtractRowArray(TmpKisekiList, I + 2) '次の作図点抽出
                Set TmpShape = TargetSheet.Shapes(TmpShapeName) 'その火薬の軌跡のシェイプ取得
                Call DrawPolyLineAddPoint(TmpShape, CenterX + TmpSakuzuKiseki(1), CenterZ + TmpSakuzuKiseki(2)) '点を追加して軌跡を延長
                
            End If
        
        Next J
        
        'アニメーション用動作
        TmpSleepTime = MaxSleepTime - (Timer - TmpTimer)
        TmpSleepTime = WorksheetFunction.Max(0, TmpSleepTime)
        Debug.Print Format(Timer - TmpTimer, "0.00000秒"), "停止時間" & TmpSleepTime '計算速度確認出力
        TmpTimer = Timer
        Sleep TmpSleepTime * 100
        Application.Calculate
        DoEvents
        
    Next I
    
    '終了時に軌跡の線を全部消す
    For I = 1 To KayakuCount
        TmpShapeName = ShapeNameList(I)
        TargetSheet.Shapes(TmpShapeName).Delete
    Next
    
End Sub

Private Function 落下軌跡計算2(N&, dt#, V0#, Ox#, Oy#, Oz#, X0#, Y0#, Z0#, Wind#, PointCount&)
'指定火薬の落下軌跡の全座標を計算する。
'20211008

'引数
'N         ・・・軌跡の点の個数/Long型
'dt        ・・・軌跡の点の間の時間変化/Double型
'V0        ・・・花火爆発の初速/Double型
'Ox        ・・・花火爆発位置の最初の位置座標X/Double型
'Oy        ・・・花火爆発位置の最初の位置座標Y/Double型
'Oz        ・・・花火爆発位置の最初の位置座標Z/Double型
'X0        ・・・火薬の花火中心に対しての相対座標X/Double型
'Y0        ・・・火薬の花火中心に対しての相対座標Y/Double型
'Z0        ・・・火薬の花火中心に対しての相対座標Z/Double型
'Wind      ・・・風速/Double型
'PointCount・・・描画軌跡の点数/Long型

    Dim Theta#, Fai#, R# 'θ,φ,r
    Dim r_xy# 'XY投影の半径
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
        
        If I <= PointCount Then  '軌跡の開始で点の数を擬似的に減らす
            '何もしない
        ElseIf I >= N + 1 Then
            '何もしない
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
        
    落下軌跡計算2 = Output
        
End Function

Private Function 火薬座標取得(R_core#, R_fire#, Optional HaimenNasiNaraTrue = True)
    
    Dim Theta# '中心角θ
    Dim KayakuCountList
    Dim DanCount As Byte
    Dim RkList
    Dim SkList
    Dim S1 As Integer
    Dim KayakuCount As Integer
    
    Theta = 2 * WorksheetFunction.Asin(R_fire / (R_core + R_fire))
    S1 = WorksheetFunction.RoundDown(2 * Pi / Theta, 0) '1段目の火薬個数
    DanCount = WorksheetFunction.RoundUp(S1 / 4, 0) '段数
    
    Dim I&, K&, J& '数え上げ用(Integer型)
    
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
        
        If K > 1 Then '下半球の分
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
    火薬座標取得 = KayakuZahyoList
    
End Function

Private Function ExtractRowArray(Array2D, TargetRow&)
'二次元配列の指定行を一次元配列で抽出する
'20210917

'引数
'Array2D  ・・・二次元配列
'TargetRow・・・抽出する対象の行番号


    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数

    If TargetRow < 1 Then
        MsgBox ("抽出する行番号は1以上の値を入れてください")
        Stop
        End
    ElseIf TargetRow > N Then
        MsgBox ("抽出する行番号は元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    End If

    '処理
    Dim Output
    ReDim Output(1 To M)
    
    For I = 1 To M
        Output(I) = Array2D(TargetRow, I)
    Next I
    
    '出力
    ExtractRowArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function ExtractArray(Array2D, StartRow&, StartCol&, EndRow&, EndCol&)
'二次元配列の指定範囲を配列として抽出する
'20210917

'引数
'Array2D ・・・二次元配列
'StartRow・・・抽出範囲の開始行番号
'StartCol・・・抽出範囲の開始列番号
'EndRow  ・・・抽出範囲の終了行番号
'EndCol  ・・・抽出範囲の終了列番号
                                   
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
    
    If StartRow > EndRow Then
        MsgBox ("抽出範囲の開始行「StartRow」は、終了行「EndRow」以下でなければなりません")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("抽出範囲の開始列「StartCol」は、終了列「EndCol」以下でなければなりません")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("抽出範囲の開始行「StartRow」は1以上の値を入れてください")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("抽出範囲の開始列「StartCol」は1以上の値を入れてください")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("抽出範囲の終了行「StartRow」は抽出元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("抽出範囲の終了列「StartCol」は抽出元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '出力
    ExtractArray = Output
    
End Function

Private Function DrawPolyLine(XYList, TargetSheet As Worksheet) As Shape
'XY座標からポリラインを描く
'シェイプをオブジェクト変数として返す
'20210921

'引数
'XYList         ・・・XY座標が入った二次元配列 X方向→右方向 Y方向→下方向
'TargetSheet    ・・・作図対象のシート

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
'ポリラインに点を追加して延長する
'20211008

'引数
'InputShape         ・・・対象のポリライン
'AddX               ・・・追加する点のX座標（右方向）
'AddY               ・・・追加する点のY座標（下方向）
'[DeleteFirstPoint] ・・・対象の曲線の最初の点を削除するかどうか


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
'現在カーソル位置のドキュメント座標取得
'カーソル位置のスクリーン座標を、
'カーソルが乗っているセルの四隅のスクリーン座標の関係性から、
'カーソルが乗っているセルの四隅のドキュメント座標をもとに補間して、
'カーソル位置のドキュメント座標を求める。
'20211005

'引数
'[ImmidiateShow]・・・イミディエイトウィンドウに計算結果などを表示するか(デフォルトはTrue)

'返り値
'Output(1 to 2)・・・1:カーソル位置のドキュメント座標X,2:カーソル位置のドキュメント座標Y(Double型)
'カーソルがシート内にない場合はEmptyを返す。

'参考：https://gist.github.com/furyutei/f0668f33d62ccac95d1643f15f19d99a?s=09#to-footnote-1

    Dim Win As Window
    Set Win = ActiveWindow
    
    'カーソルのスクリーン座標取得
    Dim Cursor As PointAPI, CursorScreenX#, CursorScreenY#
    Call GetCursorPos(Cursor)
    CursorScreenX = Cursor.X
    CursorScreenY = Cursor.Y
    
    'カーソルが乗っているセルを取得
    Dim CursorCell As Range, Dummy
    Set Dummy = Win.RangeFromPoint(CursorScreenX, CursorScreenY)
    If TypeName(Dummy) = "Range" Then
        Set CursorCell = Dummy
    Else
        'カーソルがセルに乗ってないので終了
        Exit Function
    End If
    
    '四隅のスクリーン座標を取得
    Dim X1Screen#, X2Screen#, Y1Screen#, Y2Screen# '四隅のスクリーン座標
    Dummy = GetXYCellScreenUpperLeft(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X1Screen = Dummy(1)
    Y1Screen = Dummy(2)
    
    Dummy = GetXYCellScreenLowerRight(CursorCell)
    If IsEmpty(Dummy) Then Exit Function
    X2Screen = Dummy(1)
    Y2Screen = Dummy(2)
    
    '四隅のドキュメント座標取得
    Dim X1Document#, X2Document#, Y1Document#, Y2Document# '四隅のドキュメント座標
    X1Document = CursorCell.Left
    X2Document = CursorCell.Left + CursorCell.Width
    Y1Document = CursorCell.Top
    Y2Document = CursorCell.Top + CursorCell.Height
    
    'マウスカーソルのドキュメント座標を補間で計算
    Dim CursorDocumentX#, CursorDocumentY#
    CursorDocumentX = X1Document + (X2Document - X1Document) * (CursorScreenX - X1Screen) / (X2Screen - X1Screen)
    CursorDocumentY = Y1Document + (Y2Document - Y1Document) * (CursorScreenY - Y1Screen) / (Y2Screen - Y1Screen)
        
    '出力
    Dim Output#(1 To 2)
    Output(1) = CursorDocumentX
    Output(2) = CursorDocumentY
    
    GetXYDocumentFromCursor = Output
    
    '確認表示
    If ImmidiateShow Then
        Debug.Print "カーソルの乗ったセル", CursorCell.Address(False, False)
        Debug.Print "カーソルスクリーン座標", "CursorScreenX:" & CursorScreenX, "CursorScreenY:" & CursorScreenY
        Debug.Print "カーソルドキュメント座標", "CursorDocumentX:" & WorksheetFunction.Round(CursorDocumentX, 1), "CursorDocumentY:" & WorksheetFunction.Round(CursorDocumentY, 1)
        Debug.Print "セル左上スクリーン座標", "X1Screen:" & X1Screen, , "Y1Screen:" & Y1Screen
        Debug.Print "セル左上ドキュメント座標", "X1Document:" & X1Document, "Y1Document:" & Y1Document
        Debug.Print "セル右下スクリーン座標", "X2Screen:" & X2Screen, , "Y2Screen:" & Y2Screen
        Debug.Print "セル右下ドキュメント座標", "X2Document:" & X2Document, "Y2Document:" & Y2Document
    End If

End Function

Private Function GetXYCellScreenUpperLeft(TargetCell As Range)
'指定セルの左上のスクリーン座標XYを取得する。
'20211005

'引数
'TargetCell・・・対象のセル(Range型)

'返り値
'Output(1 to 2)・・・1:セル左上のスクリーン座標X,2;セル左上のスクリーン座標Y(Double型)

    'セルが表示されているPane(ウィンドウ枠の固定を考慮した表示エリア)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
       
    '【PointsToScreenPixelsの注意事項】
    '【注】対象セルがシート上で表示されていないと取得不可。一部でも表示されていたら取得可能。
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top)
    
    GetXYCellScreenUpperLeft = Output
    
End Function

Private Function GetPaneOfCell(TargetCell As Range) As Pane
'指定セルのPaneを取得する
'ウィンドウ枠固定、ウィンドウ分割の設定でも取得できる。
'参考：http://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_100_120.html
'20211006

'引数
'TargetCell・・・対象のセル/Range型

'返り値
'指定セルが含まれるPane/Pane型
'指定セルが表示範囲外ならNothing
    
    Dim Win As Window
    Set Win = ActiveWindow
    
    Dim Output As Pane
    Dim I& '数え上げ用(Long型)
    
    ' ウィンドウ分割無しか
    If Not Win.FreezePanes And Not Win.Split Then
        'ウィンドウ枠固定でもウィンドウ分割でもない場合
        ' 表示以外にセルがある場合は無視
        If Intersect(Win.VisibleRange, TargetCell) Is Nothing Then Exit Function
        Set Output = Win.Panes(1)
    Else ' 分割あり
        If Win.FreezePanes Then
            ' ウィンドウ枠固定の場合
            ' どのウィンドウに属するか判定
            For I = 1 To Win.Panes.Count
                If Not Intersect(Win.Panes(I).VisibleRange, TargetCell) Is Nothing Then
                    'Paneの表示範囲に含まれる場合はそのPaneを取得
                    Set Output = Win.Panes(I)
                    Exit For
                End If
            Next I
            
            '見つからなかった場合
            If Output Is Nothing Then Exit Function
        Else
            'ウィンドウ分割の場合
            ' ウィンドウ分割はアクティブペインのみ判定
            If Not Intersect(Win.ActivePane.VisibleRange, TargetCell) Is Nothing Then
                Set Output = Win.ActivePane
            Else
                Exit Function
            End If
        End If
    End If
    
    '出力
    Set GetPaneOfCell = Output
    
End Function

Private Function GetXYCellScreenLowerRight(TargetCell As Range)
'指定セルの右下のスクリーン座標XYを取得する。
'20211005

'引数
'TargetCell・・・対象のセル(Range型)

'返り値
'Output(1 to 2)・・・1:セル右下のスクリーン座標X,2;セル右下のスクリーン座標Y(Double型)

    'セルが表示されているPane(ウィンドウ枠の固定を考慮した表示エリア)
    Dim Pane As Pane
    Set Pane = GetPaneOfCell(TargetCell)
    If Pane Is Nothing Then Exit Function
    
    '【PointsToScreenPixelsの注意事項】
    '【注】対象セルがシート上で表示されていないと取得不可。一部でも表示されていたら取得可能。
    Dim Output#(1 To 2)
    Output(1) = Pane.PointsToScreenPixelsX(TargetCell.Left + TargetCell.Width)
    Output(2) = Pane.PointsToScreenPixelsY(TargetCell.Top + TargetCell.Height)
    
    GetXYCellScreenLowerRight = Output
    
End Function


