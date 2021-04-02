Attribute VB_Name = "M_myinterp"
Option Explicit

Rem Ver2 shibuya 2013.1.1



Function myinterp(x As Range, y As Range, xi As Double, Optional Ex As Integer = 0) As Double

Dim tmp As Variant
Dim Lop As Long
Dim xn, yn As Long ' x size
Dim xLoPos As Long ' xLoの位置
Dim xHiPos As Long ' xHiの位置
Dim xLoExist As Boolean
Dim xHiExist As Boolean  'xLo,xHiがある場合True
Dim xLoDelVal As Double
Dim xHiDelVal As Double
Dim xLo2Pos As Long ' xLoの位置
Dim xHi2Pos As Long ' xHiの位置
Dim xLo2Exist As Boolean
Dim xHi2Exist As Boolean  'xLo,xHiがある場合True
Dim xLo2DelVal As Double
Dim xHi2DelVal As Double
Dim xMax As Double
Dim xMin As Double
Dim xMaxPos As Long
Dim xMinPos As Long
Dim xMaxMinExist As Boolean

Dim getValue As Double
Dim Delta As Double

    ' sizeチェック
    xn = myCheckSize(x)
    yn = myCheckSize(y)
    If xn <> yn Then
        myinterp = "#Err Diff Size"
        Exit Function
    End If

    ' 位置探査
    xLoExist = False
    xHiExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' Delta演算前に，セル内に値が有るか同かをチェックする。
            Delta = myGetValue(x, Lop) - xi
            'Lo側探索
            If Delta <= 0 Then
                If xLoExist = False Then
                    'xLoVal は，初回の負はそのまま代入
                    xLoDelVal = Delta
                    xLoPos = Lop
                    xLoExist = True
                ElseIf Delta > xLoDelVal Then
                    '２回目以降で　以前より大きい（0に近い）場合
                    xLoDelVal = Delta
                    xLoPos = Lop
                End If
            End If
            'Hi側探索
            If Delta > 0 Then
                If xHiExist = False Then
                    'xHiVal:初回の正はそのまま代入
                    xHiDelVal = Delta
                    xHiPos = Lop
                    xHiExist = True
                ElseIf Delta < xHiDelVal Then
                    '２回目以降で　以前より小さい（0に近い）場合
                    xHiDelVal = Delta
                    xHiPos = Lop
                End If
            End If
        End If
    Next Lop

    '二次探査
    xLo2Exist = False
    xHi2Exist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' xLoPosとxHiPosは，１次探査で見つかっているため
        ' 二次探査ではチェックしない。
            If xLoExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xLoPos)
                'Lo側探索
                If Delta < 0 Then
                    If xLo2Exist = False Then
                        'xLoVal は，初回の負はそのまま代入
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                        xLo2Exist = True
                    ElseIf Delta > xLo2DelVal Then
                        '２回目以降で　以前より大きい（0に近い）場合
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                    End If
                End If
            End If
            'Hi側探索
            If xHiExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xHiPos)
                If Delta > 0 Then
                    If xHi2Exist = False Then
                        'xHiVal:初回の正はそのまま代入
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                        xHi2Exist = True
                    ElseIf Delta < xHi2DelVal Then
                        '２回目以降で　以前より小さい（0に近い）場合
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop




    '最大値と最小値を取得する
    xMaxMinExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
            tmp = myGetValue(x, Lop)
            If xMaxMinExist = False Then
                xMax = tmp
                xMin = xMax
                xMaxPos = Lop
                xMinPos = Lop
                xMaxMinExist = True
            End If
            If tmp > xMax Then
                xMax = tmp
                xMaxPos = Lop
            End If
            If tmp < xMin Then
                xMin = tmp
                xMinPos = Lop
            End If
        End If
    Next Lop

    '外挿出力
    ' xHiPos,xHi2Pos,xLoPos,xLo2Pos は１次，２次探査で見つかっているポジション
    If Ex = 1 And xLoExist = False And xLo2Exist = False Then
       myinterp = interp_linior( _
            myGetValue(x, xHiPos), myGetValue(x, xHi2Pos), _
            myGetValue(y, xHiPos), myGetValue(y, xHi2Pos), xi)
       Exit Function
    End If
    If Ex = 1 And xHiExist = False And xHi2Exist = False Then
       myinterp = interp_linior( _
            myGetValue(x, xLoPos), myGetValue(x, xLo2Pos), _
            myGetValue(y, xLoPos), myGetValue(y, xLo2Pos), xi)
       Exit Function
    End If

    '内挿出力
    ' xMinPos,xMaxPos は１次，２次探査で見つかっているポジション

    If xLoExist = False Then
        myinterp = myGetValue(y, xMinPos)
    ElseIf xHiExist = False Then
        myinterp = myGetValue(y, xMaxPos)
    Else
        myinterp = interp_linior( _
            myGetValue(x, xLoPos), myGetValue(x, xHiPos), _
            myGetValue(y, xLoPos), myGetValue(y, xHiPos), xi)
    End If
End Function


Function myinterp3d(x As Range, y As Range, z As Range, xi As Double, yi As Range, Optional Ex As Integer = 0) As Double


Dim tmp As Variant
Dim xn As Long  'X軸のサイズ
Dim yn As Long  'Y軸のサイズ
Dim znc As Long
Dim znr As Long

Dim Lop As Long
Dim Delta As Double
Dim yLoPos As Long ' xLoの位置
Dim yHiPos As Long ' xHiの位置
Dim yLoExist As Boolean
Dim yHiExist As Boolean  'xLo,xHiがある場合True
Dim yLoDelVal As Double
Dim yHiDelVal As Double

Dim yLo2Pos As Long ' xLoの位置
Dim yHi2Pos As Long ' xHiの位置
Dim yLo2Exist As Boolean
Dim yHi2Exist As Boolean  'xLo,xHiがある場合True
Dim yLo2DelVal As Double
Dim yHi2DelVal As Double

Dim yMax As Double
Dim yMin As Double
Dim yMaxPos As Long
Dim yMinPos As Long
Dim yMaxMinExist As Boolean

Dim xLoPos As Long ' xLoの位置
Dim xHiPos As Long ' xHiの位置
Dim xLoExist As Boolean
Dim xHiExist As Boolean  'xLo,xHiがある場合True
Dim xLoDelVal As Double
Dim xHiDelVal As Double
Dim xLo2Pos As Long ' xLoの位置
Dim xHi2Pos As Long ' xHiの位置
Dim xLo2Exist As Boolean
Dim xHi2Exist As Boolean  'xLo,xHiがある場合True
Dim xLo2DelVal As Double
Dim xHi2DelVal As Double
Dim xMax As Double
Dim xMin As Double
Dim xMaxPos As Long
Dim xMinPos As Long
Dim xMaxMinExist As Boolean


Dim zDLowLow As Double
Dim zDLowHi As Double
Dim zDHiLow As Double
Dim zDHiHi As Double
Dim xDLowPos As Long
Dim xDHiPos As Long
Dim yDLowPos As Long
Dim yDHiPos As Long

Dim zDYLow As Double
Dim zDYHi As Double


Dim ydata As Range

    ' sizeチェック
    xn = myCheckSize(x)
    yn = myCheckSize(y)
    znc = myCheckSize3dRow(z)
    znr = myCheckSize3dColumn(z)
    If xn <> znr Then
            myinterp3d = "#Err x axis Size"
            Exit Function
    End If
    If yn <> znc Then
            myinterp3d = "#Err y axis Size"
            Exit Function
    End If

    'Y軸 位置探査 (先にY軸の探査をしてテーブル化する。）
    yLoExist = False
    yHiExist = False
    For Lop = 1 To yn
        If myCheckCellIsOk(y, Lop) = True Then
        ' Delta演算前に，セル内に値が有るか同かをチェックする。
            Delta = myGetValue(y, Lop) - yi
            'Lo側探索
            If Delta <= 0 Then
                If yLoExist = False Then
                    'yLoVal は，初回の負はそのまま代入
                    yLoDelVal = Delta
                    yLoPos = Lop
                    yLoExist = True
                ElseIf Delta > yLoDelVal Then
                    '２回目以降で　以前より大きい（0に近い）場合
                    yLoDelVal = Delta
                    yLoPos = Lop
                End If
            End If
            'Hi側探索
            If Delta > 0 Then
                If yHiExist = False Then
                    'yHiVal:初回の正はそのまま代入
                    yHiDelVal = Delta
                    yHiPos = Lop
                    yHiExist = True
                ElseIf Delta < yHiDelVal Then
                    '２回目以降で　以前より小さい（0に近い）場合
                    yHiDelVal = Delta
                    yHiPos = Lop
                End If
            End If
        End If
    Next Lop

    'Y軸 位置探査 2次探査(先にY軸の探査をしてテーブル化する。）

    '二次探査
    yLo2Exist = False
    yHi2Exist = False
    For Lop = 1 To yn
        If myCheckCellIsOk(y, Lop) = True Then
        ' yLoPosとyHiPosは，１次探査で見つかっているため
        ' 二次探査ではチェックしない。
            If yLoExist = True Then
                Delta = myGetValue(y, Lop) - myGetValue(y, yLoPos)
                'Lo側探索
                If Delta < 0 Then
                    If yLo2Exist = False Then
                        'yLoVal は，初回の負はそのまま代入
                        yLo2DelVal = Delta
                        yLo2Pos = Lop
                        yLo2Exist = True
                    ElseIf Delta > yLo2DelVal Then
                        '２回目以降で　以前より大きい（0に近い）場合
                        yLo2DelVal = Delta
                        yLo2Pos = Lop
                    End If
                End If
            End If
            'Hi側探索
            If yHiExist = True Then
                Delta = myGetValue(y, Lop) - myGetValue(y, yHiPos)
                If Delta > 0 Then
                    If yHi2Exist = False Then
                        'yHiVal:初回の正はそのまま代入
                        yHi2DelVal = Delta
                        yHi2Pos = Lop
                        yHi2Exist = True
                    ElseIf Delta < yHi2DelVal Then
                        '２回目以降で　以前より小さい（0に近い）場合
                        yHi2DelVal = Delta
                        yHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop


    '最大値と最小値を取得する
    yMaxMinExist = False
    For Lop = 1 To yn
        If myCheckCellIsOk(y, Lop) = True Then
            tmp = myGetValue(y, Lop)
            If yMaxMinExist = False Then
                yMax = tmp
                yMin = yMax
                yMaxPos = Lop
                yMinPos = Lop
                yMaxMinExist = True
            End If
            If tmp > yMax Then
                yMax = tmp
                yMaxPos = Lop
            End If
            If tmp < yMin Then
                yMin = tmp
                yMinPos = Lop
            End If
        End If
    Next Lop


    ' 位置探査
    xLoExist = False
    xHiExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' Delta演算前に，セル内に値が有るか同かをチェックする。
            Delta = myGetValue(x, Lop) - xi
            'Lo側探索
            If Delta <= 0 Then
                If xLoExist = False Then
                    'xLoVal は，初回の負はそのまま代入
                    xLoDelVal = Delta
                    xLoPos = Lop
                    xLoExist = True
                ElseIf Delta > xLoDelVal Then
                    '２回目以降で　以前より大きい（0に近い）場合
                    xLoDelVal = Delta
                    xLoPos = Lop
                End If
            End If
            'Hi側探索
            If Delta > 0 Then
                If xHiExist = False Then
                    'xHiVal:初回の正はそのまま代入
                    xHiDelVal = Delta
                    xHiPos = Lop
                    xHiExist = True
                ElseIf Delta < xHiDelVal Then
                    '２回目以降で　以前より小さい（0に近い）場合
                    xHiDelVal = Delta
                    xHiPos = Lop
                End If
            End If
        End If
    Next Lop

    '二次探査
    xLo2Exist = False
    xHi2Exist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' xLoPosとxHiPosは，１次探査で見つかっているため
        ' 二次探査ではチェックしない。
            If xLoExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xLoPos)
                'Lo側探索
                If Delta < 0 Then
                    If xLo2Exist = False Then
                        'xLoVal は，初回の負はそのまま代入
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                        xLo2Exist = True
                    ElseIf Delta > xLo2DelVal Then
                        '２回目以降で　以前より大きい（0に近い）場合
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                    End If
                End If
            End If
            'Hi側探索
            If xHiExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xHiPos)
                If Delta > 0 Then
                    If xHi2Exist = False Then
                        'xHiVal:初回の正はそのまま代入
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                        xHi2Exist = True
                    ElseIf Delta < xHi2DelVal Then
                        '２回目以降で　以前より小さい（0に近い）場合
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop




    '最大値と最小値を取得する
    xMaxMinExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
            tmp = myGetValue(x, Lop)
            If xMaxMinExist = False Then
                xMax = tmp
                xMin = xMax
                xMaxPos = Lop
                xMinPos = Lop
                xMaxMinExist = True
            End If
            If tmp > xMax Then
                xMax = tmp
                xMaxPos = Lop
            End If
            If tmp < xMin Then
                xMin = tmp
                xMinPos = Lop
            End If
        End If
    Next Lop




    '固定Y軸で，テーブルを作成する。
'    ReDim ydata(1 To xn) As Range


    '外挿出力
    ' xHiPos,xHi2Pos,xLoPos,xLo2Pos は１次，２次探査で見つかっているポジション
    If Ex = 1 Then
       myinterp3d = "#Err mijissou"
       Exit Function
    End If

    '内挿出力
    If Ex = 0 Then
        ' Y軸テーブルの作成

        If xLoExist = False Then
            xDLowPos = xMinPos
            xDHiPos = xMinPos
        ElseIf xHiExist = False Then
            xDLowPos = xMaxPos
            xDHiPos = xMaxPos
        Else
            xDLowPos = xLoPos
            xDHiPos = xHiPos
        End If
        If yLoExist = False Then
            yDLowPos = yMinPos
            yDHiPos = yMinPos
        ElseIf yHiExist = False Then
            yDLowPos = yMaxPos
            yDHiPos = yMaxPos
        Else
            yDLowPos = yLoPos
            yDHiPos = yHiPos
        End If

        zDLowLow = myGetValue3d(z, xDLowPos, yDLowPos)
        zDLowHi = myGetValue3d(z, xDLowPos, yDHiPos)
        zDYLow = interp_linior3d(myGetValue(y, yDLowPos), myGetValue(y, yDHiPos), _
            zDLowLow, zDLowHi, yi)

        zDHiLow = myGetValue3d(z, xDHiPos, yDLowPos)
        zDHiHi = myGetValue3d(z, xDHiPos, yDHiPos)
        zDYHi = interp_linior3d(myGetValue(y, yDLowPos), myGetValue(y, yDHiPos), _
            zDHiLow, zDHiHi, yi)

        tmp = interp_linior3d(myGetValue(x, xDLowPos), myGetValue(x, xDHiPos), _
            zDYLow, zDYHi, xi)

        myinterp3d = tmp
        Exit Function

'        Else
'            tmp = interp_linior( _
'                myGetValue(x, xLoPos), myGetValue(x, xHiPos), _
'                myGetValue(y, xLoPos), myGetValue(y, xHiPos), xi)
'        End If
'
'        tmp = interp_linior( _
'            myGetValue(x, xLoPos), myGetValue(x, xHiPos), _
'            myGetValue(y, xLoPos), myGetValue(y, xHiPos), xi)

    End If

End Function





Private Function myCheckSize(x As Range) As Long
'選択範囲（レンジ）の範囲を調べる
    If x.Rows.Count >= x.Columns.Count Then
        myCheckSize = x.Rows.Count
    Else
        myCheckSize = x.Columns.Count
    End If
End Function

Private Function myCheckSize3dRow(x As Range) As Long
'選択範囲（レンジ）の範囲を調べる

        myCheckSize3dRow = x.Rows.Count

End Function

Private Function myCheckSize3dColumn(x As Range) As Long
'選択範囲（レンジ）の範囲を調べる

        myCheckSize3dColumn = x.Columns.Count

End Function


Private Function myGetValue(x As Range, Pos As Long) As Double
'選択範囲（レンジ）の範囲を調べる
    If x.Rows.Count >= x.Columns.Count Then
        myGetValue = x.Cells(Pos, 1).Value
    Else
        myGetValue = x.Cells(1, Pos).Value
    End If
End Function

Private Function myGetValue3d(x As Range, xPos As Long, yPos As Long) As Double
'選択範囲（レンジ）の範囲を調べる

        myGetValue3d = x.Cells(yPos, xPos).Value

End Function

Function myCheckCellIsOk(x As Range, Pos As Long) As Boolean
'選択範囲（レンジ）の範囲を調べる
Dim LcIsEmpty As Boolean
Dim LcIsNumeric As Boolean
    If x.Rows.Count >= x.Columns.Count Then
        LcIsEmpty = IsEmpty(x.Cells(Pos, 1))
        LcIsNumeric = IsNumeric(x.Cells(Pos, 1))

    Else
        LcIsEmpty = IsEmpty(x.Cells(1, Pos))
        LcIsNumeric = IsNumeric(x.Cells(1, Pos))
    End If

    If LcIsEmpty = False And LcIsNumeric = True Then
        myCheckCellIsOk = True
        'numeric の場合，空セルの場合にもTrueになるため
        '空セルでない(isEmpty = false の条件とANDする
    Else
        myCheckCellIsOk = False
    End If
End Function

Function interp_linior(xLo, xHi, yLo, yHi, xi) As Double
' 三角補間
Dim ep As Double
' (xi - xLo)/(xHi - xLo) = (yi - yLo)/(yHi - yLo)

    If (xLo - xHi) = 0 Then
        interp_linior = yLo
        'interp_linior = ("#ERR 0DIV")
        ' Exit Function
    Else
        ep = (xi - xLo) / (xHi - xLo)
        ' ep が 1以上のときは，外挿の意味
        interp_linior = yLo + (yHi - yLo) * ep
    End If
End Function


Function interp_linior3d(xLo, xHi, yLo, yHi, xi) As Double
' 三角補間
Dim ep As Double
' (xi - xLo)/(xHi - xLo) = (yi - yLo)/(yHi - yLo)

    If (xLo - xHi) = 0 Then
        interp_linior3d = yLo  ' マップ補間様の特別処置
        ' Exit Function
    Else
        ep = (xi - xLo) / (xHi - xLo)
        ' ep が 1以上のときは，外挿の意味
        interp_linior3d = yLo + (yHi - yLo) * ep
    End If
End Function



