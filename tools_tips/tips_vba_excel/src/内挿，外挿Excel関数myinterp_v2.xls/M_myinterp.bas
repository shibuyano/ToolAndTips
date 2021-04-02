Attribute VB_Name = "M_myinterp"
Option Explicit

Rem Ver2 shibuya 2013.1.1



Function myinterp(x As Range, y As Range, xi As Double, Optional Ex As Integer = 0) As Double

Dim tmp As Variant
Dim Lop As Long
Dim xn, yn As Long ' x size
Dim xLoPos As Long ' xLo�̈ʒu
Dim xHiPos As Long ' xHi�̈ʒu
Dim xLoExist As Boolean
Dim xHiExist As Boolean  'xLo,xHi������ꍇTrue
Dim xLoDelVal As Double
Dim xHiDelVal As Double
Dim xLo2Pos As Long ' xLo�̈ʒu
Dim xHi2Pos As Long ' xHi�̈ʒu
Dim xLo2Exist As Boolean
Dim xHi2Exist As Boolean  'xLo,xHi������ꍇTrue
Dim xLo2DelVal As Double
Dim xHi2DelVal As Double
Dim xMax As Double
Dim xMin As Double
Dim xMaxPos As Long
Dim xMinPos As Long
Dim xMaxMinExist As Boolean

Dim getValue As Double
Dim Delta As Double

    ' size�`�F�b�N
    xn = myCheckSize(x)
    yn = myCheckSize(y)
    If xn <> yn Then
        myinterp = "#Err Diff Size"
        Exit Function
    End If

    ' �ʒu�T��
    xLoExist = False
    xHiExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' Delta���Z�O�ɁC�Z�����ɒl���L�邩�������`�F�b�N����B
            Delta = myGetValue(x, Lop) - xi
            'Lo���T��
            If Delta <= 0 Then
                If xLoExist = False Then
                    'xLoVal �́C����̕��͂��̂܂ܑ��
                    xLoDelVal = Delta
                    xLoPos = Lop
                    xLoExist = True
                ElseIf Delta > xLoDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                    xLoDelVal = Delta
                    xLoPos = Lop
                End If
            End If
            'Hi���T��
            If Delta > 0 Then
                If xHiExist = False Then
                    'xHiVal:����̐��͂��̂܂ܑ��
                    xHiDelVal = Delta
                    xHiPos = Lop
                    xHiExist = True
                ElseIf Delta < xHiDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                    xHiDelVal = Delta
                    xHiPos = Lop
                End If
            End If
        End If
    Next Lop

    '�񎟒T��
    xLo2Exist = False
    xHi2Exist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' xLoPos��xHiPos�́C�P���T���Ō������Ă��邽��
        ' �񎟒T���ł̓`�F�b�N���Ȃ��B
            If xLoExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xLoPos)
                'Lo���T��
                If Delta < 0 Then
                    If xLo2Exist = False Then
                        'xLoVal �́C����̕��͂��̂܂ܑ��
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                        xLo2Exist = True
                    ElseIf Delta > xLo2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                    End If
                End If
            End If
            'Hi���T��
            If xHiExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xHiPos)
                If Delta > 0 Then
                    If xHi2Exist = False Then
                        'xHiVal:����̐��͂��̂܂ܑ��
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                        xHi2Exist = True
                    ElseIf Delta < xHi2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop




    '�ő�l�ƍŏ��l���擾����
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

    '�O�}�o��
    ' xHiPos,xHi2Pos,xLoPos,xLo2Pos �͂P���C�Q���T���Ō������Ă���|�W�V����
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

    '���}�o��
    ' xMinPos,xMaxPos �͂P���C�Q���T���Ō������Ă���|�W�V����

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
Dim xn As Long  'X���̃T�C�Y
Dim yn As Long  'Y���̃T�C�Y
Dim znc As Long
Dim znr As Long

Dim Lop As Long
Dim Delta As Double
Dim yLoPos As Long ' xLo�̈ʒu
Dim yHiPos As Long ' xHi�̈ʒu
Dim yLoExist As Boolean
Dim yHiExist As Boolean  'xLo,xHi������ꍇTrue
Dim yLoDelVal As Double
Dim yHiDelVal As Double

Dim yLo2Pos As Long ' xLo�̈ʒu
Dim yHi2Pos As Long ' xHi�̈ʒu
Dim yLo2Exist As Boolean
Dim yHi2Exist As Boolean  'xLo,xHi������ꍇTrue
Dim yLo2DelVal As Double
Dim yHi2DelVal As Double

Dim yMax As Double
Dim yMin As Double
Dim yMaxPos As Long
Dim yMinPos As Long
Dim yMaxMinExist As Boolean

Dim xLoPos As Long ' xLo�̈ʒu
Dim xHiPos As Long ' xHi�̈ʒu
Dim xLoExist As Boolean
Dim xHiExist As Boolean  'xLo,xHi������ꍇTrue
Dim xLoDelVal As Double
Dim xHiDelVal As Double
Dim xLo2Pos As Long ' xLo�̈ʒu
Dim xHi2Pos As Long ' xHi�̈ʒu
Dim xLo2Exist As Boolean
Dim xHi2Exist As Boolean  'xLo,xHi������ꍇTrue
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

    ' size�`�F�b�N
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

    'Y�� �ʒu�T�� (���Y���̒T�������ăe�[�u��������B�j
    yLoExist = False
    yHiExist = False
    For Lop = 1 To yn
        If myCheckCellIsOk(y, Lop) = True Then
        ' Delta���Z�O�ɁC�Z�����ɒl���L�邩�������`�F�b�N����B
            Delta = myGetValue(y, Lop) - yi
            'Lo���T��
            If Delta <= 0 Then
                If yLoExist = False Then
                    'yLoVal �́C����̕��͂��̂܂ܑ��
                    yLoDelVal = Delta
                    yLoPos = Lop
                    yLoExist = True
                ElseIf Delta > yLoDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                    yLoDelVal = Delta
                    yLoPos = Lop
                End If
            End If
            'Hi���T��
            If Delta > 0 Then
                If yHiExist = False Then
                    'yHiVal:����̐��͂��̂܂ܑ��
                    yHiDelVal = Delta
                    yHiPos = Lop
                    yHiExist = True
                ElseIf Delta < yHiDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                    yHiDelVal = Delta
                    yHiPos = Lop
                End If
            End If
        End If
    Next Lop

    'Y�� �ʒu�T�� 2���T��(���Y���̒T�������ăe�[�u��������B�j

    '�񎟒T��
    yLo2Exist = False
    yHi2Exist = False
    For Lop = 1 To yn
        If myCheckCellIsOk(y, Lop) = True Then
        ' yLoPos��yHiPos�́C�P���T���Ō������Ă��邽��
        ' �񎟒T���ł̓`�F�b�N���Ȃ��B
            If yLoExist = True Then
                Delta = myGetValue(y, Lop) - myGetValue(y, yLoPos)
                'Lo���T��
                If Delta < 0 Then
                    If yLo2Exist = False Then
                        'yLoVal �́C����̕��͂��̂܂ܑ��
                        yLo2DelVal = Delta
                        yLo2Pos = Lop
                        yLo2Exist = True
                    ElseIf Delta > yLo2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                        yLo2DelVal = Delta
                        yLo2Pos = Lop
                    End If
                End If
            End If
            'Hi���T��
            If yHiExist = True Then
                Delta = myGetValue(y, Lop) - myGetValue(y, yHiPos)
                If Delta > 0 Then
                    If yHi2Exist = False Then
                        'yHiVal:����̐��͂��̂܂ܑ��
                        yHi2DelVal = Delta
                        yHi2Pos = Lop
                        yHi2Exist = True
                    ElseIf Delta < yHi2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                        yHi2DelVal = Delta
                        yHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop


    '�ő�l�ƍŏ��l���擾����
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


    ' �ʒu�T��
    xLoExist = False
    xHiExist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' Delta���Z�O�ɁC�Z�����ɒl���L�邩�������`�F�b�N����B
            Delta = myGetValue(x, Lop) - xi
            'Lo���T��
            If Delta <= 0 Then
                If xLoExist = False Then
                    'xLoVal �́C����̕��͂��̂܂ܑ��
                    xLoDelVal = Delta
                    xLoPos = Lop
                    xLoExist = True
                ElseIf Delta > xLoDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                    xLoDelVal = Delta
                    xLoPos = Lop
                End If
            End If
            'Hi���T��
            If Delta > 0 Then
                If xHiExist = False Then
                    'xHiVal:����̐��͂��̂܂ܑ��
                    xHiDelVal = Delta
                    xHiPos = Lop
                    xHiExist = True
                ElseIf Delta < xHiDelVal Then
                    '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                    xHiDelVal = Delta
                    xHiPos = Lop
                End If
            End If
        End If
    Next Lop

    '�񎟒T��
    xLo2Exist = False
    xHi2Exist = False
    For Lop = 1 To xn
        If myCheckCellIsOk(x, Lop) = True Then
        ' xLoPos��xHiPos�́C�P���T���Ō������Ă��邽��
        ' �񎟒T���ł̓`�F�b�N���Ȃ��B
            If xLoExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xLoPos)
                'Lo���T��
                If Delta < 0 Then
                    If xLo2Exist = False Then
                        'xLoVal �́C����̕��͂��̂܂ܑ��
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                        xLo2Exist = True
                    ElseIf Delta > xLo2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO���傫���i0�ɋ߂��j�ꍇ
                        xLo2DelVal = Delta
                        xLo2Pos = Lop
                    End If
                End If
            End If
            'Hi���T��
            If xHiExist = True Then
                Delta = myGetValue(x, Lop) - myGetValue(x, xHiPos)
                If Delta > 0 Then
                    If xHi2Exist = False Then
                        'xHiVal:����̐��͂��̂܂ܑ��
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                        xHi2Exist = True
                    ElseIf Delta < xHi2DelVal Then
                        '�Q��ڈȍ~�Ł@�ȑO��菬�����i0�ɋ߂��j�ꍇ
                        xHi2DelVal = Delta
                        xHi2Pos = Lop
                    End If
                End If
            End If
        End If
    Next Lop




    '�ő�l�ƍŏ��l���擾����
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




    '�Œ�Y���ŁC�e�[�u�����쐬����B
'    ReDim ydata(1 To xn) As Range


    '�O�}�o��
    ' xHiPos,xHi2Pos,xLoPos,xLo2Pos �͂P���C�Q���T���Ō������Ă���|�W�V����
    If Ex = 1 Then
       myinterp3d = "#Err mijissou"
       Exit Function
    End If

    '���}�o��
    If Ex = 0 Then
        ' Y���e�[�u���̍쐬

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
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�
    If x.Rows.Count >= x.Columns.Count Then
        myCheckSize = x.Rows.Count
    Else
        myCheckSize = x.Columns.Count
    End If
End Function

Private Function myCheckSize3dRow(x As Range) As Long
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�

        myCheckSize3dRow = x.Rows.Count

End Function

Private Function myCheckSize3dColumn(x As Range) As Long
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�

        myCheckSize3dColumn = x.Columns.Count

End Function


Private Function myGetValue(x As Range, Pos As Long) As Double
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�
    If x.Rows.Count >= x.Columns.Count Then
        myGetValue = x.Cells(Pos, 1).Value
    Else
        myGetValue = x.Cells(1, Pos).Value
    End If
End Function

Private Function myGetValue3d(x As Range, xPos As Long, yPos As Long) As Double
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�

        myGetValue3d = x.Cells(yPos, xPos).Value

End Function

Function myCheckCellIsOk(x As Range, Pos As Long) As Boolean
'�I��͈́i�����W�j�͈̔͂𒲂ׂ�
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
        'numeric �̏ꍇ�C��Z���̏ꍇ�ɂ�True�ɂȂ邽��
        '��Z���łȂ�(isEmpty = false �̏�����AND����
    Else
        myCheckCellIsOk = False
    End If
End Function

Function interp_linior(xLo, xHi, yLo, yHi, xi) As Double
' �O�p���
Dim ep As Double
' (xi - xLo)/(xHi - xLo) = (yi - yLo)/(yHi - yLo)

    If (xLo - xHi) = 0 Then
        interp_linior = yLo
        'interp_linior = ("#ERR 0DIV")
        ' Exit Function
    Else
        ep = (xi - xLo) / (xHi - xLo)
        ' ep �� 1�ȏ�̂Ƃ��́C�O�}�̈Ӗ�
        interp_linior = yLo + (yHi - yLo) * ep
    End If
End Function


Function interp_linior3d(xLo, xHi, yLo, yHi, xi) As Double
' �O�p���
Dim ep As Double
' (xi - xLo)/(xHi - xLo) = (yi - yLo)/(yHi - yLo)

    If (xLo - xHi) = 0 Then
        interp_linior3d = yLo  ' �}�b�v��ԗl�̓��ʏ��u
        ' Exit Function
    Else
        ep = (xi - xLo) / (xHi - xLo)
        ' ep �� 1�ȏ�̂Ƃ��́C�O�}�̈Ӗ�
        interp_linior3d = yLo + (yHi - yLo) * ep
    End If
End Function



