Attribute VB_Name = "Module1"

Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Declare Sub StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long)
Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Integer, ByVal hWndinsertafter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal SoundName As String, ByVal Flags As Long) As Long
Declare Function sndPlaySoundL Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal NUL As Long, ByVal uFlags As Long) As Long

Global Const HTCAPTION = 2
Global Const MF_STRING = &H0&
Global Const MF_BYCOMMAND = &H0&
Global Const SC_CLOSE = &HF060
Global Const srccopy = &HCC0020
Global Const srcand = &H8800C6
Global Const srcor = &HEE0086

Private hMenu As Long
Private CloseStr As String '記錄Close MenuItem的字串

Public cards(1 To 52)
Public users(1 To 4, 1 To 13)
Public lost_users(1 To 4)
Public comx(2 To 4, 1 To 3)
Public CardsPlace(1 To 13)     '丟牌之後被蓋住的牌能增加多少空格能按
'Public AICards(2 To 4)     '電腦要丟的牌
Public AICallKing(2 To 4, 1 To 4)   '記錄每個花色有幾張
Public BigKing(2 To 4, 1 To 3)     '1紀錄最多的花色   2最多花色有幾張   3能喊到多少
Public ComKing(1 To 4)      '電腦叫的王
'Public AICards(2 To 4, 1 To 4, 1 To 13)   '電腦各花色有的牌
Public AIPutCard(1 To 4)     '要丟的牌
Public WhoWin     '誰的牌大
Public FirstCard     '美1次第1張牌
Public needwin(1 To 2)     '2方各要贏多少
Public PlayerPoint(0 To 3)     '各有多少點數
Public cards_num
Public cardsbox_num
Public comenable
Public Speed
Public change
Public userpic
Public UserHaveCard     '玩家有沒有跟電腦一樣花色的牌
Public AWin     '我方營多少
Public BWin     '敵方贏多少
Public CheckComplay     '確認每個電腦是不是都丟過排
Public ComPlay     '電腦第幾次丟牌
Public GameKing     '整場比賽的王
Public pass     '有幾個人PASS
Public ReturnPlayer     '檢查玩家能不能出牌
Public FirstPlayer     '第一個出牌的人
Public GameSpeed     '遊戲速度

Sub RndCards()   '洗牌
Randomize Timer
For i = 1 To 52
    cards(i) = i
Next i
For i = 1 To 5000
    a = Int(Rnd * 52) + 1
    b = Int(Rnd * 52) + 1
    Swap cards(a), cards(b)
Next i
cards_num = 1
cardsbox_num = 1
End Sub

Sub UserCards()   '發牌
For i = 1 To 4
  For j = 1 To 13
      users(i, j) = cards(cards_num)
      cards_num = cards_num + 1
  Next j
Next i
End Sub

Sub CardMode()   '排牌
For a = 1 To 4
    For i = 1 To 13
        For j = i + 1 To 13
            If users(a, i) > users(a, j) Then
            change = users(a, i)
            users(a, i) = users(a, j)
            users(a, j) = change
            End If
        Next j
    Next i
Next a
End Sub

Sub Point()   '算各家有多少點數

For i = 0 To 3
    PlayerPoint(i) = 0
Next i

For i = 1 To 4
    For j = 1 To 13
        a = users(i, j)
'        If a = 0 Then a = 13
        If a = 10 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 1
        If a = 23 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 1
        If a = 36 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 1
        If a = 49 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 1
        If a = 11 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 2
        If a = 24 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 2
        If a = 37 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 2
        If a = 50 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 2
        If a = 12 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 3
        If a = 25 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 3
        If a = 38 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 3
        If a = 51 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 3
        If a = 13 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 4
        If a = 26 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 4
        If a = 39 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 4
        If a = 52 Then PlayerPoint(i - 1) = PlayerPoint(i - 1) + 4
    Next j
Next i

For i = 0 To 3
    Form8.Label1(i).Caption = PlayerPoint(i)
Next i

End Sub

Sub NewGamePut()   '開新牌局
RndCards
UserCards
CardMode
Point
CheckComplay = 0
ComPlay = 1
ReturnPlayer = 0
CheckComplay = 0
pass = 0
Unload Form4
Form1.Picture1.Line (110, 100)-(440, 293), &H8000&, BF
Form1.Picture1.Refresh
AWin = 0
BWin = 0
For i = 1 To 13
  DrawCards 120 + (i - 1) * 20, 295, users(1, i)
Next i
For i = 13 To 1 Step -1
  DrawBack 20, 100 + (i - 1) * 8
  DrawBack 190 + (i - 1) * 8, 5
  DrawBack 440, 100 + (i - 1) * 8
Next i

For i = 0 To 35
    Form8.Command1(i).Enabled = True
    Next i

For i = 1 To 4
    ComKing(i) = 0
Next i

For i = 2 To 4
    For j = 1 To 4
        AICallKing(i, j) = 0
    Next j
Next i

For i = 2 To 4
    For j = 1 To 3
        BigKing(i, j) = 0
    Next j
Next i

Form1.Command3(0).Enabled = True

BitBlt Form1.Picture5.hdc, 0, 0, 300, 300, Form1.Picture1.hdc, 0, 0, srccopy

BitBlt Form1.Picture6.hdc, 0, 0, 41, 27, Form1.Picture10.hdc, 0, 0, srccopy
BitBlt Form1.Picture6.hdc, 0, 58, 41, 25, Form1.Picture10.hdc, 0, 0, srccopy
BitBlt Form1.Picture6.hdc, 0, 114, 41, 23, Form1.Picture10.hdc, 0, 0, srccopy
BitBlt Form1.Picture7.hdc, 0, 0, 41, 27, Form1.Picture10.hdc, 0, 0, srccopy
BitBlt Form1.Picture7.hdc, 0, 58, 41, 25, Form1.Picture10.hdc, 0, 0, srccopy
BitBlt Form1.Picture7.hdc, 0, 114, 41, 23, Form1.Picture10.hdc, 0, 0, srccopy

Form1.Picture5.Refresh
Form1.Picture6.Refresh
Form1.Picture7.Refresh

For i = 2 To 4     '數花色數量
    For j = 1 To 13
        If users(i, j) < 53 Then b = 4
        If users(i, j) < 40 Then b = 3
        If users(i, j) < 27 Then b = 2
        If users(i, j) < 14 Then b = 1

        If b = 4 Then AICallKing(i, 4) = AICallKing(i, 4) + 1
        If b = 3 Then AICallKing(i, 3) = AICallKing(i, 3) + 1
        If b = 2 Then AICallKing(i, 2) = AICallKing(i, 2) + 1
        If b = 1 Then AICallKing(i, 1) = AICallKing(i, 1) + 1
    Next j
Next i

For a = 2 To 4     '決定最多的花色
    BigKing(a, 1) = 1
    For i = 1 To 3
        For j = i + 1 To 4
            If AICallKing(a, i) <= AICallKing(a, j) Then
                BigKing(a, 1) = j
            End If
        Next j
    Next i
Next a

For a = 2 To 4     '最多花色有幾張
    BigKing(a, 2) = AICallKing(a, BigKing(a, 1))
Next a

For a = 2 To 4     '看花色張數決定喊到多少
    If BigKing(a, 2) = 4 Then
        BigKing(a, 3) = 2
    End If

    If BigKing(a, 2) = 5 Then
        BigKing(a, 3) = 3
    End If

    If BigKing(a, 2) = 6 Then
        BigKing(a, 3) = 3
    End If

    If BigKing(a, 2) = 7 Then
        BigKing(a, 3) = 4
    End If

    If BigKing(a, 2) = 8 Then
        BigKing(a, 3) = 4
    End If

    If BigKing(a, 2) = 9 Then
        BigKing(a, 3) = 4
    End If

    If BigKing(a, 2) = 10 Then
        BigKing(a, 3) = 5
    End If

    If BigKing(a, 2) = 11 Then
        BigKing(a, 3) = 5
    End If

    If BigKing(a, 2) = 12 Then
        BigKing(a, 3) = 6
    End If

    If BigKing(a, 2) = 13 Then
        BigKing(a, 3) = 7
    End If
Next a

Form1.Picture1.Refresh

End Sub

Sub CheckCards()

For i = 1 To 13
CardsPlace(i) = 0
Next i

For i = 1 To 13
k = 0
If users(1, i) <> 0 Then
    For j = i + 1 To 13
    If users(1, j) = 0 Then k = k + 20
    If users(1, j) <> 0 Then GoTo out:
    Next j
    'k = k * 20
out:
    If k > 51 Then
        k = 50
    End If
    CardsPlace(i) = k
End If
Next i
    
End Sub

Sub UserShowCards(num) '玩家出牌

UserHaveCard = 0

If users(1, num) = 0 Then GoTo out

If FirstPlayer = 2 Then
    If AIPutCard(2) < 53 Then b = 4
    If AIPutCard(2) < 40 Then b = 3
    If AIPutCard(2) < 27 Then b = 2
    If AIPutCard(2) < 14 Then b = 1
    For i = 1 To 13
    If users(1, i) <> 0 Then
        If users(1, i) < 53 Then x = 4
        If users(1, i) < 40 Then x = 3
        If users(1, i) < 27 Then x = 2
        If users(1, i) < 14 Then x = 1
        If x = b Then UserHaveCard = UserHaveCard + 1
    End If
    Next i
End If

If FirstPlayer = 3 Then
    If AIPutCard(3) < 53 Then b = 4
    If AIPutCard(3) < 40 Then b = 3
    If AIPutCard(3) < 27 Then b = 2
    If AIPutCard(3) < 14 Then b = 1
    For i = 1 To 13
    If users(1, i) <> 0 Then
        If users(1, i) < 53 Then x = 4
        If users(1, i) < 40 Then x = 3
        If users(1, i) < 27 Then x = 2
        If users(1, i) < 14 Then x = 1
        If x = b Then UserHaveCard = UserHaveCard + 1
    End If
    Next i
End If

If FirstPlayer = 4 Then
    If AIPutCard(4) < 53 Then b = 4
    If AIPutCard(4) < 40 Then b = 3
    If AIPutCard(4) < 27 Then b = 2
    If AIPutCard(4) < 14 Then b = 1
    For i = 1 To 13
    If users(1, i) <> 0 Then
        If users(1, i) < 53 Then x = 4
        If users(1, i) < 40 Then x = 3
        If users(1, i) < 27 Then x = 2
        If users(1, i) < 14 Then x = 1
        If x = b Then UserHaveCard = UserHaveCard + 1
    End If
    Next i
End If

If UserHaveCard > 0 Then
    If users(1, num) < 53 Then y = 4
    If users(1, num) < 40 Then y = 3
    If users(1, num) < 27 Then y = 2
    If users(1, num) < 14 Then y = 1
    If y <> b Then GoTo out
End If
    
DrawCards 230, 198, users(1, num)
Draw 120 + (num - 1) * 20, 295, num

AIPutCard(1) = users(1, num)


users(1, num) = 0

For i = 1 To 13
If users(1, i) <> 0 Then DrawCards 120 + (i - 1) * 20, 295, users(1, i)
Next i


CheckCards

ReturnPlayer = 0

GameRule

out:
UserHaveCard = 0
End Sub

Sub GameRule()     '系統判定輸贏

If FirstPlayer = 1 Then     '玩家先出牌

    WhoWin = 1
    ComAI (AIPutCard(1))
    If AIPutCard(1) < 53 Then FirstCard = 4
    If AIPutCard(1) < 40 Then FirstCard = 3
    If AIPutCard(1) < 27 Then FirstCard = 2
    If AIPutCard(1) < 14 Then FirstCard = 1
    
    If GameKing = 5 Then GoTo noking
    
    If FirstCard = 4 Then     '玩家出黑桃
        For i = 1 To 4
        If i <> 1 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 4 Then
                
                If e = 4 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 4 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 4 Then
                    If f = 4 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 1 Then WhoWin = i
                    If WhoWin <> 1 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 3 Then     '玩家出紅心
        For i = 1 To 4
        If i <> 1 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 3 Then
                
                If e = 3 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 3 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 3 Then
                    If f = 3 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 1 Then WhoWin = i
                    If WhoWin <> 1 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 2 Then     '玩家出方塊
        For i = 1 To 4
        If i <> 1 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 2 Then
                
                If e = 2 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 2 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 2 Then
                    If f = 2 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 1 Then WhoWin = i
                    If WhoWin <> 1 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 1 Then     '玩家出梅花
        For i = 1 To 4
        If i <> 1 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 1 Then
                
                If e = 1 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 1 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 1 Then
                    If f = 1 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 1 Then WhoWin = i
                    If WhoWin <> 1 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
End If

If FirstPlayer = 2 Then     '電腦1先出牌

    WhoWin = 2
    If AIPutCard(2) < 53 Then FirstCard = 4
    If AIPutCard(2) < 40 Then FirstCard = 3
    If AIPutCard(2) < 27 Then FirstCard = 2
    If AIPutCard(2) < 14 Then FirstCard = 1
    
    If GameKing = 5 Then GoTo noking
    
    If FirstCard = 4 Then     '電腦1出黑桃
        For i = 1 To 4
        If i <> 2 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 4 Then
                
                If e = 4 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 4 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 4 Then
                    If f = 4 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 2 Then WhoWin = i
                    If WhoWin <> 2 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 3 Then     '電腦1出紅心
        For i = 1 To 4
        If i <> 2 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 3 Then
                
                If e = 3 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 3 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 3 Then
                    If f = 3 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 2 Then WhoWin = i
                    If WhoWin <> 2 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 2 Then     '電腦1出方塊
        For i = 1 To 4
        If i <> 2 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 2 Then
                
                If e = 2 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 2 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 2 Then
                    If f = 2 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 2 Then WhoWin = i
                    If WhoWin <> 2 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 1 Then     '電腦1出梅花
        For i = 1 To 4
        If i <> 2 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 1 Then
                
                If e = 1 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 1 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 1 Then
                    If f = 1 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 2 Then WhoWin = i
                    If WhoWin <> 2 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
End If

If FirstPlayer = 3 Then     '電腦2先出牌

    WhoWin = 3
    RndComAI22
    If AIPutCard(3) < 53 Then FirstCard = 4
    If AIPutCard(3) < 40 Then FirstCard = 3
    If AIPutCard(3) < 27 Then FirstCard = 2
    If AIPutCard(3) < 14 Then FirstCard = 1
    
    If GameKing = 5 Then GoTo noking
    
    If FirstCard = 4 Then     '電腦2出黑桃
        For i = 1 To 4
        If i <> 3 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 4 Then
                
                If e = 4 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 4 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 4 Then
                    If f = 4 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 3 Then WhoWin = i
                    If WhoWin <> 3 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 3 Then     '電腦2出紅心
        For i = 1 To 4
        If i <> 3 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 3 Then
                
                If e = 3 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 3 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 3 Then
                    If f = 3 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 3 Then WhoWin = i
                    If WhoWin <> 3 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 2 Then     '電腦2出方塊
        For i = 1 To 4
        If i <> 3 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 2 Then
                
                If e = 2 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 2 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 2 Then
                    If f = 2 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 3 Then WhoWin = i
                    If WhoWin <> 3 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 1 Then     '電腦2出梅花
        For i = 1 To 4
        If i <> 3 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 1 Then
                
                If e = 1 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 1 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 1 Then
                    If f = 1 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 3 Then WhoWin = i
                    If WhoWin <> 3 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
End If

If FirstPlayer = 4 Then     '電腦3先出牌

    WhoWin = 4
    RndComAI33
    If AIPutCard(4) < 53 Then FirstCard = 4
    If AIPutCard(4) < 40 Then FirstCard = 3
    If AIPutCard(4) < 27 Then FirstCard = 2
    If AIPutCard(4) < 14 Then FirstCard = 1
    
    If GameKing = 5 Then GoTo noking
    
    If FirstCard = 4 Then     '電腦3出黑桃
        For i = 1 To 4
        If i <> 4 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 4 Then
                
                If e = 4 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 4 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 4 Then
                    If f = 4 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 4 Then WhoWin = i
                    If WhoWin <> 4 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 3 Then     '電腦3出紅心
        For i = 1 To 4
        If i <> 4 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 3 Then
                
                If e = 3 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 3 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 3 Then
                    If f = 3 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 4 Then WhoWin = i
                    If WhoWin <> 4 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 2 Then     '電腦3出方塊
        For i = 1 To 4
        If i <> 4 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 2 Then
                
                If e = 2 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 2 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 2 Then
                    If f = 2 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 4 Then WhoWin = i
                    If WhoWin <> 4 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
    If FirstCard = 1 Then     '電腦3出梅花
        For i = 1 To 4
        If i <> 4 Then
            If AIPutCard(i) < 53 Then e = 4
            If AIPutCard(i) < 40 Then e = 3
            If AIPutCard(i) < 27 Then e = 2
            If AIPutCard(i) < 14 Then e = 1
            
            If GameKing = 1 Then
                
                If e = 1 Then
                    If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                End If
                
            End If
            
            If GameKing <> 1 Then
                        
                If AIPutCard(WhoWin) < 53 Then f = 4
                If AIPutCard(WhoWin) < 40 Then f = 3
                If AIPutCard(WhoWin) < 27 Then f = 2
                If AIPutCard(WhoWin) < 14 Then f = 1
                
                If e = 1 Then
                    If f = 1 Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                End If
                
                If e = GameKing Then
                    
                    If WhoWin = 4 Then WhoWin = i
                    If WhoWin <> 4 Then
                        
                        If f = GameKing Then
                            If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                        End If
                        
                        If f <> GameKing Then WhoWin = i
                        
                    End If
                    
                End If
                
            End If
            
        End If
        Next i
        
    End If
    
End If


noking:
    If GameKing = 5 Then
        
        If FirstPlayer = 1 Then
            
            For i = 1 To 4
            
                If i <> 1 Then
                    
                    If AIPutCard(i) < 53 Then e = 4
                    If AIPutCard(i) < 40 Then e = 3
                    If AIPutCard(i) < 27 Then e = 2
                    If AIPutCard(i) < 14 Then e = 1
                    
                    If e = FirstCard Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                    
                End If
                
            Next i
            
        End If
        
        If FirstPlayer = 2 Then
            
            For i = 1 To 4
            
                If i <> 2 Then
                    
                    If AIPutCard(i) < 53 Then e = 4
                    If AIPutCard(i) < 40 Then e = 3
                    If AIPutCard(i) < 27 Then e = 2
                    If AIPutCard(i) < 14 Then e = 1
                    
                    If e = FirstCard Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                    
                End If
                
            Next i
            
        End If
        
        If FirstPlayer = 3 Then
            
            For i = 1 To 4
            
                If i <> 3 Then
                    
                    If AIPutCard(i) < 53 Then e = 4
                    If AIPutCard(i) < 40 Then e = 3
                    If AIPutCard(i) < 27 Then e = 2
                    If AIPutCard(i) < 14 Then e = 1
                    
                    If e = FirstCard Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                    
                End If
                
            Next i
            
        End If
        
        If FirstPlayer = 4 Then
            
            For i = 1 To 4
            
                If i <> 4 Then
                    
                    If AIPutCard(i) < 53 Then e = 4
                    If AIPutCard(i) < 40 Then e = 3
                    If AIPutCard(i) < 27 Then e = 2
                    If AIPutCard(i) < 14 Then e = 1
                    
                    If e = FirstCard Then
                        If AIPutCard(i) > AIPutCard(WhoWin) Then WhoWin = i
                    End If
                    
                End If
                
            Next i
            
        End If
        
    End If



If WhoWin Mod 2 <> 0 Then AWin = AWin + 1
If WhoWin Mod 2 = 0 Then BWin = BWin + 1

If AWin <> 0 Then
    BitBlt Form1.Picture6.hdc, 12, 60, 15, 20, Form1.Picture8.hdc, AWin * 15, 0, srccopy
End If
If BWin <> 0 Then
    BitBlt Form1.Picture7.hdc, 12, 60, 15, 20, Form1.Picture8.hdc, BWin * 15, 0, srccopy
End If

Form1.Picture6.Refresh
Form1.Picture7.Refresh
Form1.Picture1.Refresh

If AWin = needwin(1) Then GoTo over
If BWin = needwin(2) Then GoTo over

delay GameSpeed

Form1.Picture1.Line (95, 100)-(405, 293), &H8000&, BF
Form1.Picture1.Refresh

delay GameSpeed \ 5
    
FirstPlayer = WhoWin
WhoIsFirst

over:
If AWin = needwin(1) Then GameOver
If BWin = needwin(2) Then GameOver


End Sub

Sub GameOver()     '遊戲結束

If AWin = needwin(1) Then
    BitBlt Form1.Picture6.hdc, 0, 1, 40, 25, Form1.Picture9.hdc, 0, 0, srccopy
    BitBlt Form1.Picture7.hdc, 0, 1, 40, 25, Form1.Picture9.hdc, 40, 0, srccopy
End If
If BWin = needwin(2) Then
    BitBlt Form1.Picture6.hdc, 0, 1, 40, 25, Form1.Picture9.hdc, 40, 0, srccopy
    BitBlt Form1.Picture7.hdc, 0, 1, 40, 25, Form1.Picture9.hdc, 0, 0, srccopy
End If

Form1.Picture6.Refresh
Form1.Picture7.Refresh

GameKing = 0
ReturnPlayer = 0

Form4.Show

WhoWin = 0

Exit Sub

End Sub

Sub RndComAI33()     '電腦3先的情況   玩家丟完換電腦1&2

If AIPutCard(4) < 53 Then b = 4
If AIPutCard(4) < 40 Then b = 3
If AIPutCard(4) < 27 Then b = 2
If AIPutCard(4) < 14 Then b = 1

For i = 2 To 3

j = 1
Do While j <> 14 '同花色先丟
              
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        If users(i, j) <> 0 Then
            If a = b Then
                AIPutCard(i) = users(i, j)
                x = j
            End If
        End If
        j = j + 1
'        If j = 14 Then Exit Do
    Loop
    If AIPutCard(i) <> 0 Then users(i, x) = 0
    
    j = 1
    Do While AIPutCard(i) = 0     '沒有同花色先丟王
        
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        
        If a = GameKing Then
            AIPutCard(i) = users(i, j)
            users(i, j) = 0
        End If
        j = j + 1
        If j = 14 Then Exit Do
    Loop
    
    Do While AIPutCard(i) = 0     '沒王梅同花色隨便丟
    
        y = Int(Rnd * 13) + 1
        If users(i, y) <> 0 Then
            AIPutCard(i) = users(i, y)
            users(i, y) = 0
        End If
        
    Loop
    
    ComShowCards i, AIPutCard(i)
    
Next i

End Sub

Sub RndComAI3()     '電腦3先的情況

For i = 2 To 4
    AIPutCard(i) = 0
Next i

Do While AIPutCard(4) = 0     '電腦3要丟什麼牌
    x = Int(Rnd * 13) + 1
    If users(4, x) <> 0 Then
        AIPutCard(4) = users(4, x)
        users(4, x) = 0
    End If
Loop
    
    ComShowCards 4, AIPutCard(4)
    

End Sub

Sub RndComAI22()     '電腦2先的情況   玩家丟完換電腦1

If AIPutCard(3) < 53 Then b = 4
If AIPutCard(3) < 40 Then b = 3
If AIPutCard(3) < 27 Then b = 2
If AIPutCard(3) < 14 Then b = 1

j = 1
Do While j <> 14 '同花色先丟
              
        If users(2, j) < 53 Then a = 4
        If users(2, j) < 40 Then a = 3
        If users(2, j) < 27 Then a = 2
        If users(2, j) < 14 Then a = 1
        If users(2, j) <> 0 Then
            If a = b Then
                AIPutCard(2) = users(2, j)
                x = j
            End If
        End If
        j = j + 1
'        If j = 14 Then Exit Do
Loop
    If AIPutCard(2) <> 0 Then users(2, x) = 0
    
    j = 1
    Do While AIPutCard(2) = 0     '沒有同花色先丟王
        
        If users(2, j) < 53 Then a = 4
        If users(2, j) < 40 Then a = 3
        If users(2, j) < 27 Then a = 2
        If users(2, j) < 14 Then a = 1
        
        If a = GameKing Then
            AIPutCard(2) = users(2, j)
            users(2, j) = 0
        End If
        j = j + 1
        If j = 14 Then Exit Do
    Loop
    
    Do While AIPutCard(2) = 0     '沒王梅同花色隨便丟
    
        y = Int(Rnd * 13) + 1
        If users(2, y) <> 0 Then
            AIPutCard(2) = users(2, y)
            users(2, y) = 0
        End If
        
    Loop
    
    ComShowCards 2, AIPutCard(2)

    

End Sub

Sub RndComAI2()     '電腦2先的情況

For i = 2 To 4
    AIPutCard(i) = 0
Next i

Do While AIPutCard(3) = 0     '電腦2要丟什麼牌
    x = Int(Rnd * 13) + 1
    If users(3, x) <> 0 Then
        AIPutCard(3) = users(3, x)
        users(3, x) = 0
    End If
Loop
    
    ComShowCards 3, AIPutCard(3)

If AIPutCard(3) < 53 Then b = 4
If AIPutCard(3) < 40 Then b = 3
If AIPutCard(3) < 27 Then b = 2
If AIPutCard(3) < 14 Then b = 1
    
    j = 1
    Do While j <> 14   '同花色先丟
              
        If users(4, j) < 53 Then a = 4
        If users(4, j) < 40 Then a = 3
        If users(4, j) < 27 Then a = 2
        If users(4, j) < 14 Then a = 1
        If users(4, j) <> 0 Then
            If a = b Then
                AIPutCard(4) = users(4, j)
                x = j
            End If
        End If
        j = j + 1
'        If j = 14 Then Exit Do
    Loop
    If AIPutCard(4) <> 0 Then users(4, x) = 0
    
    j = 1
    Do While AIPutCard(4) = 0     '沒有同花色先丟王
        
        If users(4, j) < 53 Then a = 4
        If users(4, j) < 40 Then a = 3
        If users(4, j) < 27 Then a = 2
        If users(4, j) < 14 Then a = 1
        
        If a = GameKing Then
            AIPutCard(4) = users(4, j)
            users(4, j) = 0
        End If
        j = j + 1
        If j = 14 Then Exit Do
    Loop
    
    Do While AIPutCard(4) = 0     '沒王梅同花色隨便丟
    
        y = Int(Rnd * 13) + 1
        If users(4, y) <> 0 Then
            AIPutCard(4) = users(4, y)
            users(4, y) = 0
        End If
        
    Loop
    
    ComShowCards 4, AIPutCard(4)
    
'Next i

End Sub

Sub RndComAI1()     '電腦1先的情況

For i = 2 To 4
    AIPutCard(i) = 0
Next i

Do While AIPutCard(2) = 0     '電腦1要丟什麼牌
    x = Int(Rnd * 13) + 1
    If users(2, x) <> 0 Then
        AIPutCard(2) = users(2, x)
        users(2, x) = 0
    End If
Loop
    
    ComShowCards 2, AIPutCard(2)

If AIPutCard(2) < 53 Then b = 4
If AIPutCard(2) < 40 Then b = 3
If AIPutCard(2) < 27 Then b = 2
If AIPutCard(2) < 14 Then b = 1


For i = 3 To 4
    
    j = 1
    Do While j <> 14   '同花色先丟
              
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        If users(i, j) <> 0 Then
            If a = b Then
                AIPutCard(i) = users(i, j)
                x = j
            End If
        End If
        j = j + 1
'        If j = 14 Then Exit Do
    Loop
    If AIPutCard(i) <> 0 Then users(i, x) = 0
    
    j = 1
    Do While AIPutCard(i) = 0     '沒有同花色先丟王
        
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        
        If a = GameKing Then
            AIPutCard(i) = users(i, j)
            users(i, j) = 0
        End If
        j = j + 1
        If j = 14 Then Exit Do
    Loop
    
    Do While AIPutCard(i) = 0     '沒王梅同花色隨便丟
    
        y = Int(Rnd * 13) + 1
        If users(i, y) <> 0 Then
            AIPutCard(i) = users(i, y)
            users(i, y) = 0
        End If
        
    Loop
    
    ComShowCards i, AIPutCard(i)
    
Next i

End Sub

Sub ComAI(num)     '電腦判斷要丟的牌

For i = 2 To 4
    AIPutCard(i) = 0
Next i

If num < 53 Then b = 4
If num < 40 Then b = 3
If num < 27 Then b = 2
If num < 14 Then b = 1


For i = 2 To 4
    
    j = 1
    Do While j <> 14   '同花色先丟
              
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        If users(i, j) <> 0 Then
            If a = b Then
                AIPutCard(i) = users(i, j)
                x = j
            End If
        End If
        j = j + 1
'        If j = 14 Then Exit Do
    Loop
    If AIPutCard(i) <> 0 Then users(i, x) = 0
    
    j = 1
    Do While AIPutCard(i) = 0     '沒有同花色先丟王
        
        If users(i, j) < 53 Then a = 4
        If users(i, j) < 40 Then a = 3
        If users(i, j) < 27 Then a = 2
        If users(i, j) < 14 Then a = 1
        
        If a = GameKing Then
            AIPutCard(i) = users(i, j)
            users(i, j) = 0
        End If
        j = j + 1
        If j = 14 Then Exit Do
    Loop
    
    Do While AIPutCard(i) = 0     '沒王梅同花色隨便丟
    
        y = Int(Rnd * 13) + 1
        If users(i, y) <> 0 Then
            AIPutCard(i) = users(i, y)
            users(i, y) = 0
        End If
        
    Loop
    
    ComShowCards i, AIPutCard(i)
    
Next i

End Sub

Sub ComShowCards(i, num)

ComDrawCards i, num
ComDraw i, ComPlay

CheckComplay = CheckComplay + 1

If CheckComplay = 3 Then
    ComPlay = ComPlay + 1
    CheckComplay = 0
End If

End Sub

Sub DrawCards(x, y, number)  '畫牌子

a = number Mod 13
If a = 0 Then
a = 13
End If

If number < 53 Then b = 4
If number < 40 Then b = 3
If number < 27 Then b = 2
If number < 14 Then b = 1

Select Case b
  Case 1
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 67, srccopy
  Case 2
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 71, 164, srccopy
  Case 3
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 261, srccopy
  Case 4
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 164, srccopy
End Select

If b = 1 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 28, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 42, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 4 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 28, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 42, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 2 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 0, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 14, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 3 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 0, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 14, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If

End Sub

Sub DrawBack(x, y)  '畫牌子背面
BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture4.hdc, 0, 0, srccopy
End Sub

Sub Draw(x, y, num)  '清除牌子
   Form1.Picture1.Line (x, y)-(x + 70, y + 95), &H8000&, BF
End Sub

Sub ComDrawCards(com, numbers)   '電腦出牌

a = numbers Mod 13
If a = 0 Then a = 13

If numbers < 53 Then b = 4
If numbers < 40 Then b = 3
If numbers < 27 Then b = 2
If numbers < 14 Then b = 1

If com = 2 Then    '95 155
x = 120
y = 155
End If

If com = 3 Then    '230 100
x = 230
y = 102
End If

If com = 4 Then    '
x = 335
y = 155
End If

Select Case b
  Case 1
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 67, srccopy
  Case 2
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 71, 164, srccopy
  Case 3
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 261, srccopy
  Case 4
    BitBlt Form1.Picture1.hdc, x, y, 71, 96, Form1.Picture2.hdc, 0, 164, srccopy
End Select

If b = 1 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 28, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 42, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 4 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 28, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 42, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 2 Then
    BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 0, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 14, srccopy
    BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
    BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If
If b = 3 Then
   BitBlt Form1.Picture1.hdc, x + 1, y + 4, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 0, srccopy
   BitBlt Form1.Picture1.hdc, x + 59, y + 79, 11, 13, Form1.Picture2.hdc, (a - 1) * 12, 14, srccopy
   BitBlt Form1.Picture1.hdc, x + 1, y + 17, 11, 10, Form1.Picture2.hdc, (b - 1) * 12, 56, srccopy
   BitBlt Form1.Picture1.hdc, x + 59, y + 69, 11, 10, Form1.Picture2.hdc, (b + 3) * 12, 56, srccopy
End If

End Sub

Sub ComDraw(i, num)   '電腦消牌

If i = 2 Then
x = 20
y = 292
a = 1
End If

If i = 3 Then
x = 357
y = 5
a = 2
End If

If i = 4 Then
x = 440
y = 292
a = 1
End If

If num = 13 Then
If a = 1 Then Form1.Picture1.Line (x, y - (8 * 12) - 96)-(x + 71, y + 96), &H8000&, BF
If a = 2 Then Form1.Picture1.Line (x - (8 * 12) - 71, y)-(x + 71, y + 96), &H8000&, BF

End If

If a = 1 Then Form1.Picture1.Line (x, y - (8 * num))-(x + 70, y + 8), &H8000&, BF
If a = 2 Then Form1.Picture1.Line (x - (8 * num), y)-(x + 8, y + 96), &H8000&, BF


End Sub

Sub WhoIsFirst()     '誰先開始

        If FirstPlayer = 1 Then
            ReturnPlayer = 1
            Exit Sub
        End If
        
        If FirstPlayer = 2 Then
            RndComAI1
            ReturnPlayer = 1
            Exit Sub
        End If
        
        If FirstPlayer = 3 Then
            RndComAI2
            ReturnPlayer = 1
            Exit Sub
        End If
        
        If FirstPlayer = 4 Then
            RndComAI3
            ReturnPlayer = 1
            Exit Sub
        End If
        
End Sub

Sub gametostart()     '叫王結束準備開始

Unload Form8

'For i = 0 To 3
'    Form8.Label1(i).Caption = 0
'Next i

For i = 1 To 4     '畫出王
    If ComKing(i) <> 36 Then
        b = ComKing(i) Mod 5
        If b = 0 Then b = 5
        GameKing = b
        If b = 1 Then BitBlt Form1.Picture5.hdc, 35, 0, 40, 40, Form1.Picture3.hdc, 40, 130, srccopy
        If b = 2 Then BitBlt Form1.Picture5.hdc, 35, 0, 40, 40, Form1.Picture3.hdc, 80, 130, srccopy
        If b = 3 Then BitBlt Form1.Picture5.hdc, 35, 0, 40, 40, Form1.Picture3.hdc, 0, 130, srccopy
        If b = 4 Then BitBlt Form1.Picture5.hdc, 35, 0, 40, 40, Form1.Picture3.hdc, 120, 130, srccopy
        If b = 5 Then BitBlt Form1.Picture5.hdc, 25, 0, 60, 40, Form1.Picture3.hdc, 120, 0, srccopy
        
        If i Mod 2 <> 0 Then     '決定雙方吃幾敦
            
            If ComKing(i) < 36 Then a = 7
            If ComKing(i) < 31 Then a = 6
            If ComKing(i) < 26 Then a = 5
            If ComKing(i) < 21 Then a = 4
            If ComKing(i) < 16 Then a = 3
            If ComKing(i) < 11 Then a = 2
            If ComKing(i) < 6 Then a = 1
            
            If a = 1 Then
                needwin(1) = 7
                needwin(2) = 7
            End If
            If a = 2 Then
                needwin(1) = 8
                needwin(2) = 6
            End If
            If a = 3 Then
                needwin(1) = 9
                needwin(2) = 5
            End If
            If a = 4 Then
                needwin(1) = 10
                needwin(2) = 4
            End If
            If a = 5 Then
                needwin(1) = 11
                needwin(2) = 3
            End If
            If a = 6 Then
                needwin(1) = 12
                needwin(2) = 2
            End If
            If a = 7 Then
                needwin(1) = 13
                needwin(2) = 1
            End If
            
        End If
        
        If i Mod 2 = 0 Then
            
            If ComKing(i) < 36 Then a = 7
            If ComKing(i) < 31 Then a = 6
            If ComKing(i) < 26 Then a = 5
            If ComKing(i) < 21 Then a = 4
            If ComKing(i) < 16 Then a = 3
            If ComKing(i) < 11 Then a = 2
            If ComKing(i) < 6 Then a = 1
            
            If a = 1 Then
                needwin(1) = 7
                needwin(2) = 7
            End If
            If a = 2 Then
                needwin(1) = 6
                needwin(2) = 8
            End If
            If a = 3 Then
                needwin(1) = 5
                needwin(2) = 9
            End If
            If a = 4 Then
                needwin(1) = 4
                needwin(2) = 10
            End If
            If a = 5 Then
                needwin(1) = 3
                needwin(2) = 11
            End If
            If a = 6 Then
                needwin(1) = 2
                needwin(2) = 12
            End If
            If a = 7 Then
                needwin(1) = 1
                needwin(2) = 13
            End If
            
        End If
        
        If i = 1 Then FirstPlayer = 2     '決定誰先開始
        If i = 2 Then FirstPlayer = 3
        If i = 3 Then FirstPlayer = 4
        If i = 4 Then FirstPlayer = 1
        
    End If
Next i

BitBlt Form1.Picture6.hdc, 12, 115, 15, 20, Form1.Picture8.hdc, needwin(1) * 15, 3, srccopy
Form1.Picture6.Refresh

BitBlt Form1.Picture7.hdc, 12, 115, 15, 20, Form1.Picture8.hdc, needwin(2) * 15, 3, srccopy
Form1.Picture7.Refresh


Form1.Picture5.Refresh

'For i = 0 To 35
'    Form8.Command1(i).Enabled = False
'Next i
'Form1.Command3(0).Enabled = False

delay 10

Form1.Picture1.Line (110, 120)-(440, 290), &H8000&, BF
Form1.Picture1.Refresh
        
WhoIsFirst

End Sub

Sub WhoIsKing(king)   '決定王牌

If king = 36 Then
    If ComKing(1) = 0 Then GoTo out
End If

Form1.Picture1.Line (110, 120)-(440, 290), &H8000&, BF
Form1.Picture1.Refresh

If pass < 3 Then pass = 0

If king = 36 Then pass = pass + 1

If pass = 3 Then gametostart: Exit Sub

AIKing king
DrewKing king

For i = 2 To 4     '把電腦喊的按鈕都取消
    If ComKing(i) <> 36 Then
        For j = 0 To ComKing(i) - 1
            Form8.Command1(j).Enabled = False
        Next j
    End If
Next i

For i = 2 To 4
    If ComKing(i) = 36 Then pass = pass + 1
Next i
If pass >= 3 Then gametostart: Exit Sub

out:
End Sub

Sub DrewKing(king)      '畫叫的王

'玩家的
If king < 36 Then a = 7
If king < 31 Then a = 6
If king < 26 Then a = 5
If king < 21 Then a = 4
If king < 16 Then a = 3
If king < 11 Then a = 2
If king < 6 Then a = 1
If king = 36 Then a = 8
b = king Mod 5
If b = 0 Then b = 5
If a = 8 Then b = 0

If b = 0 Then
    BitBlt Form1.Picture1.hdc, 230, 250, 100, 40, Form1.Picture3.hdc, 0, 0, srccopy
End If

If b = 1 Then
    BitBlt Form1.Picture1.hdc, 270, 250, 40, 40, Form1.Picture3.hdc, 40, 130, srccopy
    If a = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If a = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If a = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If a = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If a = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If a = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If a = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

If b = 2 Then
    BitBlt Form1.Picture1.hdc, 270, 250, 40, 40, Form1.Picture3.hdc, 80, 130, srccopy
    If a = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If a = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If a = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If a = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If a = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If a = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If a = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If b = 3 Then
    BitBlt Form1.Picture1.hdc, 270, 250, 40, 40, Form1.Picture3.hdc, 0, 130, srccopy
    If a = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If a = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If a = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If a = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If a = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If a = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If a = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If b = 4 Then
    BitBlt Form1.Picture1.hdc, 270, 250, 40, 40, Form1.Picture3.hdc, 120, 130, srccopy
    If a = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If a = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If a = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If a = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If a = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If a = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If a = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 250, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

If b = 5 Then
    BitBlt Form1.Picture1.hdc, 260, 250, 60, 40, Form1.Picture3.hdc, 120, 0, srccopy
    If a = 1 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If a = 2 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If a = 3 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If a = 4 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If a = 5 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If a = 6 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If a = 7 Then
        BitBlt Form1.Picture1.hdc, 230, 250, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

'電腦1
If ComKing(2) < 36 Then c = 7
If ComKing(2) < 31 Then c = 6
If ComKing(2) < 26 Then c = 5
If ComKing(2) < 21 Then c = 4
If ComKing(2) < 16 Then c = 3
If ComKing(2) < 11 Then c = 2
If ComKing(2) < 6 Then c = 1
If ComKing(2) = 36 Then c = 8
d = ComKing(2) Mod 5
If d = 0 Then d = 5
If c = 8 Then d = 0

If d = 0 Then
    BitBlt Form1.Picture1.hdc, 110, 170, 100, 40, Form1.Picture3.hdc, 0, 0, srccopy
End If

If d = 1 Then
    BitBlt Form1.Picture1.hdc, 150, 170, 40, 40, Form1.Picture3.hdc, 40, 130, srccopy
    If c = 1 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If c = 2 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If c = 3 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If c = 4 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If c = 5 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If c = 6 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If c = 7 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

If d = 2 Then
    BitBlt Form1.Picture1.hdc, 150, 170, 40, 40, Form1.Picture3.hdc, 80, 130, srccopy
    If c = 1 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If c = 2 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If c = 3 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If c = 4 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If c = 5 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If c = 6 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If c = 7 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If d = 3 Then
    BitBlt Form1.Picture1.hdc, 150, 170, 40, 40, Form1.Picture3.hdc, 0, 130, srccopy
    If c = 1 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If c = 2 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If c = 3 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If c = 4 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If c = 5 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If c = 6 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If c = 7 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If d = 4 Then
    BitBlt Form1.Picture1.hdc, 150, 170, 40, 40, Form1.Picture3.hdc, 120, 130, srccopy
    If c = 1 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If c = 2 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If c = 3 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If c = 4 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If c = 5 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If c = 6 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If c = 7 Then
        BitBlt Form1.Picture1.hdc, 120, 170, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

'電腦2
If ComKing(3) < 36 Then e = 7
If ComKing(3) < 31 Then e = 6
If ComKing(3) < 26 Then e = 5
If ComKing(3) < 21 Then e = 4
If ComKing(3) < 16 Then e = 3
If ComKing(3) < 11 Then e = 2
If ComKing(3) < 6 Then e = 1
If ComKing(3) = 36 Then e = 8
f = ComKing(3) Mod 5
If f = 0 Then f = 5
If e = 8 Then f = 0

If f = 0 Then
    BitBlt Form1.Picture1.hdc, 230, 120, 100, 40, Form1.Picture3.hdc, 0, 0, srccopy
End If

If f = 1 Then
    BitBlt Form1.Picture1.hdc, 270, 120, 40, 40, Form1.Picture3.hdc, 40, 130, srccopy
    If e = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If e = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If e = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If e = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If e = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If e = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If e = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

If f = 2 Then
    BitBlt Form1.Picture1.hdc, 270, 120, 40, 40, Form1.Picture3.hdc, 80, 130, srccopy
    If e = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If e = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If e = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If e = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If e = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If e = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If e = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If f = 3 Then
    BitBlt Form1.Picture1.hdc, 270, 120, 40, 40, Form1.Picture3.hdc, 0, 130, srccopy
    If e = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If e = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If e = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If e = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If e = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If e = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If e = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If f = 4 Then
    BitBlt Form1.Picture1.hdc, 270, 120, 40, 40, Form1.Picture3.hdc, 120, 130, srccopy
    If e = 1 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If e = 2 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If e = 3 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If e = 4 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If e = 5 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If e = 6 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If e = 7 Then
        BitBlt Form1.Picture1.hdc, 240, 120, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

'電腦3
If ComKing(4) < 36 Then g = 7
If ComKing(4) < 31 Then g = 6
If ComKing(4) < 26 Then g = 5
If ComKing(4) < 21 Then g = 4
If ComKing(4) < 16 Then g = 3
If ComKing(4) < 11 Then g = 2
If ComKing(4) < 6 Then g = 1
If ComKing(4) = 36 Then g = 8
h = ComKing(4) Mod 5
If h = 0 Then h = 5
If g = 8 Then h = 0

If h = 0 Then
    BitBlt Form1.Picture1.hdc, 340, 170, 100, 40, Form1.Picture3.hdc, 0, 0, srccopy
End If

If h = 1 Then
    BitBlt Form1.Picture1.hdc, 390, 170, 40, 40, Form1.Picture3.hdc, 40, 130, srccopy
    If g = 1 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If g = 2 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If g = 3 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If g = 4 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If g = 5 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If g = 6 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If g = 7 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If

If h = 2 Then
    BitBlt Form1.Picture1.hdc, 390, 170, 40, 40, Form1.Picture3.hdc, 80, 130, srccopy
    If g = 1 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If g = 2 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If g = 3 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If g = 4 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If g = 5 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If g = 6 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If g = 7 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If h = 3 Then
    BitBlt Form1.Picture1.hdc, 390, 170, 40, 40, Form1.Picture3.hdc, 0, 130, srccopy
    If g = 1 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 0, 45, srccopy
    End If
    If g = 2 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 30, 45, srccopy
    End If
    If g = 3 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 60, 45, srccopy
    End If
    If g = 4 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 90, 45, srccopy
    End If
    If g = 5 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 120, 45, srccopy
    End If
    If g = 6 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 150, 45, srccopy
    End If
    If g = 7 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 180, 45, srccopy
    End If
End If

If h = 4 Then
    BitBlt Form1.Picture1.hdc, 390, 170, 40, 40, Form1.Picture3.hdc, 120, 130, srccopy
    If g = 1 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 0, 85, srccopy
    End If
    If g = 2 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 30, 85, srccopy
    End If
    If g = 3 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 60, 85, srccopy
    End If
    If g = 4 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 90, 85, srccopy
    End If
    If g = 5 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 120, 85, srccopy
    End If
    If g = 6 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 150, 85, srccopy
    End If
    If g = 7 Then
        BitBlt Form1.Picture1.hdc, 360, 170, 30, 35, Form1.Picture3.hdc, 180, 85, srccopy
    End If
End If
    
    

Form1.Picture1.Refresh

End Sub

Sub AIKing(king)   '電腦叫王

If king < 36 Then a = 7
If king < 31 Then a = 6
If king < 26 Then a = 5
If king < 21 Then a = 4
If king < 16 Then a = 3
If king < 11 Then a = 2
If king < 6 Then a = 1
If king = 36 Then a = 0
b = king Mod 5
If b = 0 Then b = 5
If a = 0 Then GoTo playerpass

If king <> 36 Then
    ComKing(1) = king
End If

For i = 2 To 4

    If BigKing(i, 3) >= a Then     '電腦能喊到的數字比目前玩家喊的大
    
        If BigKing(i, 1) > b Then      '電腦要喊的花色筆玩家喊的大
            If a = 1 Then
                ComKing(i) = BigKing(i, 1)
            End If
            If a <> 1 Then
                ComKing(i) = ((a - 1) * 5) + BigKing(i, 1)
            End If
        End If
        
        If BigKing(i, 1) = b Then     '電腦要喊的跟玩家一樣   PASS
            ComKing(i) = 36
        End If
        
        If BigKing(i, 1) < b Then     '電腦要喊的筆玩家小
            If BigKing(i, 3) >= a + 1 Then     '看電腦能不能增加數字   能:喊
                ComKing(i) = a * 5 + BigKing(i, 1)
            End If
            If BigKing(i, 3) < a + 1 Then     '看電腦能不能增加數字   不能:PASS
                ComKing(i) = 36
            End If
        End If
    End If
        
    If BigKing(i, 3) < a Then     '電腦能喊的數字比目前小   PASS
        ComKing(i) = 36
    End If
    
Next i

If ComKing(2) <> 0 Then
If ComKing(3) < ComKing(2) Then     '讓第2個電腦教的數字比第1個大
    x = ComKing(3) + 5
    If x < 36 Then x = 7
    If x < 26 Then x = 5
    If x < 21 Then x = 4
    If x < 16 Then x = 3
    If x < 11 Then x = 2
    If x < 6 Then x = 1
    
    If x <= BigKing(3, 3) Then
        ComKing(3) = ComKing(3) + 5
    End If
    If x > BigKing(3, 3) Then
        ComKing(3) = 36
    End If
End If
End If

If ComKing(3) <> 0 Then
If ComKing(4) < ComKing(3) Then     '讓第4個電腦教的數字比第3個大
    y = ComKing(4) + 5
    If y < 36 Then y = 7
    If y < 26 Then y = 5
    If y < 21 Then y = 4
    If y < 16 Then y = 3
    If y < 11 Then y = 2
    If y < 6 Then y = 1
    
    If y <= BigKing(4, 3) Then
        ComKing(4) = ComKing(4) + 5
    End If
    If y > BigKing(4, 3) Then
        ComKing(4) = 36
    End If
End If
End If

If ComKing(2) <> 0 Then
If ComKing(4) < ComKing(2) Then     '讓第4個電腦教的數字比第3個大
    y = ComKing(4) + 5
    If y < 36 Then y = 7
    If y < 26 Then y = 5
    If y < 21 Then y = 4
    If y < 16 Then y = 3
    If y < 11 Then y = 2
    If y < 6 Then y = 1
    
    If y <= BigKing(4, 3) Then
        ComKing(4) = ComKing(4) + 5
    End If
    If y > BigKing(4, 3) Then
        ComKing(4) = 36
    End If
End If
End If

If BigKing(3, 1) = BigKing(2, 1) Then ComKing(3) = 36
If BigKing(4, 1) = BigKing(3, 1) Then ComKing(4) = 36
If BigKing(4, 1) = BigKing(2, 1) Then ComKing(4) = 36

playerpass:
If a = 0 Then
    
    If ComKing(2) <> 36 Then     '2跟34比
        
        If ComKing(4) <> 36 Then
            
            If ComKing(4) < 36 Then a = 7
            If ComKing(4) < 31 Then a = 6
            If ComKing(4) < 26 Then a = 5
            If ComKing(4) < 21 Then a = 4
            If ComKing(4) < 16 Then a = 3
            If ComKing(4) < 11 Then a = 2
            If ComKing(4) < 6 Then a = 1
            b = BigKing(4, 1)
            
            If BigKing(2, 3) >= a Then
                
                If BigKing(2, 1) > b Then
                    ComKing(2) = ComKing(4) + BigKing(2, 1) - b
                End If
                
                If BigKing(2, 1) < b Then
                    
                    If BigKing(2, 3) >= a + 1 Then
                        ComKing(2) = ComKing(2) + 5
                    End If
                    
                    If BigKing(2, 3) < a + 1 Then
                        ComKing(2) = 36
                    End If
                End If
            End If
            
            If BigKing(2, 3) < a Then
                ComKing(2) = 36
            End If
        End If
        
        If ComKing(4) = 36 Then
            
            If ComKing(3) <> 36 Then
                
                If ComKing(3) < 36 Then a = 7
                If ComKing(3) < 31 Then a = 6
                If ComKing(3) < 26 Then a = 5
                If ComKing(3) < 21 Then a = 4
                If ComKing(3) < 16 Then a = 3
                If ComKing(3) < 11 Then a = 2
                If ComKing(3) < 6 Then a = 1
                b = BigKing(3, 1)
            
                If BigKing(2, 3) >= a Then
                   
                    If BigKing(2, 1) > b Then
                        ComKing(2) = ComKing(3) + BigKing(2, 1) - b
                    End If
                
                    If BigKing(2, 1) < b Then
                
                        If BigKing(2, 3) >= a + 1 Then
                            ComKing(2) = ComKing(2) + 5
                        End If
                    
                        If BigKing(2, 3) < a + 1 Then
                            ComKing(2) = 36
                        End If
                    End If
                End If
            
                If BigKing(2, 3) < a Then
                    ComKing(2) = 36
                End If
            End If
        End If
    End If
    
    If ComKing(3) <> 36 Then     '3跟24比
        
        If ComKing(2) <> 36 Then
            
            If ComKing(2) < 36 Then a = 7
            If ComKing(2) < 31 Then a = 6
            If ComKing(2) < 26 Then a = 5
            If ComKing(2) < 21 Then a = 4
            If ComKing(2) < 16 Then a = 3
            If ComKing(2) < 11 Then a = 2
            If ComKing(2) < 6 Then a = 1
            b = BigKing(2, 1)
            
            If BigKing(3, 3) >= a Then
                
                If BigKing(3, 1) > b Then
                    ComKing(3) = ComKing(2) + BigKing(3, 1) - b
                End If
                
                If BigKing(3, 1) < b Then
                    
                    If BigKing(3, 3) >= a + 1 Then
                        ComKing(3) = ComKing(3) + 5
                    End If
                    
                    If BigKing(3, 3) < a + 1 Then
                        ComKing(3) = 36
                    End If
                End If
            End If
            
            If BigKing(3, 3) < a Then
                ComKing(3) = 36
            End If
        End If
        
        If ComKing(2) = 36 Then
            
            If ComKing(4) <> 36 Then
                
                If ComKing(4) < 36 Then a = 7
                If ComKing(4) < 31 Then a = 6
                If ComKing(4) < 26 Then a = 5
                If ComKing(4) < 21 Then a = 4
                If ComKing(4) < 16 Then a = 3
                If ComKing(4) < 11 Then a = 2
                If ComKing(4) < 6 Then a = 1
                b = BigKing(4, 1)
            
                If BigKing(3, 3) >= a Then
                   
                    If BigKing(3, 1) > b Then
                        ComKing(3) = ComKing(4) + BigKing(3, 1) - b
                    End If
                
                    If BigKing(3, 1) < b Then
                
                        If BigKing(3, 3) >= a + 1 Then
                            ComKing(3) = ComKing(3) + 5
                        End If
                    
                        If BigKing(3, 3) < a + 1 Then
                            ComKing(3) = 36
                        End If
                    End If
                End If
            
                If BigKing(3, 3) < a Then
                    ComKing(3) = 36
                End If
            End If
        End If
    End If
    
    If ComKing(4) <> 36 Then     '4跟23比
        
        If ComKing(3) <> 36 Then
            
            If ComKing(3) < 36 Then a = 7
            If ComKing(3) < 31 Then a = 6
            If ComKing(3) < 26 Then a = 5
            If ComKing(3) < 21 Then a = 4
            If ComKing(3) < 16 Then a = 3
            If ComKing(3) < 11 Then a = 2
            If ComKing(3) < 6 Then a = 1
            b = BigKing(3, 1)
            
            If BigKing(4, 3) >= a Then
                
                If BigKing(4, 1) > b Then
                    ComKing(4) = ComKing(3) + BigKing(4, 1) - b
                End If
                
                If BigKing(4, 1) < b Then
                    
                    If BigKing(4, 3) >= a + 1 Then
                        ComKing(4) = ComKing(4) + 5
                    End If
                    
                    If BigKing(4, 3) < a + 1 Then
                        ComKing(4) = 36
                    End If
                End If
            End If
            
            If BigKing(4, 3) < a Then
                ComKing(4) = 36
            End If
        End If
        
        If ComKing(3) = 36 Then
            
            If ComKing(2) <> 36 Then
                
                If ComKing(2) < 36 Then a = 7
                If ComKing(2) < 31 Then a = 6
                If ComKing(2) < 26 Then a = 5
                If ComKing(2) < 21 Then a = 4
                If ComKing(2) < 16 Then a = 3
                If ComKing(2) < 11 Then a = 2
                If ComKing(2) < 6 Then a = 1
                b = BigKing(2, 1)
            
                If BigKing(4, 3) >= a Then
                   
                    If BigKing(4, 1) > b Then
                        ComKing(4) = ComKing(2) + BigKing(4, 1) - b
                    End If
                
                    If BigKing(4, 1) < b Then
                
                        If BigKing(4, 3) >= a + 1 Then
                            ComKing(4) = ComKing(4) + 5
                        End If
                    
                        If BigKing(4, 3) < a + 1 Then
                            ComKing(4) = 36
                        End If
                    End If
                End If
            
                If BigKing(4, 3) < a Then
                    ComKing(4) = 36
                End If
            End If
        End If
    End If
         
End If

End Sub

Sub ClsScreen()  '清除桌面
Form1.Picture1.Line (95, 100)-(200, 100), &H8000&, BF
End Sub
Sub Swap(a, b)
c = a
a = b
b = c
End Sub

Sub delay(tim)
t = Timer
Do
If Timer - t > tim / 10 Then Exit Do
Loop
End Sub


Sub CloseEnd(ByVal hwnd As Long)   '將"X"除能
hMenu = GetSystemMenu(hwnd, 0)
 CloseStr = String(255, 0)

 Call GetMenuString(hMenu, SC_CLOSE, CloseStr, 256, MF_BYCOMMAND)
 CloseStr = Left(CloseStr, InStr(1, CloseStr, Chr(0)) - 1)

 Call DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
End Sub


