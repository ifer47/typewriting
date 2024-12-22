' 定义字符集，包含所有小写英文字母
Const charset = "abcdefghijklmnopqrstuvwxyz"

' 定义组合长度参数
Dim combinationLength

' 初始化数组来存储所有组合
Dim combinations()
Dim index
index = 0

' 生成所有可能的字母组合
Sub GenerateCombinations(length)
    ' 初始化索引
    index = 0
    Dim i, j, k, l
    For i = 1 To Len(charset)
    	TracePrint i
        If length = 1 Then
            ' 对于长度为1的组合，直接取每个字符
            ReDim Preserve combinations(index + 1)
            combinations(index) = Mid(charset, i, 1)
            index = index + 1
        ElseIf length = 2 Then
            For j = 1 To Len(charset)
                ReDim Preserve combinations(index + 1)
                combinations(index) = Mid(charset, i, 1) & Mid(charset, j, 1)
                index = index + 1
            Next
        ElseIf length = 3 Then
            For j = 1 To Len(charset)
                For k = 1 To Len(charset)
                    ReDim Preserve combinations(index + 1)
                    combinations(index) = Mid(charset, i, 1) & Mid(charset, j, 1) & Mid(charset, k, 1)
                    index = index + 1
                Next
            Next
        ElseIf length = 4 Then
            For j = 1 To Len(charset)
                For k = 1 To Len(charset)
                    For l = 1 To Len(charset)
                        ReDim Preserve combinations(index + 1)
                        combinations(index) = Mid(charset, i, 1) & Mid(charset, j, 1) & Mid(charset, k, 1) & Mid(charset, l, 1)
                        index = index + 1
                    Next
                Next
            Next
        End If
    Next
End Sub

' 处理组合
Sub ProcessCombinations(xPos, yPos)
    ' 移动鼠标到屏幕上的(xPos, yPos)位置
    MoveTo xPos, yPos
    ' 在当前鼠标位置执行一次鼠标左键点击
    LeftClick 1

    ' 定义候选数字，用于模拟输入数字1到candidate
    Const candidate = 9
    Dim letter, x, y, currentChar

    ' 遍历combinations数组中的每个组合
    For Each letter In combinations
        ' 输出当前处理的组合，用于调试和验证
        TracePrint letter
        Delay 100
        ' 遍历组合中的每个字符，并模拟按键操作
        For y = 1 To Len(letter)
            currentChar = Mid(letter, y, 1)
            KeyPress currentChar, 1
            Delay 100
        Next
        ' 模拟按键Enter，然后延迟100毫秒，再模拟按键Space
        KeyPress "Enter", 1
        Delay 100
        KeyPress "Space", 1
        ' 重复candidate次以下操作：模拟按键组合，模拟按键数字1到candidate
        For x = 1 To candidate
            Delay 100
            ' 遍历组合中的每个字符，并模拟按键操作
            For y = 1 To Len(letter)
                currentChar = Mid(letter, y, 1)
                KeyPress currentChar, 1
                Delay 100
            Next
            ' 模拟按键数字
            KeyPress CStr(x), 1
            ' 如果不是最后一个数字，模拟按键逗号
            If x <> candidate Then
                KeyPress ",", 1
            End If
        Next
        ' 模拟按键Enter两次，然后延迟200毫秒
        Delay 200
        KeyPress "Enter", 2
        Delay 200
    Next
End Sub

' 主程序
Sub Main()
    ' 根据需要设置组合长度
    combinationLength = 3 ' 或 2 或 3 或 4
    GenerateCombinations combinationLength
    ' 调用处理组合子程序，传入鼠标点击位置
    ProcessCombinations 35, 344
End Sub

' 调用主程序
Call Main