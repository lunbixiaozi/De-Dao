# workbook

* 计算workbook的数量

    Application.Workbooks.Count

* 找到active workbook的名字

    ActiveWorkbook.Name


# worksheet

* 找到当前workbook的第二个worksheet

    ActiveWorkbook.Sheets(2).Name

# 数据类型

* string里查看char的数字

    Len(myWord)

* string里取出一定数量的char

    Left(myWord, 4)

# 单元格

* 读出某cell的值

    FirstQ_C5 = ActiveWorkbook.Sheets(1).Range("C5").Value

* 读出某cell的公式

    Range("A1").Formula

* 改变字号

    Range("A1").Font.Size = 14

* 字体变粗

    Range("A1").Font.Bold = True

* 字体变斜体，改变对齐方式：左右对齐，水平对齐

    Range("A1").Font.italic = True
    Range(.Offset(0, 0), .End(xlToRight)).HorizontalAlignment = xlCenter

* 改变字的颜色

    Range(.Offset(0, 1), .End(xlDown).End(xlToRight)).Interior.Color = RGB(240, 240, 240)
    或者：Range("A1").Font.Color = vbRed

* 背景变黑

# 函数

* 函数的return值
* 用vba计算min，max，均值，方差，median

# message box

* message box的显示改一下，把变量传进去
* break a line: vbNewLine
* 修改按钮

    MsgBox "suibian", vbYesNo

# button

* assign一个宏
