'
'目前包括：地理数据相关；账本处理；学号统计
'

prov_lst = Array("北京市", "浙江省", "天津市", "安徽省", "上海市", "福建省", "重庆市", "江西省", "香港特别行政区", "山东省", "澳门特别行政区", "河南省", "内蒙古自治区", "湖北省", "新疆维吾尔自治区", "湖南省", "宁夏回族自治区", "广东省", "西藏自治区", "海南省", "广西壮族自治区", "四川省", "河北省", "贵州省", "山西省", "云南省", "辽宁省", "陕西省", "吉林省", "甘肃省", "黑龙江省", "青海省", "江苏省", "台湾省")
pro_tile = Array("北京", "浙江", "天津", "安徽", "上海", "福建", "重庆", "江西", "香港", "山东", "澳门", "河南", "内蒙古", "湖北", "新疆", "湖南", "宁夏", "广东", "西藏", "海南", "广西", "四川", "河北", "贵州", "山西", "云南", "辽宁", "陕西", "吉林", "甘肃", "黑龙", "青海", "江苏", "台湾")

Sub sumWeek() '账本处理
Dim su(0 To 7, 0 To 1) As Integer
i = 2
Do While (i < 800)
    ak = Cells(i, "D")
    If ak = "饭卡" Then
    'If ak = "早午晚餐" Or ak = "饮料" Or ak = "水果零食" Then
        
        s = Cells(i, "K")
        If (s = "星期一") Then
            su(0, 0) = su(0, 0) + Cells(i, "G")
            su(0, 1) = su(0, 1) + 1
        ElseIf (s = "星期二") Then
            su(1, 0) = su(1, 0) + Cells(i, "G")
            su(1, 1) = su(1, 1) + 1
        ElseIf (s = "星期三") Then
            su(2, 0) = su(2, 0) + Cells(i, "G")
            su(2, 1) = su(2, 1) + 1
        ElseIf (s = "星期四") Then
            su(3, 0) = su(3, 0) + Cells(i, "G")
            su(3, 1) = su(3, 1) + 1
        ElseIf (s = "星期五") Then
            su(4, 0) = su(4, 0) + Cells(i, "G")
            su(4, 1) = su(4, 1) + 1
        ElseIf (s = "星期六") Then
            su(5, 0) = su(5, 0) + Cells(i, "G")
            su(5, 1) = su(5, 1) + 1
        ElseIf (s = "星期日") Then
            su(6, 0) = su(6, 0) + Cells(i, "G")
            su(6, 1) = su(6, 1) + 1
        
        End If
    End If
    
    i = i + 1
Loop

Worksheets(2).Activate
Range("K4").Resize(8, 2) = su

End Sub

Sub sumDoom()
Dim su(0 To 7, 0 To 1) As Double
i = 2
Do While (i < 800)
    ak = Cells(i, "B")
    If ak = "食品酒水" Then
    'If ak = "早午晚餐" Or ak = "饮料" Or ak = "水果零食" Then
        
        s = Cells(i, "F")
        If (s = "学一") Then
            su(0, 0) = su(0, 0) + Cells(i, "G")
            su(0, 1) = su(0, 1) + 1
        ElseIf (s = "燕南美食") Then
            su(1, 0) = su(1, 0) + Cells(i, "G")
            su(1, 1) = su(1, 1) + 1
        ElseIf (s = "学五") Then
            su(2, 0) = su(2, 0) + Cells(i, "G")
            su(2, 1) = su(2, 1) + 1
        ElseIf (s = "松林") Then
            su(3, 0) = su(3, 0) + Cells(i, "G")
            su(3, 1) = su(3, 1) + 1
        ElseIf (s = "农园") Then
            su(4, 0) = su(4, 0) + Cells(i, "G")
            su(4, 1) = su(4, 1) + 1
        ElseIf (s = "勺园") Then
            su(5, 0) = su(5, 0) + Cells(i, "G")
            su(5, 1) = su(5, 1) + 1
        Else
            su(6, 0) = su(6, 0) + Cells(i, "G")
            su(6, 1) = su(6, 1) + 1
        
        End If
    End If
    
    i = i + 1
Loop

Worksheets(2).Activate
Range("D70").Resize(8, 2) = su
Debug.Print ("12")
End Sub

'----------------
Sub getXuehaodata() 'pku学号数据统计
Worksheets.Item(1).Activate
schFullArr = Array("城市与环境学院", "地球与空间科学学院", "对外汉语教育学院", "法学院", "分子医学研究所", "歌剧研究院", "工学院", "光华管理学院", "深圳研究生院", "国际关系学院", "国家发展研究院", "化学与分子工程学院", "环境科学与工程学院", "建筑与景观设计学院", "教育学院", "经济学院", "考古文博学院", "历史学系", "马克思主义学院", "前沿交叉学科研究院", "人口研究所", "软件与微电子学院", "社会学系", "生命科学学院", "数学科学学院", "体育教研部", "外国语学院", "物理学院", "心理学系", "新闻与传播学院", "信息管理系", "信息科学技术学院", "艺术学院", "元培学院", "哲学系", "政府管理学院", "中国语言文学系", "燕京学堂", "新媒体研究院") '1-39：也就是0-38
Dim al(0 To 40, 0 To 20) As Variant
 
For q = 0 To 40
    al(q, 0) = 2
    If q < 39 Then
        al(q, 1) = schFullArr(q)
    End If
Next q
al(39, 1) = ""
al(40, 1) = "其他院"

iRows = ActiveSheet.UsedRange.Rows.Count
'行的最大长度

For k = 2 To iRows
    i = 0
    schname = Cells(k, "B")
    For v = 0 To 40:
        
        If al(v, 1) = schname Then  '=其他 的情况之后再说
            xueh = Cells(k, "A")
            xueha = Right(xueh, 5)
            xuehao = Left(xueha, 3) 'xuehao
            indx = al(v, 0)
            If indx = 2 Then
                al(v, indx) = xuehao
                al(v, 0) = indx + 1  '3
            Else
                For yc = 2 To indx
                    If yc = 3 Then
                        ask = 0
                    End If
                    If al(v, yc) = xuehao Then '已经存在
                        Exit For
                    ElseIf yc = indx And al(v, yc) <> xuehao Then '到末位仍然不存在
                        If v = 39 And indx = 20 Then
                        ElseIf indx = 20 Then 'indx=20了
                        
                        Else
                            al(v, 0) = indx + 1
                            al(v, yc) = xuehao
                        End If
                    ElseIf yc <> indx And al(v, yc) <> xuehao Then '非末位，未存在
                        'pass
                    End If
                Next yc
            End If
        End If
    Next v
Next k
'写入Excel表格
Dim tempSheet As Worksheet
Set tempSheet = Worksheets.Add(after:=Worksheets(1))
tempSheet.Range("A1").Resize(40, 20) = al

Application.ScreenUpdating = True

End Sub