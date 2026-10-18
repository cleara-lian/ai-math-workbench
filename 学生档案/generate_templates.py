# -*- coding: utf-8 -*-
"""生成学生档案模板 + 学情诊断报告模板（带格式，含示例数据）"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OUT_DIR = "/workspace/AI教学工作台/学生档案"
thin = Side(style="thin", color="cccccc")
border = Border(left=thin,right=thin,top=thin,bottom=thin)
HF = Font(name="微软雅黑", size=11, bold=True, color="ffffff")
CF = Font(name="微软雅黑", size=10)
TF = Font(name="微软雅黑", size=14, bold=True, color="ffffff")
HFill = PatternFill("solid", fgColor="1e3a8a")
TitleFill = PatternFill("solid", fgColor="2563eb")
SecFill = PatternFill("solid", fgColor="eff6ff")
wrap = Alignment(wrap_text=True, vertical="center")
ctr = Alignment(horizontal="center", vertical="center")

def style_header(ws, row, ncol):
    for j in range(1, ncol+1):
        c = ws.cell(row=row, column=j)
        c.font = HF; c.fill = HFill; c.border = border; c.alignment = ctr

def put_title(ws, text, ncol, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row=row, column=1, value=text)
    c.font = TF; c.fill = TitleFill; c.alignment = ctr
    ws.row_dimensions[row].height = 30

def set_widths(ws, widths):
    for j,w in enumerate(widths,1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

# ============ 1. 学生档案模板 ============
wb = openpyxl.Workbook()

# Sheet1 档案总览
ws = wb.active; ws.title = "档案总览"
put_title(ws, "学生教学档案（示例：张同学）", 6)
ws["A2"] = "一、基本信息"; ws["A2"].font = Font(name="微软雅黑", size=11, bold=True, color="1e3a8a")
info = [
    ("姓名","张同学"),("年级","初二"),("学校","（填写）"),("班型","精进"),("层次","110-130分"),
    ("家长期望","查漏优先，冲重点高中"),("目标分数","115+"),("薄弱模块","含参不等式组、一次函数综合"),
    ("学习习惯","审题粗心、几何证明思路易卡壳"),("当前进度","反比例函数二（已上12讲）"),
]
r = 3
for i in range(0, len(info), 2):
    ws.cell(r,1,info[i][0]).font = CF; ws.cell(r,1).fill = SecFill; ws.cell(r,1).border=border
    ws.cell(r,2,info[i][1]).font = CF; ws.cell(r,2).border=border
    if i+1 < len(info):
        ws.cell(r,4,info[i+1][0]).font = CF; ws.cell(r,4).fill = SecFill; ws.cell(r,4).border=border
        ws.cell(r,5,info[i+1][1]).font = CF; ws.cell(r,5).border=border
    r += 1
r += 1
ws.cell(r,1,"二、学生画像摘要（AI自动生成）").font = Font(name="微软雅黑", size=11, bold=True, color="1e3a8a")
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=6)
c = ws.cell(r,1,"初二精进班，目标冲刺重点高中。致命漏洞：含参不等式组、一次函数综合；习惯：几何证明思路易卡壳、审题粗心。已上12讲，当前进度：反比例函数二。家长配合度高，期望本学期提至115+。")
c.font = CF; c.alignment = wrap; c.fill = SecFill
for rr in range(r, r+3):
    for j in range(1,7): ws.cell(rr,j).border = border
set_widths(ws, [14,22,3,14,22,3])

# Sheet2 学情诊断记录
ws2 = wb.create_sheet("学情诊断记录")
put_title(ws2, "学情诊断记录（历次）", 8)
heads = ["日期","试卷来源","总分","致命漏洞","薄弱模块","提分方向","建议切入讲次","备注"]
for j,h in enumerate(heads,1): ws2.cell(2,j,h)
style_header(ws2, 2, 8)
sample = [
    ["2026-09-15","开学考","98","含参不等式组(-6)、一次函数综合(-8)","分式方程应用","先补含参不等式边界，再衔接一次函数","FY26初一好学寒假·含参不等式组★","几何证明思路需强化"],
    ["2026-10-18","月考","102","一次函数综合(-6)","反比例函数k几何意义","巩固一次函数后推进反比例","FY26初二精进寒假·一次函数综合","较上次+4分"],
]
for i,row in enumerate(sample,3):
    for j,v in enumerate(row,1):
        c=ws2.cell(i,j,v); c.font=CF; c.border=border; c.alignment=wrap
    ws2.row_dimensions[i].height = 38
set_widths(ws2,[12,12,7,24,16,24,22,16])

# Sheet3 成绩记录
ws3 = wb.create_sheet("成绩记录")
put_title(ws3, "成绩记录（历次考试）", 9)
heads = ["日期","考试名称","总分","班级排名","代数","几何","函数","统计概率","备注"]
for j,h in enumerate(heads,1): ws3.cell(2,j,h)
style_header(ws3, 2, 9)
for i,row in enumerate([["2026-09-15","开学考","98","15","34","22","30","12","函数失分多"],["2026-10-18","月考","102","12","36","24","32","10","较上次+4"]],3):
    for j,v in enumerate(row,1):
        c=ws3.cell(i,j,v); c.font=CF; c.border=border; c.alignment=ctr
set_widths(ws3,[12,12,7,9,7,7,7,9,16])

# Sheet4 课堂表现
ws4 = wb.create_sheet("课堂表现")
put_title(ws4, "课堂表现速记（每节课）", 7)
heads = ["日期","课次","本节内容","亮点","问题","反馈已发","备注"]
for j,h in enumerate(heads,1): ws4.cell(2,j,h)
style_header(ws4, 2, 7)
for i,row in enumerate([["2026-10-16","第11讲","反比例函数二","k几何意义理解到位","与一次函数综合时交点算错","是","下次补交点训练"]],3):
    for j,v in enumerate(row,1):
        c=ws4.cell(i,j,v); c.font=CF; c.border=border; c.alignment=wrap
    ws4.row_dimensions[i].height = 34
set_widths(ws4,[12,9,16,18,18,9,16])

# Sheet5 家长沟通
ws5 = wb.create_sheet("家长沟通")
put_title(ws5, "家长沟通记录", 5)
heads = ["日期","场景","沟通要点","话术要点/回复","备注"]
for j,h in enumerate(heads,1): ws5.cell(2,j,h)
style_header(ws5, 2, 5)
for i,row in enumerate([
    ["2026-10-16","课后反馈","反比例函数二课堂情况","套用场景A，问题具体到交点计算",""],
    ["2026-10-19","提问应答","家长问为什么没提分","套用B1，摆数据98→102+4，解释函数综合滞后期","家长认可"],
],3):
    for j,v in enumerate(row,1):
        c=ws5.cell(i,j,v); c.font=CF; c.border=border; c.alignment=wrap
    ws5.row_dimensions[i].height = 34
set_widths(ws5,[12,12,20,28,16])

# Sheet6 阶段目标
ws6 = wb.create_sheet("阶段目标")
put_title(ws6, "阶段目标与达成", 5)
heads = ["阶段","起止","目标","达成情况","备注"]
for j,h in enumerate(heads,1): ws6.cell(2,j,h)
style_header(ws6, 2, 5)
for i,row in enumerate([["寒假阶段","2026.01-02","补含参不等式+一次函数综合，开学考105+","进行中","已补含参不等式，一次函数综合强化中"]],3):
    for j,v in enumerate(row,1):
        c=ws6.cell(i,j,v); c.font=CF; c.border=border; c.alignment=wrap
    ws6.row_dimensions[i].height = 34
set_widths(ws6,[14,16,28,14,18])

wb.save(f"{OUT_DIR}/学生档案模板.xlsx")

# ============ 2. 学情诊断报告模板 ============
wb2 = openpyxl.Workbook()
ws = wb2.active; ws.title = "学情诊断报告"
put_title(ws, "学情诊断报告", 6, 1)
# 学生信息
ws["A2"]="学生信息"; ws["A2"].font=Font(name="微软雅黑",size=11,bold=True,color="1e3a8a")
info2 = [("姓名","（填写）"),("年级","（填写）"),("层次","好学/精进/志高"),("诊断日期","（填写）")]
r=3
for i in range(0,len(info2),2):
    ws.cell(r,1,info2[i][0]).fill=SecFill; ws.cell(r,1).font=CF; ws.cell(r,1).border=border
    ws.cell(r,2,info2[i][1]).font=CF; ws.cell(r,2).border=border
    if i+1<len(info2):
        ws.cell(r,4,info2[i+1][0]).fill=SecFill; ws.cell(r,4).font=CF; ws.cell(r,4).border=border
        ws.cell(r,5,info2[i+1][1]).font=CF; ws.cell(r,5).border=border
    r+=1
# 试卷信息
r+=1; ws.cell(r,1,"试卷信息").font=Font(name="微软雅黑",size=11,bold=True,color="1e3a8a"); r+=1
ws.cell(r,1,"试卷来源").fill=SecFill; ws.cell(r,1).font=CF; ws.cell(r,1).border=border
ws.cell(r,2,"（填写）").font=CF; ws.cell(r,2).border=border
ws.cell(r,3,"总分").fill=SecFill; ws.cell(r,3).font=CF; ws.cell(r,3).border=border
ws.cell(r,4,"（填写）").font=CF; ws.cell(r,4).border=border
ws.cell(r,5,"本次得分").fill=SecFill; ws.cell(r,5).font=CF; ws.cell(r,5).border=border
ws.cell(r,6,"（填写）").font=CF; ws.cell(r,6).border=border
r+=2
# 知识漏洞清单
ws.cell(r,1,"知识漏洞清单（按严重度排序）").font=Font(name="微软雅黑",size=11,bold=True,color="1e3a8a"); r+=1
heads=["漏洞知识点","严重度","失分","对应大纲讲次","所属模块","备注"]
for j,h in enumerate(heads,1): ws.cell(r,j,h)
style_header(ws, r, 6); r+=1
for row in [["（如含参不等式组）","致命","-6","FY26初一好学寒假·含参不等式组★","代数","边界问题"],["（如一次函数综合）","薄弱","-8","FY26初二精进寒假·一次函数综合","函数","分类讨论不全"]]:
    for j,v in enumerate(row,1):
        c=ws.cell(r,j,v); c.font=CF; c.border=border; c.alignment=wrap
    ws.row_dimensions[r].height=30; r+=1
r+=1
# 失分模块统计
ws.cell(r,1,"失分模块统计（AI自动汇总）").font=Font(name="微软雅黑",size=11,bold=True,color="1e3a8a"); r+=1
heads=["模块","代数","几何","函数","统计概率","合计"]
for j,h in enumerate(heads,1): ws.cell(r,j,h)
style_header(ws, r, 6); r+=1
for j,v in enumerate(["失分","-6","-4","-8","0","-18"],1):
    c=ws.cell(r,j,v); c.font=CF; c.border=border; c.alignment=ctr
r+=2
# 提分方向
ws.cell(r,1,"提分方向建议").font=Font(name="微软雅黑",size=11,bold=True,color="1e3a8a"); r+=1
ws.merge_cells(start_row=r,start_column=1,end_row=r+2,end_column=6)
c=ws.cell(r,1,"（AI生成）建议先用2讲补含参不等式组的边界问题（致命漏洞），再衔接一次函数综合（薄弱点）；几何证明思路需配合模型识别训练。预计补漏后提分空间8-12分。")
c.font=CF; c.alignment=wrap; c.fill=SecFill
for rr in range(r,r+3):
    for j in range(1,7): ws.cell(rr,j).border=border
r+=4
ws.cell(r,1,"⚠ 本报告生成后将自动写入该学生档案的「学情诊断记录」").font=Font(name="微软雅黑",size=9,italic=True,color="dc2626")
set_widths(ws,[18,10,7,28,12,16])
wb2.save(f"{OUT_DIR}/学情诊断报告模板.xlsx")

print("已生成:")
print(f"  {OUT_DIR}/学生档案模板.xlsx")
print(f"  {OUT_DIR}/学情诊断报告模板.xlsx")
