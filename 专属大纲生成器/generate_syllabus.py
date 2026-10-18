# -*- coding: utf-8 -*-
"""
专属大纲生成器（MVP核心能力）
输入: 年级/学期/层次/家长期望/学生漏洞/学生姓名
处理: 加载机构大纲底库 → 按学情裁剪重排 → 标注重点与备注
输出: 专属教学大纲 .xlsx（带格式，可编辑）

用法(命令行):
  python3 generate_syllabus.py --grade 初二 --semester 寒假 --layer 精进 \
      --expectation 查漏 --weak "含参不等式组,一次函数综合" --student 张同学
用法(被调用):
  from generate_syllabus import generate
  generate(grade="初二", semester="寒假", layer="精进", expectation="查漏",
           weak_points=["含参不等式组","一次函数综合"], student_name="张同学")
"""
import json, os, argparse, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

LIB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "大纲底库", "syllabus_library.json")
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# 学期 → 文件 + 学期名
SEM_TO_FILE = {
    "寒假": ("FY26寒春", "寒假"), "春季": ("FY26寒春", "春季"),
    "暑假": ("FY27暑秋", "暑假"), "秋季": ("FY27暑秋", "秋季"),
}
LAYER_DESC = {"好学":"110分以下·基础巩固", "精进":"110-130分·中等综合", "志高":"约140分·拔高压轴", "壮志":"竞赛/冲刺"}

# ---------- 漏洞匹配 ----------
def match_weak(lesson, weak_points):
    """判断该讲次是否与某漏洞相关，返回匹配的漏洞列表"""
    text = (lesson.get("title","") + lesson.get("detail",""))
    hit = []
    for w in weak_points:
        w = w.strip()
        if not w: continue
        # 双向匹配：漏洞关键词在讲次文本里，或讲次标题在漏洞里
        if w in text or any(k in w for k in [lesson.get("title","")] if k):
            hit.append(w)
    return hit

# ---------- 重排逻辑 ----------
def rearrange(lessons, expectation, weak_points, mastered=None):
    """
    expectation: '查漏'(查漏优先) 或 '同步'(同步优先)
    weak_points: 学生漏洞知识点列表
    mastered: 已掌握知识点列表（标记快速过）
    返回: (排好序的讲次列表, 每讲备注说明)
    """
    mastered = mastered or []
    notes = {}  # id(lesson) -> 备注
    annotated = []
    for les in lessons:
        title = les.get("title","")
        detail = les.get("detail","")
        hits = match_weak(les, weak_points)
        is_mastered = any(m.strip() in (title+detail) for m in mastered if m.strip())
        is_review = les.get("is_review", False)
        tag = []
        if hits:
            tag.append("🔴漏洞巩固：" + "/".join(hits))
        if is_mastered:
            tag.append("🟢已掌握·快速过")
        if les.get("starred"):
            tag.append("★重点")
        annotated.append({
            "les": les,
            "hits": hits,
            "is_mastered": is_mastered,
            "is_review": is_review,
            "note": "；".join(tag) if tag else "",
            "weak_hit": bool(hits),
        })

    if expectation == "查漏":
        # 1) 漏洞相关讲次优先（按漏洞出现顺序）
        weak_lessons = [a for a in annotated if a["weak_hit"]]
        # 漏洞讲次按其命中的第一个漏洞在weak_points中的顺序排
        def weak_order(a):
            for i,w in enumerate(weak_points):
                if w.strip() in (a["les"].get("title","")+a["les"].get("detail","")):
                    return i
            return 999
        weak_lessons.sort(key=weak_order)
        # 2) 其余常规讲次（排除已掌握可后置，但保留顺序）
        rest = [a for a in annotated if not a["weak_hit"]]
        # 已掌握的放最后
        rest.sort(key=lambda a: (1 if a["is_mastered"] else 0))
        ordered = weak_lessons + rest
        # 给漏洞讲次补备注
        for a in weak_lessons:
            if not a["note"]:
                a["note"] = "🔴漏洞巩固（前置）"
            else:
                a["note"] = "【查漏前置】" + a["note"]
    else:
        # 同步优先：保持原顺序，漏洞讲次加重点关注
        ordered = annotated
        for a in ordered:
            if a["weak_hit"]:
                a["note"] = ("【重点关注】" if not a["note"] else "【重点关注】" + a["note"])
    return ordered

# ---------- 输出 xlsx ----------
def write_xlsx(ordered, meta, out_path, cross_suggest=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "专属大纲"
    # 样式
    thin = Side(style="thin", color="cccccc")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    title_font = Font(name="微软雅黑", size=14, bold=True, color="ffffff")
    head_font = Font(name="微软雅黑", size=11, bold=True, color="ffffff")
    cell_font = Font(name="微软雅黑", size=10)
    head_fill = PatternFill("solid", fgColor="1e3a8a")
    weak_fill = PatternFill("solid", fgColor="fee2e2")     # 漏洞巩固 浅红
    master_fill = PatternFill("solid", fgColor="dcfce7")   # 已掌握 浅绿
    review_fill = PatternFill("solid", fgColor="fef3c7")   # 复盘/测试 浅黄
    star_font = Font(name="微软雅黑", size=10, bold=True, color="b91c1c")
    wrap = Alignment(wrap_text=True, vertical="center")

    # 标题行
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = f"{meta['student']} 专属教学大纲"
    c.font = title_font
    c.fill = PatternFill("solid", fgColor="2563eb")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 信息行
    ws.merge_cells("A2:H2")
    info = f"年级：{meta['grade']}　|　层次：{meta['layer']}（{LAYER_DESC.get(meta['layer'],'')}）　|　学期：{meta['semester']}　|　家长期望：{meta['expectation']}　|　生成日期：{meta['date']}"
    c = ws["A2"]
    c.value = info
    c.font = Font(name="微软雅黑", size=10, color="475569")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # 表头
    headers = ["讲次","大标题","具体内容/小标题","题型","难度","分值","重点","备注（裁剪说明）"]
    for j,h in enumerate(headers,1):
        c = ws.cell(row=3, column=j, value=h)
        c.font = head_font; c.fill = head_fill; c.border = border; c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 24

    # 数据行
    r = 4
    for i, a in enumerate(ordered, 1):
        les = a["les"]
        vals = [
            str(i),
            les.get("title",""),
            les.get("detail",""),
            les.get("qtype",""),
            les.get("difficulty",""),
            les.get("score",""),
            "★" if les.get("starred") else "",
            a["note"],
        ]
        for j,v in enumerate(vals,1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = star_font if (j==7 and v) else cell_font
            c.border = border; c.alignment = wrap
        # 行底色
        fill = None
        if a["weak_hit"]: fill = weak_fill
        elif a["is_mastered"]: fill = master_fill
        elif a["is_review"]: fill = review_fill
        if fill:
            for j in range(1,9):
                ws.cell(row=r, column=j).fill = fill
        ws.row_dimensions[r].height = 38
        r += 1

    # 列宽
    widths = [6,18,42,16,8,8,6,28]
    for j,w in enumerate(widths,1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    # 图例说明
    r += 1
    ws.cell(row=r, column=1, value="图例：").font = Font(name="微软雅黑", size=9, bold=True)
    legends = [("🔴浅红=漏洞巩固（查漏前置）", weak_fill),
               ("🟢浅绿=已掌握·快速过", master_fill),
               ("🟡浅黄=复盘/测试", review_fill)]
    for k,(txt,fl) in enumerate(legends):
        c = ws.cell(row=r, column=2+k*2, value=txt)
        c.font = Font(name="微软雅黑", size=9); c.fill = fl
    r += 2
    ws.cell(row=r, column=1, value="说明：").font = Font(name="微软雅黑", size=9, bold=True)
    ws.cell(row=r, column=2, value="本大纲由机构底库×学生学情自动裁剪生成，可直接编辑调整；右键行可插入/删除讲次。").font = Font(name="微软雅黑", size=9, color="64748b")

    # 未覆盖漏洞补课建议（跨学期/跨年级查找）
    if cross_suggest:
        r += 2
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(row=r, column=1, value="⚠ 未在本学期大纲中覆盖的漏洞 · 跨学期/跨年级补课建议")
        c.font = Font(name="微软雅黑", size=10, bold=True, color="b91c1c")
        c.fill = PatternFill("solid", fgColor="fef2f2")
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 22
        for sug in cross_suggest:
            r += 1
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            if sug["found"]:
                founds = "；".join([f"{f['src']}「{f['title']}」{'★' if f['starred'] else ''}" for f in sug["found"]])
                txt = f"· 漏洞【{sug['weak']}】→ 建议补：{founds}"
            else:
                txt = f"· 漏洞【{sug['weak']}】→ 底库未找到对应讲次，建议安排专题补讲"
            c = ws.cell(row=r, column=1, value=txt)
            c.font = Font(name="微软雅黑", size=9, color="475569")
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[r].height = 28

    wb.save(out_path)
    return out_path

# ---------- 主函数 ----------
def generate(grade, semester, layer, expectation, weak_points, student_name="学生", mastered=None, out_path=None):
    # 1. 选底库
    if semester not in SEM_TO_FILE:
        raise ValueError(f"学期应为 寒假/春季/暑假/秋季，当前：{semester}")
    file_key, sem_name = SEM_TO_FILE[semester]
    with open(LIB_PATH, encoding="utf-8") as f:
        lib = json.load(f)
    # 2. 加载讲次
    if grade not in lib.get(file_key, {}):
        raise ValueError(f"{file_key} 暂无 {grade} 大纲底库")
    grade_data = lib[file_key][grade]
    if layer not in grade_data:
        # 志高在FY26缺失等场景：降级提示
        avail = "、".join(grade_data.keys())
        raise ValueError(f"{grade} 在 {file_key}({semester}) 暂无 [{layer}] 层次底库，现有：{avail}。志高生建议选「暑假/秋季」学期（FY27暑秋含志高）。")
    sem_data = grade_data[layer]
    if sem_name not in sem_data:
        raise ValueError(f"{grade}/{layer} 在 {semester} 无数据，请检查学期与层次搭配。")
    base_lessons = sem_data[sem_name]

    # 3. 重排+标注
    ordered = rearrange(base_lessons, expectation, weak_points or [], mastered or [])

    # 4. 输出
    os.makedirs(OUT_DIR, exist_ok=True)
    if not out_path:
        safe = student_name.replace(" ","")
        out_path = os.path.join(OUT_DIR, f"{safe}_{grade}{layer}_{semester}专属大纲.xlsx")
    meta = {
        "student": student_name, "grade": grade, "layer": layer,
        "semester": semester, "expectation": expectation,
        "date": datetime.date.today().strftime("%Y-%m-%d"),
    }

    # 4.5 统计未匹配漏洞 + 跨学期/跨年级查找补课建议
    matched_weak = set()
    for a in ordered:
        for h in a["hits"]:
            matched_weak.add(h)
    unmatched_weak = [w for w in weak_points if w not in matched_weak]
    cross_suggest = []
    if unmatched_weak:
        GRADE_ORDER = ["初一","初二","初三"]
        for w in unmatched_weak:
            found = []
            for fk, grades in lib.items():
                for g, layers in grades.items():
                    if g == grade: continue  # 跨年级查找
                    for ly, sems in layers.items():
                        if not isinstance(sems, dict) or "error" in sems: continue
                        for sm, ls in sems.items():
                            for les in ls:
                                if w in (les.get("title","")+les.get("detail","")):
                                    found.append({"src":f"{fk}·{g}/{ly}/{sm}", "title":les.get("title",""), "starred":les.get("starred",False)})
            cross_suggest.append({"weak":w, "found":found[:4]})

    write_xlsx(ordered, meta, out_path, cross_suggest)

    # 5. 生成说明
    weak_n = sum(1 for a in ordered if a["weak_hit"])
    master_n = sum(1 for a in ordered if a["is_mastered"])
    review_n = sum(1 for a in ordered if a["is_review"])
    summary = {
        "out_path": out_path,
        "total_lessons": len(ordered),
        "weak_lessons": weak_n,
        "mastered_lessons": master_n,
        "review_lessons": review_n,
        "base_source": f"{file_key}/{grade}/{layer}/{sem_name}（{len(base_lessons)}讲）",
        "expectation": expectation,
        "weak_points": weak_points,
        "unmatched_weak": unmatched_weak,
        "cross_suggest": cross_suggest,
    }
    return summary

# ---------- 命令行 ----------
if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="生成学生专属教学大纲")
    ap.add_argument("--grade", required=True, choices=["初一","初二","初三"])
    ap.add_argument("--semester", required=True, choices=["寒假","春季","暑假","秋季"])
    ap.add_argument("--layer", required=True, choices=["好学","精进","志高","壮志"])
    ap.add_argument("--expectation", required=True, choices=["查漏","同步"])
    ap.add_argument("--weak", default="", help="学生漏洞知识点，逗号分隔")
    ap.add_argument("--mastered", default="", help="已掌握知识点，逗号分隔")
    ap.add_argument("--student", default="学生")
    args = ap.parse_args()
    weak = [w.strip() for w in args.weak.split(",") if w.strip()]
    mastered = [m.strip() for m in args.mastered.split(",") if m.strip()]
    res = generate(args.grade, args.semester, args.layer, args.expectation, weak, args.student, mastered)
    print("=== 专属大纲生成完成 ===")
    for k,v in res.items():
        print(f"  {k}: {v}")
