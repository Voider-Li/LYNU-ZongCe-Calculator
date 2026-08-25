"""读取输入的 xlsx 文件，解析原始成绩、参评科目、量化数据"""
import re
import openpyxl


def _parse_course_header(cell_value):
    """从 '[00001022]食品安全与日常饮食' 提取 (编号, 科目名)"""
    if not cell_value:
        return None, None
    s = str(cell_value).strip()
    m = re.match(r'\[([^\]]+)\](.+)', s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return None, s


def _parse_category_credit(cell_value):
    """从 '[任选1.0]' 提取 (类别, 学分)"""
    if not cell_value:
        return None, None
    s = str(cell_value).strip()
    m = re.match(r'\[([^\d]+)([\d.]+)\]', s)
    if m:
        return m.group(1).strip(), float(m.group(2))
    return None, None


# 中文月份数字映射（用于量化表扣分列表头识别）
_CN_MONTH = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
             '七': 7, '八': 8, '九': 9, '十': 10, '十一': 11, '十二': 12}


def _parse_month_header(val):
    """识别月份表头，返回月份数字(1-12)；非月份返回 None。
    支持 '9月'/'10月'/'1月' 与 '一月'~'十二月'/'三月' 等中文写法。"""
    if not val:
        return None
    s = str(val).strip()
    m = re.match(r'(\d+)月', s)
    if m:
        n = int(m.group(1))
        return n if 1 <= n <= 12 else None
    m = re.match(r'(十[一二]|[一二三四五六七八九十])月', s)
    if m:
        return _CN_MONTH.get(m.group(1))
    return None


def _school_year_order(month_num):
    """学年排序：9月=1, 10月=2, 11月=3, 12月=4, 1月=5, ..., 8月=12。
    跨年学期中 1月排在 12月之后。"""
    if month_num >= 9:
        return month_num - 8   # 9→1 ... 12→4
    return month_num + 4        # 1→5 ... 8→12


# 表头标签黑名单（用于区分"科目行"与"序号/学号/姓名"等标签行）
_LABEL_BLACKLIST = {'序号', '学号', '姓名', '总成绩', '排名', '成绩', '平均成绩',
                    '学分绩点', '平均学分绩点', '课程/环节', '类别', '学分'}


def _detect_header_rows(ws, name_col=4):
    """自动检测表头行位置，兼容两种格式：
    ① 有标题行（如大二上：标题在 A1，科目在第2行，学分第3行，学生第4行起）
    ② 直接从科目行开始（如使用说明所述：无标题，科目在第1行，学分第2行，学生第3行起）
    返回 (科目名行, 学分行, 学生起始行)。"""
    name_row = None
    for r in range(1, 7):
        v = ws.cell(r, name_col).value
        if v is None:
            continue
        s = str(v).strip()
        if s == '':
            continue
        # 纯数字（学分/成绩）跳过
        if re.fullmatch(r'-?[\d.]+', s):
            continue
        if s in _LABEL_BLACKLIST:
            continue
        name_row = r   # 第一个在 name_col 出现的非数字、非标签文本行 → 科目行
        break
    if name_row is None:
        name_row = 2   # 兜底
    # 学分行 = 科目行下方第一个非空行
    credits_row = name_row + 1
    for r in range(name_row + 1, min(name_row + 4, ws.max_row + 1)):
        v = ws.cell(r, name_col).value
        if v is not None and str(v).strip() != '':
            credits_row = r
            break
    return name_row, credits_row, credits_row + 1


def read_raw_scores(path):
    """
    读取原始成绩表 sheet。
    返回:
      title: str  表标题
      courses: [{code, name, category, credit, col_idx}]
      students: [{seq, id, name, scores: {col_idx: score}}]
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    # 找原始表 sheet（名字含"原始"）
    raw_sheet = None
    for name in wb.sheetnames:
        if '原始' in name:
            raw_sheet = wb[name]
            break
    if raw_sheet is None:
        # 退化：用第一个 sheet
        raw_sheet = wb[wb.sheetnames[0]]

    ws = raw_sheet
    # 自动检测表头行（兼容有标题 / 从科目行开始 两种格式）
    header_row, credits_row, student_start = _detect_header_rows(ws)

    # 表标题：A1 若是普通文本则用，否则留空
    raw_title = str(ws.cell(1, 1).value or '').strip()
    if raw_title in _LABEL_BLACKLIST or re.match(r'\[', raw_title):
        raw_title = ''
    title = raw_title

    # 课程列从第4列开始；学生信息在前几列（序号、学号、姓名）
    id_col = 2
    name_col = 3
    courses = []
    for col in range(4, ws.max_column + 1):
        val = ws.cell(header_row, col).value
        code, cname = _parse_course_header(val)
        if cname:
            cat_cred = ws.cell(credits_row, col).value
            cat, cred = _parse_category_credit(cat_cred)
            # 末尾汇总列跳过
            if cname in ('总成绩', '平均成绩', '平均\n成绩', '学分\n绩点', '平均\n学分\n绩点', '学分绩点', '平均学分绩点'):
                continue
            courses.append({
                'col_idx': col,
                'code': code,
                'name': cname,
                'category': cat,
                'credit': cred
            })

    # 读学生数据
    students = []
    for r in range(student_start, ws.max_row + 1):
        sid = ws.cell(r, id_col).value
        sname = ws.cell(r, name_col).value
        if not sid:
            continue
        scores = {}
        for c in courses:
            v = ws.cell(r, c['col_idx']).value
            if v is not None and v != '':
                try:
                    scores[c['col_idx']] = float(v)
                except (ValueError, TypeError):
                    pass
        students.append({
            'seq': ws.cell(r, 1).value,
            'id': str(sid),
            'name': str(sname) if sname else '',
            'scores': scores
        })

    wb.close()
    return {
        'title': title,
        'courses': courses,
        'students': students,
    }


def read_eval_subjects(path):
    """
    读取参评科目 sheet。
    返回: [{name, credit}] 或 None（如果不存在）
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = None
    for name in wb.sheetnames:
        if '参评' in name:
            sheet = wb[name]
            break
    if sheet is None:
        wb.close()
        return None

    ws = sheet
    # 自动检测表头行（兼容有标题 / 从科目行开始 两种格式）
    subj_row, credit_row, student_start = _detect_header_rows(ws)
    # 说明：部分示例文件的「参评科目」表会在右侧并列第二张表
    # （如「成绩排序」表，行序与左侧不同）。为避免把第二张表的同名列
    # 误并入并覆盖左侧成绩，遇到空列即视为分表间隙，停止继续读取。
    subjects = []
    started = False
    for col in range(4, ws.max_column + 1):
        name_val = ws.cell(subj_row, col).value
        if not name_val or str(name_val).strip() == '':
            if started:
                break  # 第一张表结束（遇到间隙）
            continue
        cname = str(name_val).strip()
        if cname in ('总成绩', '总成绩\n', '总成绩\n'):
            continue
        started = True
        cred_val = ws.cell(credit_row, col).value
        try:
            cred = float(cred_val) if cred_val is not None else None
        except (ValueError, TypeError):
            cred = None
        subjects.append({'name': cname, 'credit': cred, 'col_idx': col})

    # 同时读取学生成绩
    students = []
    for r in range(student_start, ws.max_row + 1):
        sid = ws.cell(r, 2).value
        sname = ws.cell(r, 3).value
        if not sid:
            continue
        scores = {}
        for s in subjects:
            v = ws.cell(r, s['col_idx']).value
            if v is not None and v != '':
                try:
                    scores[s['name']] = float(v)
                except (ValueError, TypeError):
                    pass
        students.append({
            'id': str(sid),
            'name': str(sname) if sname else '',
            'scores': scores
        })

    wb.close()
    return {'subjects': subjects, 'students': students}


def read_quantization(path):
    """
    读取量化 sheet。
    返回: {students: [{id, name, second_class_raw, deductions: {月: val}}], months: [月份列表]}
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet = None
    for name in wb.sheetnames:
        if '量化' in name:
            sheet = wb[name]
            break
    if sheet is None:
        wb.close()
        return None

    ws = sheet
    # 量化表结构（模板）：
    # R3: 序号|学号|姓名|二课量化(本学期)...|德育量化(基础分|扣分9月~1月)|量化最终值
    # 二课量化原始值在第4列(D)
    # 德育基础分在第7列(G)，扣分在8-12列(H-L)，对应9月~1月
    # 需要动态扫描表头

    # 找二课量化列：表头含"二课"的
    second_class_col = None
    base_score_col = None
    base_mgmt_col = None   # "总分"/"基础管理量化"列（源表已算好的扣分后值）
    month_cols = {}   # {col: (month_num, label)} —— 任意月份、任意列数
    final_col = None

    # 扫描前6行找表头（不同模板的月份行位置不同：大二上在第5行、大一下在第4行、老格式在第2行）
    for r in range(1, 7):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or '').strip()
            if not val:
                continue
            if '二课' in val and second_class_col is None:
                second_class_col = c
            if ('基础' in val or '基礎' in val) and '量化最终' not in val and '最终值' not in val:
                if base_score_col is None:
                    base_score_col = c
            # "总分"或"基础管理量化"列：源表里 Excel 已算好的扣分后值
            if ('总分' in val or '基础管理' in val or '基础管理量化' in val) \
                    and '量化最终' not in val and '最终值' not in val:
                if base_mgmt_col is None:
                    base_mgmt_col = c
            mn = _parse_month_header(val)
            if mn is not None:
                month_cols[c] = (mn, val)   # 同列出现在多行时取最下层
            elif ('量化最终' in val or '最终值' in val) and final_col is None:
                final_col = c

    # 按学年顺序排序（9月起），月份列数不限
    ordered = sorted(month_cols.items(), key=lambda kv: _school_year_order(kv[1][0]))
    months = [label for _, (_, label) in ordered]
    deduction_cols = {label: col for col, (_, label) in ordered}

    # ---- 新模板：无月份列，改为单个"扣分/德育量化"合并区（如 E:I）----
    # 扣分区只在无月份时启用（避免与旧模板的"扣分"段标题冲突）
    deduct_col = None       # 扣分区起始列
    deduct_end_col = None   # 扣分区结束列（含）
    if not month_cols:
        for r in range(1, 7):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if not val:
                    continue
                if ('扣分' in val or '德育' in val) \
                        and '量化最终' not in val and '最终值' not in val and '二课' not in val:
                    if deduct_col is None:
                        deduct_col = c
                        # 由合并范围确定列跨度（如 E2:I2 → E..I）
                        for mr in ws.merged_cells.ranges:
                            if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                                deduct_end_col = mr.max_col
                                break
                        if deduct_end_col is None:
                            deduct_end_col = c  # 未合并，单列
        if deduct_col is not None:
            months = ['扣分']
            deduction_cols = {'扣分': deduct_col}   # 标记用，实际读取用区间

    students = []
    # 学生数据从第6行开始（模板），但动态找
    start_row = 6
    for r in range(1, ws.max_row + 1):
        sid = ws.cell(r, 2).value
        if sid and str(sid).strip().isdigit() or (sid and len(str(sid)) >= 6):
            # 确认是学号行
            sname = ws.cell(r, 3).value
            sc_raw = ws.cell(r, second_class_col).value if second_class_col else None
            try:
                sc_raw = float(sc_raw) if sc_raw else 0
            except (ValueError, TypeError):
                sc_raw = 0

            deductions = {}
            if deduct_col is not None and not month_cols:
                # 新模板：扣分区多格按日期填 -2，对整个区间求和作为总扣分
                total = 0.0
                for c in range(deduct_col, deduct_end_col + 1):
                    v = ws.cell(r, c).value
                    if v is None or v == '':
                        continue
                    try:
                        total += float(v)
                    except (ValueError, TypeError):
                        pass
                deductions['扣分'] = total
            else:
                for m, col in deduction_cols.items():
                    v = ws.cell(r, col).value
                    try:
                        deductions[m] = float(v) if v is not None and v != '' else 0
                    except (ValueError, TypeError):
                        deductions[m] = 0

            # 读源表"总分/基础管理量化"列（Excel 已算好的扣分后值）
            base_mgmt_src = None
            if base_mgmt_col:
                v = ws.cell(r, base_mgmt_col).value
                if v is not None and v != '':
                    try:
                        base_mgmt_src = float(v)
                    except (ValueError, TypeError):
                        base_mgmt_src = None

            students.append({
                'id': str(sid),
                'name': str(sname) if sname else '',
                'second_class_raw': sc_raw,
                'deductions': deductions,
                'base_mgmt_source': base_mgmt_src,
            })
            if r >= start_row:
                pass  # 已记录

    # 去重：只保留真正学生行（学号长度>=6）
    students = [s for s in students if len(s['id']) >= 6]

    wb.close()
    return {
        'students': students,
        'months': months,
        'second_class_col': second_class_col,
        'deduction_cols': deduction_cols,
        'base_mgmt_col': base_mgmt_col,
    }
