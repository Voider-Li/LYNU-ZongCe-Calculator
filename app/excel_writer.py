"""写出综测结果 xlsx（多 sheet，写 Excel 公式）

结构：
  1. 保留上传文件里的原 sheet（原始成绩 / 参评科目 / 量化等）在最前；
  2. 程序计算的 sheet 追加在后面：量化处理、成绩学分绩点处理、排名百分比；
  3. 最后一步是三个可切换的小表格：成绩处理、量化处理、综测成绩（含 不含量化/含量化）。

排版：每个 sheet 第一行合并标题，表头暖橙渐变 + 白字，数据居中、
微软雅黑 11、细边框白色填充。
"""
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ---------- 通用样式 ----------
FONT_FAMILY = 'Microsoft YaHei'

TITLE_FONT = Font(name=FONT_FAMILY, size=16, bold=True, color='1F4E79')
TITLE_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
TITLE_FILL = PatternFill(fill_type='solid', start_color='D9E1F2', end_color='D9E1F2')

HEADER_FONT = Font(name=FONT_FAMILY, size=11, bold=True, color='FFFFFFFF')
HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
HEADER_FILL = PatternFill(fill_type='solid', start_color='4472C4', end_color='4472C4')

DATA_FONT = Font(name=FONT_FAMILY, size=11, bold=False)
DATA_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
DATA_FILL = PatternFill(fill_type='solid', start_color='FFFFFFFF', end_color='FFFFFFFF')
ZEBRA_FILL = PatternFill(fill_type='solid', start_color='EDF2F9', end_color='EDF2F9')

THIN_SIDE = Side(style='thin', color='B4C7E7')
ALL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

FMT_INT = '0'
FMT_NUM2 = '0.00'
FMT_NUM3 = '0.000'          # 学习成绩 / 综测成绩：3 位小数
FMT_QUANT = 'General'       # 量化成绩：整数无小数点，小数自然 1-2 位
FMT_PCT = '0.00%'

ROW_TITLE_HEIGHT = 30
ROW_HEADER_HEIGHT = 22
ROW_DATA_HEIGHT = 16.5

# 程序生成/会覆盖的计算sheet名（从源文件里剔除，避免重复追加）
COMPUTED_NAMES = ['量化', '量化处理', '成绩学分绩点处理', '排名百分比', '排名百分比计算',
                  '成绩处理', '综测成绩', '综测成绩备查表', '备查表', '备查表填写',
                  '参评', '参评科目']


def _apply_header(cell):
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = HEADER_FILL
    cell.border = ALL_BORDER


def _apply_data(cell, fmt=None, zebra=False):
    cell.font = DATA_FONT
    cell.alignment = DATA_ALIGN
    cell.fill = ZEBRA_FILL if zebra else DATA_FILL
    cell.border = ALL_BORDER
    if fmt:
        cell.number_format = fmt


def _write_title(ws, row, last_col, text):
    """合并第一行写标题，并给标题行加浅蓝底衬"""
    ws.cell(row, 1, text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    for c in range(1, last_col + 1):
        cc = ws.cell(row, c)
        cc.font = TITLE_FONT
        cc.alignment = TITLE_ALIGN
        cc.fill = TITLE_FILL
        cc.border = ALL_BORDER
    ws.row_dimensions[row].height = ROW_TITLE_HEIGHT


def _build_title(result, class_name, purpose):
    suffix = purpose if purpose.endswith('表') else f'{purpose}表'
    cn = str(class_name).strip() if class_name else ''
    # 班级名和标题用空格隔开；无班级名则空着
    return f'{cn} {suffix}'


def _set_widths(ws, widths):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


def write_result(workbook_path, result, class_name='综测', source_path=None):
    """写出结果 xlsx。result: calculator.process_all 的返回值。

    source_path: 上传的原始 xlsx 路径，保留其原有 sheet；为 None 则新建。
    """
    students = result['students']
    subjects = result['selected_subjects']
    months = result['months']
    quant_results = result['quant_results']
    grade_results = result['grade_results']
    n = len(students)
    nsubj = len(subjects)
    total_credits = result['total_credits']
    total_credits_with_quant = result['total_credits_with_quant']

    # ---------- 加载源文件（保留原 sheet） ----------
    if source_path:
        wb = openpyxl.load_workbook(source_path)  # 保留公式
        for name in list(wb.sheetnames):
            if name.strip() in COMPUTED_NAMES:
                del wb[name]
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    # ==================== Sheet (插入到第二个位置): 参评科目 ====================
    ws_eval = wb.create_sheet('参评科目', 1)  # index=1 → 第二个位置
    nsubj = len(subjects)
    last_eval = 3 + nsubj
    _write_title(ws_eval, 1, last_eval, _build_title(result, class_name, '参评科目'))

    # 第2行：表头（科目名）
    eval_headers = ['序号', '学号', '姓名']
    eval_headers += [s['name'] for s in subjects]
    for c, h in enumerate(eval_headers, 1):
        _apply_header(ws_eval.cell(2, c, h))
    ws_eval.row_dimensions[2].height = ROW_HEADER_HEIGHT

    # 第3行：学分行
    credit_labels = ['', '', '']
    credit_labels += [s.get('credit', 0) for s in subjects]
    for c, lab in enumerate(credit_labels, 1):
        cc = ws_eval.cell(3, c, lab)
        cc.font = Font(name=FONT_FAMILY, size=10, bold=True, color='1F4E79')
        cc.alignment = DATA_ALIGN
        cc.fill = PatternFill(fill_type='solid', start_color='D9E1F2', end_color='D9E1F2')
        cc.border = ALL_BORDER
        if c >= 4:
            cc.number_format = FMT_NUM2
    ws_eval.row_dimensions[3].height = ROW_HEADER_HEIGHT

    # 数据行（学生成绩）
    students = result['students']  # scores 已按科目名索引
    for i, st in enumerate(students):
        r = i + 4
        ws_eval.row_dimensions[r].height = ROW_DATA_HEIGHT
        zebra = (i % 2 == 1)
        ws_eval.cell(r, 1, i + 1)
        ws_eval.cell(r, 2, st.get('id', ''))
        ws_eval.cell(r, 3, st.get('name', ''))
        for j, subj in enumerate(subjects):
            score = st['scores'].get(subj['name'])
            cc = ws_eval.cell(r, 4 + j)
            cc.value = score if score is not None else None
            _apply_data(cc, zebra=zebra)
        # 应用边框/斑马纹到前3列
        for c in range(1, 4):
            _apply_data(ws_eval.cell(r, c), FMT_INT if c == 1 else None, zebra)

    _set_widths(ws_eval, {'A': 7, 'B': 13, 'C': 10})
    for j in range(nsubj):
        wcol = get_column_letter(4 + j)
        ws_eval.column_dimensions[wcol].width = 10
    ws_eval.freeze_panes = 'D4'

    # ==================== Sheet1: 量化 ====================
    ws1 = wb.create_sheet('量化')
    n_months = len(months)
    base_col = 7 + n_months        # 基础管理量化
    final_col = 8 + n_months       # 量化最终值
    last1 = final_col
    headers1 = ['排名', '学号', '姓名', '二课量化', '最终二课量化', '基础分']
    headers1 += list(months) + ['基础管理量化', '量化最终值']

    _write_title(ws1, 1, last1, _build_title(result, class_name, '量化处理'))
    for c, h in enumerate(headers1, 1):
        _apply_header(ws1.cell(2, c, h))
    ws1.row_dimensions[2].height = ROW_HEADER_HEIGHT

    final_letter = get_column_letter(final_col)
    for i, q in enumerate(quant_results):
        r = i + 3
        ws1.row_dimensions[r].height = ROW_DATA_HEIGHT
        zebra = (i % 2 == 1)
        ws1.cell(r, 1, f'=RANK({final_letter}{r},${final_letter}$3:${final_letter}${n+2},0)')
        ws1.cell(r, 2, q['id'])
        ws1.cell(r, 3, q['name'])
        ws1.cell(r, 4, q['second_class_raw'])
        ws1.cell(r, 5, f'=MIN(D{r},100)')
        ws1.cell(r, 6, 100)
        for j, m in enumerate(months):
            ws1.cell(r, 7 + j, q['deductions'].get(m, 0))
        if months:
            first_m = get_column_letter(7)
            last_m = get_column_letter(base_col - 1)
            # 若 base_management 与扣分格推算一致，用公式（可审计）；
            # 否则说明扣分格读不到、用了源表"总分"列，直接写值
            computed = 100 - sum(abs(v) for v in q['deductions'].values())
            if abs(computed - q['base_management']) < 0.01:
                ws1.cell(r, base_col, f'=100-SUMPRODUCT(ABS({first_m}{r}:{last_m}{r}))')
            else:
                ws1.cell(r, base_col, q['base_management'])
        else:
            ws1.cell(r, base_col, q['base_management'])
        ws1.cell(r, final_col, f'=(E{r}+{get_column_letter(base_col)}{r})/2')
        for c in range(1, last1 + 1):
            cc = ws1.cell(r, c)
            if c == 1:
                _apply_data(cc, FMT_INT, zebra)
            elif c in (4, 5, base_col, final_col):
                _apply_data(cc, FMT_QUANT, zebra)   # 量化类：整数无小数点
            else:
                _apply_data(cc, zebra=zebra)

    w1 = {'A': 7, 'B': 13, 'C': 10, 'D': 11, 'E': 13, 'F': 9}
    for j in range(n_months):
        w1[get_column_letter(7 + j)] = 7
    w1[get_column_letter(base_col)] = 14
    w1[get_column_letter(final_col)] = 13
    _set_widths(ws1, w1)
    ws1.freeze_panes = 'D3'

    # ==================== Sheet2: 成绩学分绩点处理（明细） ====================
    ws2 = wb.create_sheet('成绩学分绩点处理')
    score_start = 7
    quant_score_col = 7 + nsubj
    subj_point_start = 8 + nsubj
    subj_point_end = 7 + 2 * nsubj
    quant_point_col = 8 + 2 * nsubj
    avg_no_quant_col = 9 + 2 * nsubj
    avg_with_quant_col = 10 + 2 * nsubj
    last2 = avg_with_quant_col

    _write_title(ws2, 1, last2, _build_title(result, class_name, '成绩学分绩点处理'))

    sections = [
        (1, 3, '学生信息'),
        (4, 6, '排名部分'),
        (7, quant_score_col, '成绩部分'),
        (subj_point_start, quant_point_col, '学分绩点部分'),
        (avg_no_quant_col, avg_with_quant_col, '平均学分绩点部分'),
    ]
    for c1, c2, label in sections:
        ws2.cell(2, c1, label)
        if c2 > c1:
            ws2.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        for c in range(c1, c2 + 1):
            _apply_header(ws2.cell(2, c))
    ws2.row_dimensions[2].height = ROW_HEADER_HEIGHT

    col_names = ['序号', '学号', '姓名', '综合排名', '量化排名', '成绩排名']
    for j, subj in enumerate(subjects):
        col_names.append(subj['name'])
    col_names.append('综合量化测评')
    for j, subj in enumerate(subjects):
        col_names.append(subj['name'])
    col_names.append('综合量化测评')
    col_names.append('平均学分绩点\n不含量化')
    col_names.append('平均学分绩点\n含量化')
    for c, name in enumerate(col_names, 1):
        _apply_header(ws2.cell(3, c, name))
    ws2.row_dimensions[3].height = ROW_HEADER_HEIGHT

    labels4 = ['', '', '']
    labels4 += ['[量化+学习]', '[量化]', '[学习]']
    for j, subj in enumerate(subjects):
        labels4.append(f'[{subj.get("category", "")}{subj["credit"]}]')
    labels4.append('[3.0]')
    for j, subj in enumerate(subjects):
        labels4.append(f'[{subj.get("category", "")}{subj["credit"]}]')
    labels4.append('[3.0]')
    labels4.append(f'[{total_credits}]')
    labels4.append(f'[{total_credits_with_quant}]')
    for c, lab in enumerate(labels4, 1):
        _apply_header(ws2.cell(4, c, lab))
    ws2.row_dimensions[4].height = ROW_HEADER_HEIGHT

    quant_score_letter = get_column_letter(quant_score_col)
    avg_no_quant_letter = get_column_letter(avg_no_quant_col)
    avg_with_quant_letter = get_column_letter(avg_with_quant_col)
    sps = get_column_letter(subj_point_start)
    spe = get_column_letter(subj_point_end)
    qp = get_column_letter(quant_point_col)

    for i in range(n):
        r = i + 5
        ws2.row_dimensions[r].height = ROW_DATA_HEIGHT
        zebra = (i % 2 == 1)
        ws2.cell(r, 1, i + 1)
        ws2.cell(r, 2, grade_results[i]['id'])
        ws2.cell(r, 3, grade_results[i]['name'])
        ws2.cell(r, 4, f'=RANK({avg_with_quant_letter}{r},${avg_with_quant_letter}$5:${avg_with_quant_letter}${n+4},0)')
        ws2.cell(r, 5, f'=RANK({quant_score_letter}{r},${quant_score_letter}$5:${quant_score_letter}${n+4},0)')
        ws2.cell(r, 6, f'=RANK({avg_no_quant_letter}{r},${avg_no_quant_letter}$5:${avg_no_quant_letter}${n+4},0)')
        for j, subj in enumerate(subjects):
            score = grade_results[i]['subject_scores'].get(subj['name'])
            ws2.cell(r, score_start + j, score if score is not None else None)
        ws2.cell(r, quant_score_col, f'=量化!{final_letter}{r-2}')
        for j, subj in enumerate(subjects):
            score_ref = f'{get_column_letter(score_start + j)}{r}'
            ws2.cell(r, subj_point_start + j, f'=IF({score_ref}<60,0,({score_ref}-50)/10*{subj["credit"]})')
        ws2.cell(r, quant_point_col, f'=IF({quant_score_letter}{r}<60,0,({quant_score_letter}{r}-50)/10*3)')
        ws2.cell(r, avg_no_quant_col, f'=SUM({sps}{r}:{spe}{r})/{total_credits}')
        ws2.cell(r, avg_with_quant_col, f'=SUM({sps}{r}:{qp}{r})/{total_credits_with_quant}')
        for c in range(1, last2 + 1):
            cc = ws2.cell(r, c)
            if c in (1, 4, 5, 6):
                _apply_data(cc, FMT_INT, zebra)
            elif c == quant_score_col:
                _apply_data(cc, FMT_QUANT, zebra)        # 量化成绩：0.##
            elif c == avg_no_quant_col:
                _apply_data(cc, FMT_NUM3, zebra)         # 学习成绩：0.000
            elif c == avg_with_quant_col:
                _apply_data(cc, FMT_NUM3, zebra)         # 综测成绩：0.000
            else:
                _apply_data(cc, FMT_NUM2, zebra)

    w2 = {'A': 6, 'B': 13, 'C': 10, 'D': 11, 'E': 11, 'F': 11}
    for c in range(score_start, last2 + 1):
        w2[get_column_letter(c)] = 10
    w2[get_column_letter(avg_no_quant_col)] = 14
    w2[get_column_letter(avg_with_quant_col)] = 14
    _set_widths(ws2, w2)
    ws2.freeze_panes = 'D5'

    # ==================== Sheet3: 排名百分比（明细） ====================
    ws3 = wb.create_sheet('排名百分比')
    last3 = 11
    _write_title(ws3, 1, last3, _build_title(result, class_name, '排名百分比'))

    sections3 = [
        (1, 3, '学生信息'),
        (4, 5, '数据'),
        (6, 8, '排名部分'),
        (9, 11, '百分比部分'),
    ]
    for c1, c2, label in sections3:
        ws3.cell(2, c1, label)
        if c2 > c1:
            ws3.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        for c in range(c1, c2 + 1):
            _apply_header(ws3.cell(2, c))
    ws3.row_dimensions[2].height = ROW_HEADER_HEIGHT

    names3 = ['序号', '学号', '姓名', '平均学分绩点\n不含量化', '平均学分绩点\n含量化',
              '综合成绩排名', '量化成绩排名', '学习成绩排名',
              '综合成绩百分比', '量化成绩百分比', '学习成绩百分比']
    for c, name in enumerate(names3, 1):
        _apply_header(ws3.cell(3, c, name))
    ws3.row_dimensions[3].height = ROW_HEADER_HEIGHT

    labels3 = ['', '', '', '', '', '[量化+学习]', '[量化]', '[学习]', '[量化+学习]', '[量化]', '[学习]']
    for c, lab in enumerate(labels3, 1):
        _apply_header(ws3.cell(4, c, lab))
    ws3.row_dimensions[4].height = ROW_HEADER_HEIGHT

    for i in range(n):
        r = i + 5
        ws3.row_dimensions[r].height = ROW_DATA_HEIGHT
        zebra = (i % 2 == 1)
        ws3.cell(r, 1, i + 1)
        ws3.cell(r, 2, grade_results[i]['id'])
        ws3.cell(r, 3, grade_results[i]['name'])
        ws3.cell(r, 4, f'=成绩学分绩点处理!{avg_no_quant_letter}{r}')
        ws3.cell(r, 5, f'=成绩学分绩点处理!{avg_with_quant_letter}{r}')
        ws3.cell(r, 6, f'=成绩学分绩点处理!D{r}')
        ws3.cell(r, 7, f'=成绩学分绩点处理!E{r}')
        ws3.cell(r, 8, f'=成绩学分绩点处理!F{r}')
        ws3.cell(r, 9, f'=F{r}/{n}')
        ws3.cell(r, 10, f'=G{r}/{n}')
        ws3.cell(r, 11, f'=H{r}/{n}')
        for c in range(1, last3 + 1):
            cc = ws3.cell(r, c)
            if c == 1:
                _apply_data(cc, FMT_INT, zebra)
            elif c in (6, 7, 8):
                _apply_data(cc, FMT_INT, zebra)
            elif c in (4, 5):
                _apply_data(cc, FMT_NUM3, zebra)         # 学习/综测成绩：0.000
            else:
                _apply_data(cc, FMT_PCT, zebra)

    w3 = {'A': 6, 'B': 13, 'C': 10, 'D': 16, 'E': 16,
          'F': 14, 'G': 14, 'H': 14, 'I': 16, 'J': 16, 'K': 16}
    _set_widths(ws3, w3)
    ws3.freeze_panes = 'D5'

    # ==================== Sheet4: 备查表（12列，按综测名次排序） ====================
    # 学生 id -> 排名百分比/成绩学分绩点处理 中的数据行号 (i+5)
    rank_row_map = {g['id']: i + 5 for i, g in enumerate(grade_results)}
    # 按综测成绩(含量化平均分)降序 => 综测名次升序
    comp_order = sorted(grade_results, key=lambda g: (-g['avg_with_quant'], g['id']))

    ws4 = wb.create_sheet('备查表')
    last4 = 12
    _write_title(ws4, 1, last4, _build_title(result, class_name, '备查表'))

    # 三级表头：第2行分组，第3行明细
    groups = [(1, 3, '学生信息'), (4, 6, '学习成绩'), (7, 9, '量化成绩'), (10, 12, '综测成绩')]
    for c1, c2, label in groups:
        ws4.cell(2, c1, label)
        if c2 > c1:
            ws4.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        for c in range(c1, c2 + 1):
            _apply_header(ws4.cell(2, c))
    ws4.row_dimensions[2].height = ROW_HEADER_HEIGHT

    names4 = ['排名', '学号', '姓名',
              '学习成绩', '学习名次', '百分比',
              '量化成绩', '量化名次', '百分比',
              '综测成绩', '综测名次', '百分比']
    for c, name in enumerate(names4, 1):
        _apply_header(ws4.cell(3, c, name))
    ws4.row_dimensions[3].height = ROW_HEADER_HEIGHT

    for idx, g in enumerate(comp_order):
        r = idx + 4
        rr = rank_row_map[g['id']]
        ws4.row_dimensions[r].height = ROW_DATA_HEIGHT
        zebra = (idx % 2 == 1)
        ws4.cell(r, 1, idx + 1)
        ws4.cell(r, 2, g['id'])
        ws4.cell(r, 3, g['name'])
        # 学习
        ws4.cell(r, 4, f'=排名百分比!D{rr}')
        ws4.cell(r, 5, f'=排名百分比!H{rr}')
        ws4.cell(r, 6, f'=排名百分比!K{rr}')
        # 量化
        ws4.cell(r, 7, f'=成绩学分绩点处理!{quant_score_letter}{rr}')
        ws4.cell(r, 8, f'=排名百分比!G{rr}')
        ws4.cell(r, 9, f'=排名百分比!J{rr}')
        # 综测
        ws4.cell(r, 10, f'=排名百分比!E{rr}')
        ws4.cell(r, 11, f'=排名百分比!F{rr}')
        ws4.cell(r, 12, f'=排名百分比!I{rr}')
        for c in range(1, last4 + 1):
            cc = ws4.cell(r, c)
            if c == 1:
                _apply_data(cc, FMT_INT, zebra)
            elif c in (5, 8, 11):
                _apply_data(cc, FMT_INT, zebra)
            elif c == 4:
                _apply_data(cc, FMT_NUM3, zebra)         # 学习成绩：0.000
            elif c == 7:
                _apply_data(cc, FMT_QUANT, zebra)        # 量化成绩：0.##
            elif c == 10:
                _apply_data(cc, FMT_NUM3, zebra)         # 综测成绩：0.000
            else:
                _apply_data(cc, FMT_PCT, zebra)

    w4 = {'A': 6, 'B': 13, 'C': 10,
          'D': 12, 'E': 10, 'F': 11,
          'G': 12, 'H': 10, 'I': 11,
          'J': 12, 'K': 10, 'L': 11}
    _set_widths(ws4, w4)
    ws4.freeze_panes = 'D4'

    wb.save(workbook_path)
    return workbook_path