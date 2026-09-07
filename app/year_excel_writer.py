"""学年综测结果 xlsx 写出模块

结构（参考示例 Sheet1）：
  Sheet1: 学年综测主表（4行表头 + 数据行）
    - 第1行：标题合并
    - 第2行：分区标题合并（第一学期成绩部分/绩点部分/第二学期.../量化+学习/学习）
    - 第3行：列名（学号|姓名|各科名称×组|综测平均|排名|学习平均|排名|学号|姓名）
    - 第4行：学分标签 [必修0.5]/[限选3.0]等
    - 第5行起：数据行包含
      - 第一学期成绩列(s1_n列)
      - 第一学期绩点列(s1_n列) + 补考黄色标记
      - 量化绩点(3学分)
      - 第二学期成绩列(s2_n列)
      - 第二学期绩点列(s2_n列) + 补考黄色标记
      - 学年综测平均学分绩点(含量化)
      - 学年综测排名(RANK公式)
      - 学年学习平均学分绩点(不含量化)
      - 学年学习排名(RANK公式)
  Sheet2: 备查表 [学号 | 姓名 | 综测含量量 | 排名 | =排名/N% | 空 | 空 | 空 | 学习不含量量 | 排名 | =排名/N%]
"""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


FONT_FAMILY = 'Microsoft YaHei'
DATA_FONT = Font(name=FONT_FAMILY, size=11, bold=False)
HEADER_FONT = Font(name=FONT_FAMILY, size=11, bold=True)
TITLE_FONT = Font(name=FONT_FAMILY, size=14, bold=True)
SECTION_FONT = Font(name=FONT_FAMILY, size=11, bold=True)

DATA_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
HEADER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
TITLE_ALIGN = Alignment(horizontal='right', vertical='center')
SECTION_ALIGN = Alignment(horizontal='center', vertical='center')

DATA_FILL = PatternFill(fill_type='solid', start_color='FFFFFFFF', end_color='FFFFFFFF')
ZEBRA_FILL = PatternFill(fill_type='solid', start_color='EDF2F9', end_color='EDF2F9')
YELLOW_FILL = PatternFill(fill_type='solid', start_color='FFFF00', end_color='FFFF00')  # 补考黄色标记
ALL_BORDER = Border(
    left=Side(style='thin', color='B4C7E7'),
    right=Side(style='thin', color='B4C7E7'),
    top=Side(style='thin', color='B4C7E7'),
    bottom=Side(style='thin', color='B4C7E7')
)

FMT_INT = '0'
FMT_NUM2 = '0.00'
FMT_NUM3 = '0.000'
FMT_PCT = '0.00%'


def _apply_data(cell, fmt=None, zebra=False, no_fill=False):
    cell.font = DATA_FONT
    cell.alignment = DATA_ALIGN
    if not no_fill:
        cell.fill = ZEBRA_FILL if zebra else DATA_FILL
    cell.border = ALL_BORDER
    if fmt:
        cell.number_format = fmt


def write_year_result(workbook_path, result, class_name='', s1_path=None, s2_path=None):
    """写出学年综测结果 xlsx"""
    yearly_results = result['yearly_results']
    n = len(yearly_results)
    s1_subjects = result['s1_subjects']
    s2_subjects = result['s2_subjects']
    s1_total_credits_with_quant = result['s1_total_credits_with_quant']
    s2_total_credits_with_quant = result['s2_total_credits_with_quant']
    yearly_total_credits = result['yearly_total_credits']
    yearly_total_credits_with_quant = result['yearly_total_credits_with_quant']
    makeup_passed = set(result.get('makeup_passed', []))

    wb = openpyxl.Workbook()

    # ==================== Sheet1: 学年综测主表 ====================
    ws = wb.active
    ws.title = '学年综测'

    # ---- 计算列位置 ----
    # A=1: 学号, B=2: 姓名
    # C~C+N-1 = 2+len(s1_subjects): 第一学期成绩
    s1_score_start = 3  # C列
    s1_score_end = s1_score_start + len(s1_subjects) - 1

    # 第一学期的综合量化测评
    s1_quant_score_col = s1_score_end + 1

    # 接下来是 S1绩点列
    s1_point_start = s1_score_start + len(s1_subjects)  # 在成绩后面？不，参照示例：成绩和绩点是分开的两块
    # 看示例：S1成绩(11门)+S1量化(1列)+S1绩点(11门)+S1量化绩点(1列)+S2成绩(11门)+S2量化(1列)+S2绩点(11门)+S2量化绩点(1列)+综测平均+排名+学习平均+排名+末尾学号姓名

    # 重新规划列位置（完全参照示例）：
    col_id = 1            # A: 学号
    col_name = 2          # B: 姓名
    col_s1_score = 3      # C: 第一学期成绩开始
    col_s1_quant_score = col_s1_score + len(s1_subjects)  # 第一学期综合量化测评
    col_s1_point = col_s1_quant_score + 1  # 第一学期绩点开始
    col_s1_quant_point = col_s1_point + len(s1_subjects)  # 第一学期量化绩点
    col_s2_score = col_s1_quant_point + 1  # 第二学期成绩开始
    col_s2_quant_score = col_s2_score + len(s2_subjects)  # 第二学期综合量化测评
    col_s2_point = col_s2_quant_score + 1  # 第二学期绩点开始
    col_s2_quant_point = col_s2_point + len(s2_subjects)  # 第二学期量化绩点
    col_avg = col_s2_quant_point + 1  # 学年综测平均
    col_rank = col_avg + 1  # 学年综测排名
    col_study = col_rank + 1  # 学年学习平均
    col_study_rank = col_study + 1  # 学年学习排名
    col_end_id = col_study_rank + 1  # 末尾学号
    col_end_name = col_study_rank + 2  # 末尾姓名

    total_cols = col_end_name

    # ---- 第1行：分区标题合并 ----
    # 第一学期成绩部分
    if s1_subjects:
        ws.merge_cells(f'{get_column_letter(col_s1_score)}1:{get_column_letter(col_s1_quant_score - 1)}1')
        cell = ws.cell(1, col_s1_score, f'第一学期成绩部分[{s1_total_credits_with_quant}]')
        cell.font = SECTION_FONT
        cell.alignment = SECTION_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
        cell.border = ALL_BORDER

    # 第一学期绩点部分
    if s1_subjects:
        ws.merge_cells(f'{get_column_letter(col_s1_point)}1:{get_column_letter(col_s1_quant_point - 1)}1')
        cell = ws.cell(1, col_s1_point, f'第一学期学分绩点部分[{s1_total_credits_with_quant}]')
        cell.font = SECTION_FONT
        cell.alignment = SECTION_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
        cell.border = ALL_BORDER

    # 第二学期成绩部分
    if s2_subjects:
        ws.merge_cells(f'{get_column_letter(col_s2_score)}1:{get_column_letter(col_s2_quant_score - 1)}1')
        cell = ws.cell(1, col_s2_score, f'第二学期成绩部分[{s2_total_credits_with_quant}]')
        cell.font = SECTION_FONT
        cell.alignment = SECTION_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
        cell.border = ALL_BORDER

    # 第二学期绩点部分
    if s2_subjects:
        ws.merge_cells(f'{get_column_letter(col_s2_point)}1:{get_column_letter(col_s2_quant_point - 1)}1')
        cell = ws.cell(1, col_s2_point, f'第二学期学分绩点部分[{s2_total_credits_with_quant}]')
        cell.font = SECTION_FONT
        cell.alignment = SECTION_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
        cell.border = ALL_BORDER

    # 量化+学习（覆盖综测平均+综测排名）
    ws.merge_cells(f'{get_column_letter(col_avg)}1:{get_column_letter(col_rank)}1')
    cell = ws.cell(1, col_avg, '量化+学习')
    cell.font = SECTION_FONT
    cell.alignment = SECTION_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
    cell.border = ALL_BORDER

    # 学习（覆盖学习平均+学习排名）
    ws.merge_cells(f'{get_column_letter(col_study)}1:{get_column_letter(col_study_rank)}1')
    cell = ws.cell(1, col_study, '学习')
    cell.font = SECTION_FONT
    cell.alignment = SECTION_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
    cell.border = ALL_BORDER

    # 末尾学号姓名
    for col, label in [(col_end_id, '学号'), (col_end_name, '姓名')]:
        cell = ws.cell(1, col, label)
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='BDD7EE', end_color='BDD7EE')
        cell.border = ALL_BORDER

    ws.row_dimensions[1].height = 25

    # ---- 第2行：列名 ----
    row_header = 2
    for col, label in [(col_id, '学号'), (col_name, '姓名')]:
        cell = ws.cell(row_header, col, label)
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
        cell.border = ALL_BORDER

    # 第一学期成绩列名
    for j, subj in enumerate(s1_subjects):
        col = col_s1_score + j
        cell = ws.cell(row_header, col, subj['name'])
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
        cell.border = ALL_BORDER

    # 综合量化测评
    cell = ws.cell(row_header, col_s1_quant_score, '综合量化测评')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
    cell.border = ALL_BORDER

    # 第一学期绩点列名
    for j, subj in enumerate(s1_subjects):
        col = col_s1_point + j
        cell = ws.cell(row_header, col, subj['name'])
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
        cell.border = ALL_BORDER

    # 量化绩点
    cell = ws.cell(row_header, col_s1_quant_point, '综合量化测评')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
    cell.border = ALL_BORDER

    # 第二学期成绩列名
    for j, subj in enumerate(s2_subjects):
        col = col_s2_score + j
        cell = ws.cell(row_header, col, subj['name'])
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
        cell.border = ALL_BORDER

    # 第二学期综合量化测评
    cell = ws.cell(row_header, col_s2_quant_score, '综合量化测评')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
    cell.border = ALL_BORDER

    # 第二学期绩点列名
    for j, subj in enumerate(s2_subjects):
        col = col_s2_point + j
        cell = ws.cell(row_header, col, subj['name'])
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
        cell.border = ALL_BORDER

    # 量化绩点
    cell = ws.cell(row_header, col_s2_quant_point, '综合量化测评')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
    cell.border = ALL_BORDER

    # 学年综测平均学分绩
    cell = ws.cell(row_header, col_avg, '学年综测\n平均学分绩')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='FCE4D6', end_color='FCE4D6')
    cell.border = ALL_BORDER

    # 排名
    cell = ws.cell(row_header, col_rank, '排名')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='FCE4D6', end_color='FCE4D6')
    cell.border = ALL_BORDER

    # 学年学习平均学分绩
    cell = ws.cell(row_header, col_study, '学年学习\n平均学分绩')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='FCE4D6', end_color='FCE4D6')
    cell.border = ALL_BORDER

    # 排名
    cell = ws.cell(row_header, col_study_rank, '排名')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='FCE4D6', end_color='FCE4D6')
    cell.border = ALL_BORDER

    # 末尾学号姓名
    ws.cell(row_header, col_end_id, '学号')
    ws.cell(row_header, col_end_name, '姓名')

    for c in [col_end_id, col_end_name]:
        cc = ws.cell(row_header, c)
        cc.font = HEADER_FONT
        cc.alignment = HEADER_ALIGN
        cc.fill = PatternFill(fill_type='solid', start_color='FCE4D6', end_color='FCE4D6')
        cc.border = ALL_BORDER

    ws.row_dimensions[2].height = 35

    # ---- 第3行：学分标签 ----
    row_credit = 3
    for j, subj in enumerate(s1_subjects):
        col = col_s1_score + j
        credit = subj.get('credit', 0)
        credit_type = subj.get('credit_type', '必修')
        cell = ws.cell(row_credit, col, f'[{credit_type}{credit}]')
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
        cell.border = ALL_BORDER

    # 综合量化测评学分
    cell = ws.cell(row_credit, col_s1_quant_score, '[3.0]')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
    cell.border = ALL_BORDER

    for j, subj in enumerate(s1_subjects):
        col = col_s1_point + j
        credit = subj.get('credit', 0)
        credit_type = subj.get('credit_type', '必修')
        cell = ws.cell(row_credit, col, f'[{credit_type}{credit}]')
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
        cell.border = ALL_BORDER

    # 量化绩点学分
    cell = ws.cell(row_credit, col_s1_quant_point, '[3.0]')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
    cell.border = ALL_BORDER

    for j, subj in enumerate(s2_subjects):
        col = col_s2_score + j
        credit = subj.get('credit', 0)
        credit_type = subj.get('credit_type', '必修')
        cell = ws.cell(row_credit, col, f'[{credit_type}{credit}]')
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
        cell.border = ALL_BORDER

    # 第二学期综合量化测评学分
    cell = ws.cell(row_credit, col_s2_quant_score, '[3.0]')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='D6EAF8', end_color='D6EAF8')
    cell.border = ALL_BORDER

    for j, subj in enumerate(s2_subjects):
        col = col_s2_point + j
        credit = subj.get('credit', 0)
        credit_type = subj.get('credit_type', '必修')
        cell = ws.cell(row_credit, col, f'[{credit_type}{credit}]')
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
        cell.border = ALL_BORDER

    # 量化绩点学分
    cell = ws.cell(row_credit, col_s2_quant_point, '[3.0]')
    cell.font = HEADER_FONT
    cell.alignment = HEADER_ALIGN
    cell.fill = PatternFill(fill_type='solid', start_color='E2EFDA', end_color='E2EFDA')
    cell.border = ALL_BORDER

    ws.row_dimensions[3].height = 20

    # ---- 第4行起：数据行 ----
    data_start_row = 4

    # 按学年综测平均学分绩降序排列
    yearly_sorted = sorted(yearly_results, key=lambda x: -x['yearly_avg_with_quant'])

    for i, r in enumerate(yearly_sorted):
        row = data_start_row + i
        ws.row_dimensions[row].height = 20
        zebra = (i % 2 == 1)

        fill = ZEBRA_FILL if zebra else DATA_FILL

        # A: 学号
        cc = ws.cell(row, col_id, str(r['id']))
        _apply_data(cc, fmt=FMT_INT, zebra=zebra)

        # B: 姓名
        cc = ws.cell(row, col_name, r['name'])
        _apply_data(cc, zebra=zebra)

        # C~C+N-1: 第一学期成绩
        for j, subj in enumerate(s1_subjects):
            col = col_s1_score + j
            score = r['s1_scores'].get(subj['name'])
            cc = ws.cell(row, col, score if score is not None else '')
            _apply_data(cc, fmt=FMT_INT if isinstance(score, (int, float)) else None, zebra=zebra)

        # 综合量化测评(第一学期的)
        quant_final_1 = r.get('s1_quant_final', r['yearly_quant_final'])
        cc = ws.cell(row, col_s1_quant_score, round(quant_final_1, 3))
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra)

        # N+1~N+M: 第一学期绩点（公式）
        for j, subj in enumerate(s1_subjects):
            col = col_s1_point + j
            score_col = col_s1_score + j
            score_cell_ref = f'{get_column_letter(score_col)}{row}'
            credit = subj.get('credit', 0)
            course_name = subj['name']
            key = f"{r['id']}||{course_name}"
            is_makeup = key in makeup_passed

            # 绩点: 补考通过直接写1；否则用公式 =(分数-50)/10*学分，挂科则0，空值则0
            if is_makeup:
                cc = ws.cell(row, col, 1)
                _apply_data(cc, zebra=zebra)
                cc.fill = YELLOW_FILL
            else:
                formula = f'=IF(OR({score_cell_ref}<60,{score_cell_ref}=""),0,({score_cell_ref}-50)/10*{credit})'
                cc = ws.cell(row, col, formula)
                _apply_data(cc, zebra=zebra)

        # 量化绩点（第一学期）
        s1_qp = r.get('s1_quant_point')
        if s1_qp is None:
            from calculator import calc_quant_grade_point
            s1_qp = calc_quant_grade_point(r.get('s1_quant_final', r['yearly_quant_final']))
        cc = ws.cell(row, col_s1_quant_point, round(s1_qp, 2))
        _apply_data(cc, fmt=FMT_NUM2, zebra=zebra)

        # 第二学期成绩
        for j, subj in enumerate(s2_subjects):
            col = col_s2_score + j
            score = r['s2_scores'].get(subj['name'])
            cc = ws.cell(row, col, score if score is not None else '')
            _apply_data(cc, fmt=FMT_INT if isinstance(score, (int, float)) else None, zebra=zebra)

        # 综合量化测评(第二学期的)
        cc = ws.cell(row, col_s2_quant_score, round(r.get('s2_quant_final', r['yearly_quant_final']), 3))
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra)

        # 第二学期绩点（公式）
        for j, subj in enumerate(s2_subjects):
            col = col_s2_point + j
            score_col = col_s2_score + j
            score_cell_ref = f'{get_column_letter(score_col)}{row}'
            credit = subj.get('credit', 0)
            course_name = subj['name']
            key = f"{r['id']}||{course_name}"
            is_makeup = key in makeup_passed

            # 绩点: 补考通过直接写1；否则用公式
            if is_makeup:
                cc = ws.cell(row, col, 1)
                _apply_data(cc, zebra=zebra)
                cc.fill = YELLOW_FILL
            else:
                formula = f'=IF(OR({score_cell_ref}<60,{score_cell_ref}=""),0,({score_cell_ref}-50)/10*{credit})'
                cc = ws.cell(row, col, formula)
                _apply_data(cc, zebra=zebra)

        # 量化绩点（第二学期）
        s2_qp = r.get('s2_quant_point')
        if s2_qp is None:
            from calculator import calc_quant_grade_point
            s2_qp = calc_quant_grade_point(r.get('s2_quant_final', r['yearly_quant_final']))
        cc = ws.cell(row, col_s2_quant_point, round(s2_qp, 2))
        _apply_data(cc, fmt=FMT_NUM2, zebra=zebra)

        # 学年综测平均学分绩（含量化）
        # (S1绩点总和 + S2绩点总和 + S1量化绩点 + S2量化绩点) / 总学分
        s1_pts_start = get_column_letter(col_s1_point)
        s1_pts_end = get_column_letter(col_s1_quant_point - 1)
        s2_pts_start = get_column_letter(col_s2_point)
        s2_pts_end = get_column_letter(col_s2_quant_point - 1)
        s1_qp_col = get_column_letter(col_s1_quant_point)
        s2_qp_col = get_column_letter(col_s2_quant_point)
        avg_formula = (f'=SUM({s1_pts_start}{row}:{s1_pts_end}{row},'
                       f'{s2_pts_start}{row}:{s2_pts_end}{row},'
                       f'{s1_qp_col}{row},{s2_qp_col}{row})'
                       f'/{yearly_total_credits_with_quant}')
        cc = ws.cell(row, col_avg, avg_formula)
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra)

        # 学年综测排名（RANK公式）
        avg_letter = get_column_letter(col_avg)
        rank_range = f'{avg_letter}${data_start_row}:${avg_letter}${data_start_row + n - 1}'
        cc = ws.cell(row, col_rank, f'=RANK({avg_letter}{row},{rank_range},0)')
        _apply_data(cc, fmt=FMT_INT, zebra=zebra)

        # 学年学习平均学分绩（不含量化）
        study_formula = (f'=(SUM({s1_pts_start}{row}:{s1_pts_end}{row})'
                         f'+SUM({s2_pts_start}{row}:{s2_pts_end}{row}))'
                         f'/{yearly_total_credits}')
        cc = ws.cell(row, col_study, study_formula)
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra)

        # 学年学习排名（RANK公式）
        study_letter = get_column_letter(col_study)
        study_rank_range = f'{study_letter}${data_start_row}:${study_letter}${data_start_row + n - 1}'
        cc = ws.cell(row, col_study_rank, f'=RANK({study_letter}{row},{study_rank_range},0)')
        _apply_data(cc, fmt=FMT_INT, zebra=zebra)

        # 末尾学号姓名
        ws.cell(row, col_end_id, str(r['id']))
        ws.cell(row, col_end_name, r['name'])

        for c in [col_end_id, col_end_name]:
            cc = ws.cell(row, c)
            cc.font = DATA_FONT
            cc.alignment = DATA_ALIGN
            cc.fill = fill
            cc.border = ALL_BORDER

    # ---- 列宽设置 ----
    ws.column_dimensions['A'].width = 14.5
    ws.column_dimensions['B'].width = 10
    for c in range(col_s1_score, col_end_id):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions[get_column_letter(col_avg)].width = 18
    ws.column_dimensions[get_column_letter(col_rank)].width = 8
    ws.column_dimensions[get_column_letter(col_study)].width = 18
    ws.column_dimensions[get_column_letter(col_study_rank)].width = 8
    ws.column_dimensions[get_column_letter(col_end_id)].width = 14.5
    ws.column_dimensions[get_column_letter(col_end_name)].width = 10

    # 冻结窗格（A:B列，从第5行开始）
    ws.freeze_panes = f'C{data_start_row}'

    # ==================== Sheet2: 备查表（旧版Sheet1简化格式）====================
    ws2 = wb.create_sheet('备查表')

    # 按综测含量量降序排列
    backup_sorted = sorted(yearly_results, key=lambda x: -x['yearly_avg_with_quant'])

    backup_align = Alignment(horizontal='center', vertical='center')
    backup_font = Font(name=FONT_FAMILY, size=11, bold=False, color='FF000000')
    backup_font_bold = Font(name='SimHei', size=14, bold=True, color='FF000000')
    all_border2 = Border(
        left=Side(style='thin', color='B4C7E7'),
        right=Side(style='thin', color='B4C7E7'),
        top=Side(style='thin', color='B4C7E7'),
        bottom=Side(style='thin', color='B4C7E7')
    )

    # 列定义（参考旧版Sheet1）
    b_col_id = 1           # A: 学号
    b_col_name = 2         # B: 姓名
    b_col_qcgp = 3         # C: 综测含量量
    b_col_qcrank = 4       # D: 排名
    b_col_qcpct = 5        # E: =排名/N%
    b_col_blank1 = 6       # F: 空
    b_col_blank2 = 7       # G: 空
    b_col_blank3 = 8       # H: 空
    b_col_studyavg = 9     # I: 学习不含量量
    b_col_study_rank = 10  # J: 排名
    b_col_studypct = 11    # K: =排名/N%

    for i, r in enumerate(backup_sorted):
        row = i + 1
        ws2.row_dimensions[row].height = 20
        zebra = (i % 2 == 1)

        # A: 学号
        cc = ws2.cell(row, b_col_id, str(r['id']))
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_INT, zebra=zebra, no_fill=True)

        # B: 姓名
        cc = ws2.cell(row, b_col_name, r['name'])
        cc.font = backup_font
        _apply_data(cc, zebra=zebra, no_fill=True)

        # C: 综测含量量（保留3位小数）
        cc = ws2.cell(row, b_col_qcgp, round(r['yearly_avg_with_quant'], 3))
        cc.font = backup_font_bold
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra, no_fill=True)

        # D: 排名（RANK 公式）
        qc_gp_letter = get_column_letter(b_col_qcgp)
        rank_letter = get_column_letter(b_col_qcrank)
        cc = ws2.cell(row, b_col_qcrank, f'=RANK({qc_gp_letter}{row},${qc_gp_letter}$1:${qc_gp_letter}${n},0)')
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_INT, zebra=zebra, no_fill=True)

        # E: =排名/总人数（百分比格式）
        pct_letter = get_column_letter(b_col_qcpct)
        cc = ws2.cell(row, b_col_qcpct, f'={rank_letter}{row}/{n}')
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_PCT, zebra=zebra, no_fill=True)

        # F, G, H: 空白占位列
        for c in range(b_col_blank1, b_col_blank1 + 3):
            cc = ws2.cell(row, c)
            cc.font = backup_font
            _apply_data(cc, zebra=zebra, no_fill=True)

        # I: 学习不含量量（保留3位小数）
        cc = ws2.cell(row, b_col_studyavg, round(r['yearly_avg_no_quant'], 3))
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_NUM3, zebra=zebra, no_fill=True)

        # J: 排名（RANK 公式）
        study_avg_letter = get_column_letter(b_col_studyavg)
        study_rank_letter = get_column_letter(b_col_study_rank)
        cc = ws2.cell(row, b_col_study_rank, f'=RANK({study_avg_letter}{row},${study_avg_letter}$1:${study_avg_letter}${n},0)')
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_INT, zebra=zebra, no_fill=True)

        # K: =排名/总人数（百分比格式）
        study_pct_letter = get_column_letter(b_col_studypct)
        cc = ws2.cell(row, b_col_studypct, f'={study_rank_letter}{row}/{n}')
        cc.font = backup_font
        _apply_data(cc, fmt=FMT_PCT, zebra=zebra, no_fill=True)

    ws2.column_dimensions['A'].width = 14.5
    for c in range(2, b_col_studypct + 1):
        ws2.column_dimensions[get_column_letter(c)].width = 10
    ws2.column_dimensions[get_column_letter(b_col_qcpct)].width = 13
    ws2.column_dimensions[get_column_letter(b_col_study_rank)].width = 13
    ws2.column_dimensions[get_column_letter(b_col_studypct)].width = 13

    wb.save(workbook_path)
    return workbook_path
