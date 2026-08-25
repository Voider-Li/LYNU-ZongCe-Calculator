"""综测计算核心逻辑"""
from openpyxl.utils import get_column_letter


def calc_second_class_final(raw):
    """二课量化最终值：>100 取 100"""
    return min(raw, 100)


def calc_base_management(deductions):
    """基础管理量化 = 100 - Σ|各月扣分|（扣分后值）。
    无论用户填 -2 还是 2，均按扣 2 分处理，确保结果为扣分后的分值。"""
    total = sum(abs(v) for v in deductions.values())
    return 100 - total


def calc_quant_final(second_class_raw, base_management):
    """量化最终值 = (最终二课量化 + 基础管理量化) / 2"""
    sc_final = calc_second_class_final(second_class_raw)
    return (sc_final + base_management) / 2


def calc_grade_point(score, credit):
    """单科绩点: IF(成绩<60, 0, (成绩-50)/10 × 学分)"""
    if score < 60:
        return 0
    return ((score - 50) / 10) * credit


def calc_quant_grade_point(quant_final):
    """量化绩点（3学分）: IF(量化<60, 0, (量化-50)/10 × 3)"""
    if quant_final < 60:
        return 0
    return ((quant_final - 50) / 10) * 3


def calc_rank_desc(values):
    """
    RANK 降序：值越大名次越小（第1名最好）
    返回 {index: rank}
    """
    n = len(values)
    indexed = list(enumerate(values))
    # 降序排列
    sorted_vals = sorted(indexed, key=lambda x: -x[1])
    ranks = {}
    i = 0
    while i < n:
        # 找并列
        j = i
        while j + 1 < n and sorted_vals[j + 1][1] == sorted_vals[i][1]:
            j += 1
        rank = i + 1
        for k in range(i, j + 1):
            ranks[sorted_vals[k][0]] = rank
        i = j + 1
    return ranks


def calc_percentage(rank, total):
    """百分比 = 排名 / 总人数"""
    if total == 0:
        return 0
    return rank / total


def process_all(raw_data, eval_subjects, quant_data, selected_subjects):
    """
    全流程计算。
    参数:
      raw_data: read_raw_scores 返回值
      eval_subjects: read_eval_subjects 返回值 或 None
      quant_data: read_quantization 返回值
      selected_subjects: [{name, credit}] 用户选定的参评科目

    返回: 完整结果字典
    """
    # 1. 确定学生列表和科目
    # 优先用 eval_subjects 的学生和成绩
    if eval_subjects:
        students = eval_subjects['students']
    else:
        # 从原始表提取
        students = raw_data['students']
        # 把 scores 的 key 从 col_idx 转成科目名
        course_map = {c['col_idx']: c['name'] for c in raw_data['courses']}
        for s in students:
            new_scores = {}
            for col, score in s['scores'].items():
                cname = course_map.get(col)
                if cname:
                    new_scores[cname] = score
            s['scores'] = new_scores

    # 2. 量化计算
    quant_map = {s['id']: s for s in quant_data['students']} if quant_data else {}
    quant_results = []
    for s in students:
        q = quant_map.get(s['id'], {})
        sc_raw = q.get('second_class_raw', 0)
        deductions = q.get('deductions', {})
        sc_final = calc_second_class_final(sc_raw)
        # 基础管理量化：优先用源表"总分"列（Excel 已算好的扣分后值），
        # 没有则由扣分格计算（100 - Σ|扣分|）
        base_mgmt_src = q.get('base_mgmt_source')
        if base_mgmt_src is not None:
            base_mgmt = base_mgmt_src
        else:
            base_mgmt = calc_base_management(deductions)
        quant_final = calc_quant_final(sc_raw, base_mgmt)
        quant_results.append({
            'id': s['id'],
            'name': s['name'],
            'second_class_raw': sc_raw,
            'second_class_final': sc_final,
            'base_score': 100,
            'deductions': deductions,
            'base_management': base_mgmt,
            'quant_final': quant_final,
        })

    # 量化排名
    quant_finals = [q['quant_final'] for q in quant_results]
    quant_ranks = calc_rank_desc(quant_finals)
    for i, q in enumerate(quant_results):
        q['quant_rank'] = quant_ranks[i]

    # 3. 成绩绩点计算
    total_credits = sum(s['credit'] for s in selected_subjects)
    total_credits_with_quant = total_credits + 3  # 量化3学分

    grade_results = []
    for i, s in enumerate(students):
        subject_scores = {}
        subject_points = {}
        total_point = 0  # 不含量化
        for subj in selected_subjects:
            score = s['scores'].get(subj['name'])
            if score is None:
                score = 0
            point = calc_grade_point(score, subj['credit'])
            subject_scores[subj['name']] = score
            subject_points[subj['name']] = point
            total_point += point

        # 量化绩点
        quant_final = quant_results[i]['quant_final']
        quant_point = calc_quant_point(quant_final)
        total_point_with_quant = total_point + quant_point

        # 平均学分绩点
        avg_no_quant = total_point / total_credits if total_credits > 0 else 0
        avg_with_quant = total_point_with_quant / total_credits_with_quant if total_credits_with_quant > 0 else 0

        grade_results.append({
            'id': s['id'],
            'name': s['name'],
            'subject_scores': subject_scores,
            'subject_points': subject_points,
            'quant_score': quant_final,
            'quant_point': quant_point,
            'total_point': total_point,
            'total_point_with_quant': total_point_with_quant,
            'avg_no_quant': avg_no_quant,
            'avg_with_quant': avg_with_quant,
        })

    # 排名
    avg_no_quants = [g['avg_no_quant'] for g in grade_results]
    avg_with_quants = [g['avg_with_quant'] for g in grade_results]

    study_ranks = calc_rank_desc(avg_no_quants)
    composite_ranks = calc_rank_desc(avg_with_quants)

    for i, g in enumerate(grade_results):
        g['study_rank'] = study_ranks[i]
        g['composite_rank'] = composite_ranks[i]

    # 4. 百分比
    n = len(students)
    for i, q in enumerate(quant_results):
        q['quant_percent'] = calc_percentage(q['quant_rank'], n)
    for i, g in enumerate(grade_results):
        g['study_percent'] = calc_percentage(g['study_rank'], n)
        g['composite_percent'] = calc_percentage(g['composite_rank'], n)

    # 5. 备查表（按综测成绩降序）
    backup = []
    for i in range(len(students)):
        backup.append({
            'id': students[i]['id'],
            'name': students[i]['name'],
            'study_score': grade_results[i]['avg_no_quant'],
            'study_rank': grade_results[i]['study_rank'],
            'study_percent': grade_results[i]['study_percent'],
            'quant_score': quant_results[i]['quant_final'],
            'quant_rank': quant_results[i]['quant_rank'],
            'quant_percent': quant_results[i]['quant_percent'],
            'composite_score': grade_results[i]['avg_with_quant'],
            'composite_rank': grade_results[i]['composite_rank'],
            'composite_percent': grade_results[i]['composite_percent'],
        })
    backup.sort(key=lambda x: -x['composite_score'])
    # 重新排备查表名次
    for idx, b in enumerate(backup):
        b['backup_rank'] = idx + 1

    return {
        'students': students,
        'selected_subjects': selected_subjects,
        'total_credits': total_credits,
        'total_credits_with_quant': total_credits_with_quant,
        'quant_results': quant_results,
        'grade_results': grade_results,
        'backup': backup,
        'months': quant_data['months'] if quant_data else [],
        'title': raw_data.get('title', '') if raw_data else '',
    }


# 兼容调用
def calc_quant_point(quant_final):
    return calc_quant_grade_point(quant_final)
