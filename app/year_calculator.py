"""学年综测计算模块

流程：
  1. 读取两个学期的综测结果文件（学期综测导出的 xlsx）
  2. 直接使用参评科目表已选好的科目和成绩，不重新选、不调用 process_all
  3. 识别挂科科目（分数 < 60 且非空）
  4. 用户勾选补考通过 → 该科绩点改为 1
  5. 合并两学期，计算学年平均学分绩点 + 排名
"""
from excel_reader import read_eval_subjects, read_quantization
from calculator import (
    calc_grade_point, calc_quant_grade_point,
    calc_rank_desc, calc_percentage, calc_quant_final, calc_base_management,
    calc_second_class_final,
)


def read_semester(filepath):
    """读取一个学期的综测结果文件，直接用参评科目表的数据。

    不调用 process_all，不重新选科目，直接从参评科目表读成绩+学分算绩点。

    返回: {
        students: [{id, name, scores: {course_name: score}, subject_points: {course_name: gp}}],
        subjects: [{name, credit}],
        total_credits: float,
        total_credits_with_quant: float,
        quant_points: {student_id: quant_point},
    }
    """
    eval_sub = read_eval_subjects(filepath)
    quant = read_quantization(filepath)

    if eval_sub:
        students_raw = eval_sub['students']
        subjects = [{'name': s['name'], 'credit': s.get('credit') or 0} for s in eval_sub['subjects']]
    else:
        # 没有参评科目表，退化：读不到就报错
        raise ValueError('未找到参评科目表，请确保上传的是学期综测导出的结果文件')

    total_credits = sum(s['credit'] for s in subjects)
    total_credits_with_quant = total_credits + 3  # 量化3学分

    # 量化数据
    quant_map = {s['id']: s for s in quant['students']} if quant else {}

    # 直接算每科绩点 + 量化绩点
    students = []
    quant_points = {}
    quant_finals = {}
    for s in students_raw:
        sid = s['id']
        sname = s['name']
        scores = {}
        subject_points = {}
        for subj in subjects:
            score = s['scores'].get(subj['name'])
            # 空值 = 没选这门课（如英语生→日语列空），保持 None，不转0
            if score is None or score == '':
                scores[subj['name']] = None
                subject_points[subj['name']] = 0  # 没选 = 绩点0，但不算挂科
            else:
                scores[subj['name']] = score
                subject_points[subj['name']] = calc_grade_point(score, subj['credit'])

        # 量化绩点
        q = quant_map.get(sid, {})
        sc_raw = q.get('second_class_raw', 0)
        deductions = q.get('deductions', {})
        base_mgmt_src = q.get('base_mgmt_source')
        if base_mgmt_src is not None:
            base_mgmt = base_mgmt_src
        else:
            base_mgmt = calc_base_management(deductions)
        quant_final = calc_quant_final(sc_raw, base_mgmt)
        qp = calc_quant_grade_point(quant_final)
        quant_points[sid] = qp
        quant_finals[sid] = quant_final

        students.append({
            'id': sid,
            'name': sname,
            'scores': scores,
            'subject_points': subject_points,
        })

    return {
        'students': students,
        'subjects': subjects,
        'total_credits': total_credits,
        'total_credits_with_quant': total_credits_with_quant,
        'quant_points': quant_points,
        'quant_finals': quant_finals,
    }


def find_failed_courses(semester_data):
    """找出所有挂科（分数是数值且 < 60，空值跳过）。

    返回: [{student_id, student_name, course, score}]
    """
    failed = []
    for s in semester_data['students']:
        for subj in semester_data['subjects']:
            score = s['scores'].get(subj['name'])
            if score is None or score == '':
                continue  # 空值 = 没选这门课，跳过
            if isinstance(score, (int, float)) and score < 60:
                failed.append({
                    'student_id': s['id'],
                    'student_name': s['name'],
                    'course': subj['name'],
                    'score': score,
                })
    return failed


def calc_yearly(semester1_data, semester2_data, makeup_passed):
    """计算学年综测。

    makeup_passed: list of "student_id||course_name" 字符串，
                   表示该学生该科补考通过 → 绩点 = 1
    """
    makeup_set = set(makeup_passed)

    s1_total_credits = semester1_data['total_credits']
    s1_total_credits_with_quant = semester1_data['total_credits_with_quant']
    s2_total_credits = semester2_data['total_credits']
    s2_total_credits_with_quant = semester2_data['total_credits_with_quant']

    yearly_total_credits = s1_total_credits + s2_total_credits
    yearly_total_credits_with_quant = s1_total_credits_with_quant + s2_total_credits_with_quant

    # 按学号匹配学生
    s1_map = {s['id']: s for s in semester1_data['students']}
    s2_map = {s['id']: s for s in semester2_data['students']}
    s1_qp = semester1_data['quant_points']
    s2_qp = semester2_data['quant_points']
    s1_qf = semester1_data['quant_finals']
    s2_qf = semester2_data['quant_finals']

    all_ids = set(s1_map.keys()) & set(s2_map.keys())  # 两学期都有的学生

    yearly_results = []
    for sid in all_ids:
        s1 = s1_map[sid]
        s2 = s2_map[sid]
        name = s1['name']

        s1_subject_points = dict(s1['subject_points'])
        s2_subject_points = dict(s2['subject_points'])

        # 补考通过：绩点改为 1
        for key in makeup_set:
            parts = key.split('||')
            if len(parts) == 2 and parts[0] == sid:
                course = parts[1]
                if course in s1_subject_points:
                    s1_subject_points[course] = 1
                if course in s2_subject_points:
                    s2_subject_points[course] = 1

        s1_total_point = sum(s1_subject_points.values())
        s2_total_point = sum(s2_subject_points.values())
        s1_quant_point = s1_qp.get(sid, 0)
        s2_quant_point = s2_qp.get(sid, 0)

        # 量化最终分（两学期平均）
        s1_quant_final = s1_qf.get(sid, 0)
        s2_quant_final = s2_qf.get(sid, 0)
        yearly_quant_final = (s1_quant_final + s2_quant_final) / 2

        s1_total_with_quant = s1_total_point + s1_quant_point
        s2_total_with_quant = s2_total_point + s2_quant_point

        yearly_total_point = s1_total_point + s2_total_point
        yearly_total_with_quant = s1_total_with_quant + s2_total_with_quant

        yearly_avg_no_quant = yearly_total_point / yearly_total_credits if yearly_total_credits > 0 else 0
        yearly_avg_with_quant = yearly_total_with_quant / yearly_total_credits_with_quant if yearly_total_credits_with_quant > 0 else 0

        yearly_results.append({
            'id': sid,
            'name': name,
            's1_subject_points': s1_subject_points,
            's2_subject_points': s2_subject_points,
            's1_scores': s1['scores'],
            's2_scores': s2['scores'],
            's1_quant_point': s1_quant_point,
            's2_quant_point': s2_quant_point,
            's1_total_point': s1_total_point,
            's2_total_point': s2_total_point,
            'yearly_total_point': yearly_total_point,
            'yearly_total_with_quant': yearly_total_with_quant,
            'yearly_avg_no_quant': yearly_avg_no_quant,
            'yearly_avg_with_quant': yearly_avg_with_quant,
            'yearly_quant_final': yearly_quant_final,
            'yearly_total_point_raw': yearly_total_point,
        })

    # 排名
    yearly_composite_scores = [r['yearly_avg_with_quant'] for r in yearly_results]
    yearly_study_scores = [r['yearly_avg_no_quant'] for r in yearly_results]

    composite_ranks = calc_rank_desc(yearly_composite_scores)
    study_ranks = calc_rank_desc(yearly_study_scores)

    n = len(yearly_results)
    for i, r in enumerate(yearly_results):
        r['composite_rank'] = composite_ranks[i]
        r['study_rank'] = study_ranks[i]
        r['composite_percent'] = calc_percentage(r['composite_rank'], n)
        r['study_percent'] = calc_percentage(r['study_rank'], n)

    return {
        'yearly_results': yearly_results,
        's1_subjects': [{'name': s['name'], 'credit': s['credit']} for s in semester1_data['subjects']],
        's2_subjects': [{'name': s['name'], 'credit': s['credit']} for s in semester2_data['subjects']],
        's1_total_credits': s1_total_credits,
        's2_total_credits': s2_total_credits,
        's1_total_credits_with_quant': s1_total_credits_with_quant,
        's2_total_credits_with_quant': s2_total_credits_with_quant,
        'yearly_total_credits': yearly_total_credits,
        'yearly_total_credits_with_quant': yearly_total_credits_with_quant,
        'makeup_passed': makeup_passed,
    }
