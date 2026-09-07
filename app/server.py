"""Flask 后端服务器"""
import os
import sys
import tempfile
import json
import threading
import time
from flask import Flask, request, jsonify, send_from_directory

# 添加当前目录到 path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_reader import read_raw_scores, read_eval_subjects, read_quantization
from calculator import process_all
from excel_writer import write_result

# ---------- 资源路径（兼容开发模式与 PyInstaller 打包模式） ----------
if hasattr(sys, '_MEIPASS'):
    # 打包后：资源被解压到 _MEIPASS
    _STATIC_DIR = os.path.join(sys._MEIPASS, 'static')
    _USAGE_IMG_DIR = os.path.join(sys._MEIPASS, '使用说明', 'images')
else:
    # 开发模式：基于本文件位置
    _HERE = os.path.dirname(os.path.abspath(__file__))
    _STATIC_DIR = os.path.join(_HERE, 'static')
    _PROJECT_ROOT = os.path.dirname(_HERE)
    _USAGE_IMG_DIR = os.path.join(_PROJECT_ROOT, '使用说明', 'images')

app = Flask(__name__, static_folder=_STATIC_DIR, static_url_path='')

# 临时存储上传的文件路径
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'zongce_uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------- 心跳：浏览器关闭后自动退出进程 ----------
_last_ping = time.time()
_ping_lock = threading.Lock()
_HEARTBEAT_TIMEOUT = 120  # 秒：超过此时间无心跳则退出


@app.before_request
def _touch_heartbeat():
    """任何请求都刷新心跳，确保用户操作期间服务不会退出"""
    global _last_ping
    with _ping_lock:
        _last_ping = time.time()


@app.route('/api/ping')
def api_ping():
    return jsonify({'ok': True})


def _heartbeat_watcher():
    """后台线程：浏览器关闭（无心跳）后自动结束进程"""
    while True:
        time.sleep(5)
        with _ping_lock:
            idle = time.time() - _last_ping
        if idle > _HEARTBEAT_TIMEOUT:
            os._exit(0)


def start_heartbeat():
    threading.Thread(target=_heartbeat_watcher, daemon=True).start()


@app.route('/')
def index():
    return send_from_directory(_STATIC_DIR, 'index.html')


@app.route('/usage_img/<path:filename>')
def usage_img(filename):
    """提供「使用说明/images」下的图片"""
    return send_from_directory(_USAGE_IMG_DIR, filename)


@app.route('/api/upload', methods=['POST'])
def upload():
    """上传 xlsx 文件，解析返回科目列表"""
    f = request.files.get('file')
    if not f:
        return jsonify({'error': '未收到文件'}), 400
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请上传 xlsx 文件'}), 400

    filepath = os.path.join(UPLOAD_DIR, f.filename)
    f.save(filepath)

    try:
        raw = read_raw_scores(filepath)
        eval_sub = read_eval_subjects(filepath)
        quant = read_quantization(filepath)

        # 所有科目列表（供用户勾选）
        all_subjects = []
        for c in raw['courses']:
            all_subjects.append({
                'name': c['name'],
                'code': c.get('code', ''),
                'category': c.get('category', ''),
                'credit': c.get('credit', 0),
            })

        # 参评科目（如果有）
        eval_subject_names = []
        if eval_sub:
            eval_subject_names = [s['name'] for s in eval_sub['subjects']]

        # 量化月份
        months = quant['months'] if quant else []

        return jsonify({
            'filename': f.filename,
            'filepath': filepath,
            'title': raw['title'],
            'student_count': len(raw['students']),
            'all_subjects': all_subjects,
            'eval_subject_names': eval_subject_names,
            'has_eval_sheet': eval_sub is not None,
            'has_quant_sheet': quant is not None,
            'quant_months': months,
            'quant_student_count': len(quant['students']) if quant else 0,
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/preview_scores', methods=['POST'])
def preview_scores():
    """预览选中科目的成绩表（学生×科目矩阵）"""
    data = request.json
    filepath = data.get('filepath')
    subject_names = data.get('subject_names', [])
    try:
        raw = read_raw_scores(filepath)
        # 找到选中科目的列索引
        selected_cols = []
        for c in raw['courses']:
            if c['name'] in subject_names:
                selected_cols.append(c)
        # 构造成绩矩阵
        rows = []
        for s in raw['students']:
            row = {
                'id': s['id'],
                'name': s['name'],
                'scores': {}
            }
            for c in selected_cols:
                row['scores'][c['name']] = s['scores'].get(c['col_idx'], None)
            rows.append(row)
        return jsonify({
            'subjects': [c['name'] for c in selected_cols],
            'students': rows,
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/preview_quant', methods=['POST'])
def preview_quant():
    """预览量化数据"""
    data = request.json
    filepath = data.get('filepath')
    try:
        quant = read_quantization(filepath)
        if not quant:
            return jsonify({'error': '未找到量化表'}), 400

        students = []
        for s in quant['students']:
            from calculator import calc_second_class_final, calc_base_management, calc_quant_final
            sc_raw = s['second_class_raw']
            sc_final = calc_second_class_final(sc_raw)
            # 优先用源表"总分"列（扣分后值），否则由扣分格计算
            base_src = s.get('base_mgmt_source')
            base = base_src if base_src is not None else calc_base_management(s['deductions'])
            qf = calc_quant_final(sc_raw, base)
            students.append({
                'id': s['id'],
                'name': s['name'],
                'second_class_raw': sc_raw,
                'second_class_final': sc_final,
                'deductions': s['deductions'],
                'base_management': base,
                'quant_final': round(qf, 2),
            })

        return jsonify({
            'students': students,
            'months': quant['months'],
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/calc', methods=['POST'])
def calc():
    """计算并返回结果预览"""
    data = request.json
    filepath = data.get('filepath')
    selected_subjects = data.get('selected_subjects', [])

    try:
        raw = read_raw_scores(filepath)
        eval_sub = read_eval_subjects(filepath)
        quant = read_quantization(filepath)

        # 构造 selected_subjects 格式
        subs = [{'name': s['name'], 'credit': s['credit']} for s in selected_subjects]

        result = process_all(raw, eval_sub, quant, subs)

        # 转成可 JSON 序列化的格式
        preview = {
            'quant': [{
                'id': q['id'],
                'name': q['name'],
                'second_class_raw': q['second_class_raw'],
                'second_class_final': q['second_class_final'],
                'base_management': round(q['base_management'], 2),
                'quant_final': round(q['quant_final'], 2),
                'quant_rank': q['quant_rank'],
            } for q in result['quant_results']],
            'grade': [{
                'id': g['id'],
                'name': g['name'],
                'subject_scores': g['subject_scores'],
                'subject_points': {k: round(v, 2) for k, v in g['subject_points'].items()},
                'quant_point': round(g['quant_point'], 2),
                'avg_no_quant': round(g['avg_no_quant'], 4),
                'avg_with_quant': round(g['avg_with_quant'], 4),
                'study_rank': g['study_rank'],
                'composite_rank': g['composite_rank'],
                'study_percent': round(g['study_percent'], 4),
                'composite_percent': round(g['composite_percent'], 4),
            } for g in result['grade_results']],
            'backup': [{
                'id': b['id'],
                'name': b['name'],
                'study_score': round(b['study_score'], 4),
                'study_rank': b['study_rank'],
                'study_percent': round(b['study_percent'], 4),
                'quant_score': round(b['quant_score'], 2),
                'quant_rank': b['quant_rank'],
                'quant_percent': round(b['quant_percent'], 4),
                'composite_score': round(b['composite_score'], 4),
                'composite_rank': b['composite_rank'],
                'composite_percent': round(b['composite_percent'], 4),
            } for b in result['backup']],
            'subjects': subs,
            'total_credits': result['total_credits'],
            'total_credits_with_quant': result['total_credits_with_quant'],
        }
        return jsonify(preview)
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/export', methods=['POST'])
def export():
    """导出结果 xlsx"""
    data = request.json
    filepath = data.get('filepath')
    selected_subjects = data.get('selected_subjects', [])
    class_name = (data.get('class_name') or '').strip()

    try:
        raw = read_raw_scores(filepath)
        eval_sub = read_eval_subjects(filepath)
        quant = read_quantization(filepath)

        subs = [{'name': s['name'], 'credit': s['credit']} for s in selected_subjects]
        result = process_all(raw, eval_sub, quant, subs)

        file_prefix = class_name if class_name else '综测'
        output_name = f'{file_prefix}_综测结果.xlsx'
        output_path = os.path.join(UPLOAD_DIR, output_name)
        write_result(output_path, result, class_name, source_path=filepath)

        return jsonify({
            'success': True,
            'output_path': output_path,
            'filename': output_name,
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/download')
def download():
    """下载导出的文件"""
    filename = request.args.get('filename', '')
    if not filename:
        return jsonify({'error': '文件名不能为空'}), 400
    return send_from_directory(UPLOAD_DIR, filename, as_attachment=True)


# ==================== 学年综测 API ====================

# 临时缓存学期数据（按会话）
_year_sessions = {}


@app.route('/api/year/upload', methods=['POST'])
def year_upload():
    """上传两个学期的综测结果文件，返回挂科清单"""
    files = request.files.getlist('files')
    if len(files) != 2:
        return jsonify({'error': '请上传两个学期综测文件'}), 400

    try:
        from year_calculator import read_semester, find_failed_courses

        semesters = []
        filenames = []
        for f in files:
            if not f.filename.endswith(('.xlsx', '.xls')):
                return jsonify({'error': '请上传 xlsx 文件'}), 400
            filepath = os.path.join(UPLOAD_DIR, f.filename)
            f.save(filepath)
            data = read_semester(filepath)
            semesters.append(data)
            filenames.append({'name': f.filename, 'filepath': filepath})

        # 合并两个学期的挂科清单
        failed1 = find_failed_courses(semesters[0])
        failed2 = find_failed_courses(semesters[1])

        # 生成会话 ID
        import uuid
        session_id = uuid.uuid4().hex[:12]
        _year_sessions[session_id] = {
            's1_data': semesters[0],
            's2_data': semesters[1],
            's1_filepath': filenames[0]['filepath'],
            's2_filepath': filenames[1]['filepath'],
        }

        # 标记挂科属于第一学期还是第二学期
        for f in failed1:
            f['semester'] = 1
        for f in failed2:
            f['semester'] = 2

        return jsonify({
            'session_id': session_id,
            's1_filename': filenames[0]['name'],
            's2_filename': filenames[1]['name'],
            's1_student_count': len(semesters[0]['students']),
            's2_student_count': len(semesters[1]['students']),
            's1_subjects': [s['name'] for s in semesters[0]['subjects']],
            's2_subjects': [s['name'] for s in semesters[1]['subjects']],
            'failed': failed1 + failed2,
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/year/calc', methods=['POST'])
def year_calc():
    """计算学年综测（含补考勾选）"""
    data = request.json
    session_id = data.get('session_id')
    makeup_passed = data.get('makeup_passed', [])  # ["学号||课程名", ...]

    session = _year_sessions.get(session_id)
    if not session:
        return jsonify({'error': '会话已过期，请重新上传'}), 400

    try:
        from year_calculator import calc_yearly

        result = calc_yearly(session['s1_data'], session['s2_data'], makeup_passed)

        # 转成可 JSON 序列化
        preview = {
            'students': [{
                'id': r['id'],
                'name': r['name'],
                'yearly_avg_no_quant': round(r['yearly_avg_no_quant'], 4),
                'yearly_avg_with_quant': round(r['yearly_avg_with_quant'], 4),
                'study_rank': r['study_rank'],
                'composite_rank': r['composite_rank'],
                'study_percent': round(r['study_percent'], 4),
                'composite_percent': round(r['composite_percent'], 4),
            } for r in sorted(result['yearly_results'], key=lambda x: -x['yearly_avg_with_quant'])],
            's1_subjects': [s['name'] for s in result['s1_subjects']],
            's2_subjects': [s['name'] for s in result['s2_subjects']],
            'yearly_total_credits': result['yearly_total_credits'],
            'yearly_total_credits_with_quant': result['yearly_total_credits_with_quant'],
        }

        # 缓存完整结果供导出
        session['year_result'] = result
        session['makeup_passed'] = makeup_passed

        return jsonify(preview)
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/year/export', methods=['POST'])
def year_export():
    """导出学年综测结果 xlsx"""
    data = request.json
    session_id = data.get('session_id')
    class_name = (data.get('class_name') or '').strip()

    session = _year_sessions.get(session_id)
    if not session or 'year_result' not in session:
        return jsonify({'error': '请先计算再导出'}), 400

    try:
        from year_excel_writer import write_year_result

        result = session['year_result']
        file_prefix = class_name if class_name else '综测'
        output_name = f'{file_prefix}_学年综测结果.xlsx'
        output_path = os.path.join(UPLOAD_DIR, output_name)
        write_year_result(output_path, result, class_name,
                          session['s1_filepath'], session['s2_filepath'])

        return jsonify({
            'success': True,
            'output_path': output_path,
            'filename': output_name,
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


if __name__ == '__main__':
    import socket
    import webbrowser

    is_packaged = hasattr(sys, '_MEIPASS')

    if is_packaged:
        # 打包模式：自动找空闲端口（5174 被占用时换端口）
        port = 5174
        try:
            s = socket.socket()
            s.bind(('127.0.0.1', port))
            s.close()
        except OSError:
            s = socket.socket()
            s.bind(('127.0.0.1', 0))
            port = s.getsockname()[1]
            s.close()
    else:
        port = 5174

    url = f'http://127.0.0.1:{port}'

    if is_packaged:
        # 打包模式：无命令行窗口，等待服务器就绪后自动打开浏览器 + 心跳自动退出
        def _open_browser():
            time.sleep(1.8)
            webbrowser.open(url)
        threading.Thread(target=_open_browser, daemon=True).start()
        start_heartbeat()  # 浏览器关闭后自动退出（仅打包模式）

    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False, threaded=True)
