# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, jsonify, send_file, Response
import pandas as pd
import requests
import time
import os
import tempfile
import shutil
import threading
import uuid
import json
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 允许上传文件最大 64MB

# 任务状态存储: task_id -> state
# state:
# {
#   "status": "idle|running|done|error|canceled",
#   "percent": 0,
#   "eta_seconds": None,
#   "message": "",
#   "started_at": 0,
#   "finished_at": None,
#   "duration_seconds": None,
#   "download_filename": None,
#   "total": 0,
#   "current": 0,
#   "filename": "",
#   "cancel_requested": False,
#   "tmp_dir": ""
# }
TASKS = {}
TASKS_LOCK = threading.Lock()
TASK_RETENTION_SECONDS = 1800  # 任务状态保留 30 分钟后从内存移除（不删已生成文件）

def translate_text(text: str) -> str:
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {
            'client': 'gtx',
            'sl': 'auto',
            'tl': 'zh',
            'dt': 't',
            'q': text
        }
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if data and len(data) > 0 and len(data[0]) > 0:
            # 兼容多段返回
            parts = []
            for seg in data[0]:
                if isinstance(seg, list) and seg:
                    parts.append(seg[0])
            return ''.join(parts) if parts else "翻译失败"
        return "翻译失败"
    except Exception as e:
        print(f"[ERROR] translate_text: {e}")
        return "翻译失败"

def _safe_update(task_id: str, kv: dict):
    with TASKS_LOCK:
        st = TASKS.get(task_id)
        if st:
            st.update(kv)

def _get_state(task_id: str):
    with TASKS_LOCK:
        return dict(TASKS.get(task_id) or {})

def _schedule_state_cleanup(task_id: str):
    def _cleanup():
        time.sleep(TASK_RETENTION_SECONDS)
        with TASKS_LOCK:
            if task_id in TASKS:
                del TASKS[task_id]
    t = threading.Thread(target=_cleanup, daemon=True)
    t.start()

def _run_task(file_path: str, task_id: str):
    _safe_update(task_id, {"status": "running", "message": "读取文件...", "percent": 0, "started_at": time.time()})
    try:
        df = pd.read_excel(file_path)
        st = _get_state(task_id)
        if 'Title' not in df.columns:
            _safe_update(task_id, {"status": "error", "message": "表格中没有找到'Title'列", "finished_at": time.time()})
            _safe_update(task_id, {"duration_seconds": int((_get_state(task_id).get("finished_at") or time.time()) - (st.get("started_at") or time.time()))})
            _schedule_state_cleanup(task_id)
            return

        title_col_index = df.columns.get_loc('Title')
        insert_position = title_col_index + 1
        if '中文' not in df.columns:
            df.insert(insert_position, '中文', '')

        rows = []
        for i in range(len(df)):
            title_text = str(df.iloc[i, title_col_index])
            if title_text != 'nan' and title_text.strip():
                rows.append(i)

        total = len(rows)
        _safe_update(task_id, {"total": total, "current": 0, "message": "开始翻译..."})
        if total == 0:
            output_filename = os.path.splitext(os.path.basename(file_path))[0] + '_中文翻译.xlsx'
            output_path = os.path.join(tempfile.gettempdir(), output_filename)
            df.to_excel(output_path, index=False)
            _safe_update(task_id, {
                "status": "done",
                "percent": 100,
                "eta_seconds": 0,
                "message": "无需翻译，已完成",
                "download_filename": os.path.basename(output_path),
                "finished_at": time.time()
            })
            st2 = _get_state(task_id)
            _safe_update(task_id, {"duration_seconds": int((st2.get("finished_at") or time.time()) - (st2.get("started_at") or time.time()))})
            _schedule_state_cleanup(task_id)
            return

        start = _get_state(task_id).get("started_at") or time.time()
        for idx, i in enumerate(rows, start=1):
            st_loop = _get_state(task_id)
            if st_loop.get("cancel_requested"):
                _safe_update(task_id, {"status": "canceled", "message": f"已取消({idx-1}/{total})", "finished_at": time.time()})
                st3 = _get_state(task_id)
                _safe_update(task_id, {"duration_seconds": int((st3.get("finished_at") or time.time()) - (st3.get("started_at") or time.time()))})
                _schedule_state_cleanup(task_id)
                return

            title_text = str(df.iloc[i, title_col_index])
            translated = translate_text(title_text)
            df.iloc[i, insert_position] = translated

            elapsed = max(time.time() - start, 1e-6)
            done = idx
            remaining = max(total - done, 0)
            rate = done / elapsed
            eta = int(remaining / rate) if rate > 0 and remaining > 0 else 0
            percent = int(done * 100 / total)

            _safe_update(task_id, {
                "current": done,
                "percent": percent,
                "eta_seconds": eta,
                "message": f"翻译中({done}/{total})"
            })

            time.sleep(0.3)  # 适当降速，降低限流风险

        output_filename = os.path.splitext(os.path.basename(file_path))[0] + '_中文翻译.xlsx'
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        df.to_excel(output_path, index=False)

        _safe_update(task_id, {
            "status": "done",
            "percent": 100,
            "eta_seconds": 0,
            "message": f"翻译完成，共处理 {total} 行",
            "download_filename": os.path.basename(output_path),
            "finished_at": time.time()
        })
        st4 = _get_state(task_id)
        _safe_update(task_id, {"duration_seconds": int((st4.get("finished_at") or time.time()) - (st4.get("started_at") or time.time()))})
    except Exception as e:
        _safe_update(task_id, {"status": "error", "message": f"处理失败: {e}", "finished_at": time.time()})
        st5 = _get_state(task_id)
        _safe_update(task_id, {"duration_seconds": int((st5.get("finished_at") or time.time()) - (st5.get("started_at") or time.time()))})
    finally:
        try:
            base = os.path.dirname(file_path)
            if os.path.isdir(base):
                shutil.rmtree(base, ignore_errors=True)
        except Exception:
            pass
        _schedule_state_cleanup(task_id)

def start_background_task(file_path: str, filename: str) -> str:
    task_id = uuid.uuid4().hex[:12]
    with TASKS_LOCK:
        TASKS[task_id] = {
            "status": "idle",
            "percent": 0,
            "eta_seconds": None,
            "message": "",
            "started_at": None,
            "finished_at": None,
            "duration_seconds": None,
            "download_filename": None,
            "total": 0,
            "current": 0,
            "filename": filename,
            "cancel_requested": False,
            "tmp_dir": os.path.dirname(file_path)
        }
    t = threading.Thread(target=_run_task, args=(file_path, task_id), daemon=True)
    t.start()
    return task_id

def sse_progress(task_id: str):
    def event_stream():
        last_payload = None
        last_send = time.time()
        while True:
            state = _get_state(task_id)
            if not state:
                payload = json.dumps({"task_id": task_id, "status": "error", "message": "任务不存在"})
                yield f"data: {payload}\n\n"
                break

            payload_dict = {
                "task_id": task_id,
                "status": state.get("status"),
                "percent": state.get("percent"),
                "eta_seconds": state.get("eta_seconds"),
                "message": state.get("message"),
                "download_url": f"/download/{state.get('download_filename')}" if state.get("download_filename") else None,
                "filename": state.get("filename"),
                "started_at": state.get("started_at"),
                "finished_at": state.get("finished_at"),
                "duration_seconds": state.get("duration_seconds")
            }
            payload = json.dumps(payload_dict, ensure_ascii=False)
            now = time.time()

            if payload != last_payload:
                # event: progress\n 可选：让前端区分事件类型
                yield f"event: progress\ndata: {payload}\n\n"
                last_payload = payload
                last_send = now
            else:
                # 发送心跳，保持连接活跃（某些代理/网关会在无数据时断开）
                if now - last_send > 15:
                    yield f": heartbeat\n\n"  # 注释行作为 SSE 心跳
                    last_send = now

            if state.get("status") in ("done", "error", "canceled"):
                break
            time.sleep(0.2)
    return Response(event_stream(), mimetype="text/event-stream")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': '请上传Excel文件(.xlsx)'}), 400

    try:
        filename = secure_filename(file.filename)
        temp_dir = tempfile.mkdtemp()
        file_path = os.path.join(temp_dir, filename)
        file.save(file_path)
        print(f"[UPLOAD] saved to: {file_path}, size={os.path.getsize(file_path)} bytes")

        task_id = start_background_task(file_path, filename)
        _safe_update(task_id, {"started_at": time.time(), "status": "running", "message": "任务已启动"})

        return jsonify({
            'success': True,
            'message': '任务已启动',
            'task_id': task_id,
            'filename': filename
        }), 202

    except Exception as e:
        return jsonify({'error': f'处理失败: {str(e)}'}), 500

@app.route('/progress/<task_id>')
def progress(task_id):
    return sse_progress(task_id)

@app.route('/tasks/<task_id>', methods=['GET'])
def task_status(task_id):
    st = _get_state(task_id)
    if not st:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify({
        "task_id": task_id,
        "status": st.get("status"),
        "percent": st.get("percent"),
        "eta_seconds": st.get("eta_seconds"),
        "message": st.get("message"),
        "download_url": f"/download/{st.get('download_filename')}" if st.get("download_filename") else None,
        "filename": st.get("filename"),
        "started_at": st.get("started_at"),
        "finished_at": st.get("finished_at"),
        "duration_seconds": st.get("duration_seconds")
    })

@app.route('/tasks/<task_id>/cancel', methods=['POST'])
def task_cancel(task_id):
    st = _get_state(task_id)
    if not st:
        return jsonify({'error': '任务不存在'}), 404
    if st.get("status") in ("done", "error", "canceled"):
        return jsonify({'success': True, 'message': '任务已结束'}), 200
    _safe_update(task_id, {"cancel_requested": True, "message": "取消中..."})
    return jsonify({'success': True, 'message': '已请求取消'}), 202

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(tempfile.gettempdir(), filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        return f"文件不存在: {filename}", 404
    except Exception as e:
        return f"文件下载失败: {str(e)}", 500

if __name__ == '__main__':
    # 启动后浏览器中访问 http://localhost:5000
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)

if __name__ == '__main__':
    app.run(debug=True)