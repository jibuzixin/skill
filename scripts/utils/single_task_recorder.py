#!/usr/bin/env python3
import json
import sys
import queue
import threading
from datetime import datetime
from pathlib import Path
from mitmproxy import http

# ======================== 【异步日志线程】完全不阻塞主流程 ========================
log_queue = queue.Queue(maxsize=100000)


def log_worker():
    """后台异步写文件 + 异步解析流式数据，完全不影响抓包"""
    while True:
        item = log_queue.get()
        if item is None:
            break

        log_path, raw_data = item
        try:
            # 1. 流式响应：自动解析 SSE → 合并成标准大模型返回格式
            if raw_data.get("is_stream"):
                raw_data["response"]["body"] = parse_stream_to_normal(raw_data["response"]["body"])

            # 2. 安全写入文件
            Path(log_path).parent.mkdir(parents=True, exist_ok=True)
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(raw_data, ensure_ascii=False, ) + "\n")
        except Exception:
            pass


threading.Thread(target=log_worker, daemon=True).start()


# ======================== 流式 SSE → 标准结构解析（异步执行） ========================
def parse_stream_to_normal(sse_text):
    """
    完整版流式 SSE 解析：
    支持 回答内容 + 思考内容 + 工具调用 + usage + id/model
    输出结构 = 标准非流式大模型返回格式
    """
    lines = sse_text.strip().split('\n')
    full_content = ""
    full_reasoning = ""
    usage = None
    first_chunk = None
    tool_calls_dict = {}
    sse_data = []

    for line in lines:
        line = line.strip()
        if not line.startswith('data: '):
            continue

        data_str = line[6:].strip()
        if data_str == "[DONE]":
            break

        try:
            chunk = json.loads(data_str)
            sse_data.append(chunk)

            # 记录第一个 chunk（取 id / model）
            if first_chunk is None and chunk.get('id'):
                first_chunk = chunk

            # 遍历 choices
            if chunk.get('choices'):
                for choice in chunk['choices']:
                    delta = choice.get('delta', {})

                    # 拼接回答内容
                    full_content += delta.get('content', '')

                    # 拼接思考内容
                    full_reasoning += delta.get('reasoning', delta.get('reasoning_content', ''))

                    # 拼接工具调用（支持分片 arguments）
                    for tc in delta.get('tool_calls', []):
                        idx = tc.get('index', 0)
                        if idx not in tool_calls_dict:
                            tool_calls_dict[idx] = {
                                'id': tc.get('id', ''),
                                'type': tc.get('type', 'function'),
                                'function': {
                                    'name': tc.get('function', {}).get('name', ''),
                                    'arguments': ''
                                }
                            }
                        # 增量更新字段
                        if tc.get('id'):
                            tool_calls_dict[idx]['id'] = tc['id']
                        if tc.get('type'):
                            tool_calls_dict[idx]['type'] = tc['type']
                        if tc.get('function', {}).get('name'):
                            tool_calls_dict[idx]['function']['name'] = tc['function']['name']
                        # 拼接参数
                        tool_calls_dict[idx]['function']['arguments'] += tc.get('function', {}).get('arguments', '')

            # 提取 usage（有些模型在最后一个chunk返回）
            if chunk.get('usage'):
                usage = chunk['usage']

        except Exception:
            continue

    # 如果没有任何有效数据
    if not first_chunk:
        return {
            "error": "parse failed",
            "content": full_content,
            "reasoning_content": full_reasoning
        }

    # 构建标准非流式结构
    message = {"role": "assistant", "content": []}
    if full_reasoning.strip():
        message["content"].append({"type": "thinking", "thinking": full_reasoning.strip()})
    if full_content.strip():
        message["content"].append({"type": "text", "text": full_content})
    if tool_calls_dict:
        for i in sorted(tool_calls_dict.keys()):
            message["content"].append({
                "type": tool_calls_dict[i]["type"],
                "id": tool_calls_dict[i]["id"],
                "function": {
                    "name": tool_calls_dict[i]["function"]["name"],
                    "arguments": json.loads(tool_calls_dict[i]["function"]["arguments"]),
                }
            })

    finish_reason = "tool_calls" if tool_calls_dict else "stop"

    return {
        "id": first_chunk.get("id"),
        "model": first_chunk.get("model"),
        "choices": [
            {
                "message": message,
                "finish_reason": finish_reason
            }
        ],
        "usage": usage,
        "stream": True,
        "chunks_count": len(sse_data)
    }


# ======================== 端口 / 日志路径读取 ========================
def get_mitm_port():
    try:
        idx = sys.argv.index("-p")
        return sys.argv[idx + 1]
    except:
        return "8083"


LAST_LOG_PATH = None
THIS_DIR = Path(__file__).parent


def get_log_path():
    global LAST_LOG_PATH
    port = get_mitm_port()
    sync_file = THIS_DIR / f"mitm_log_path_{port}.tmp"
    try:
        if sync_file.exists():
            p = sync_file.read_text(encoding="utf-8").strip()
            if p:
                LAST_LOG_PATH = p
                return p
    except:
        pass
    return LAST_LOG_PATH


# ======================== 【核心】全局请求隔离：绝对不串流 ========================
REQUEST_CTX = {}  # { flow.id: { "log_path": str } }


# ======================== 抓包主类 ========================
class LLMCapture:
    def request(self, flow: http.HTTPFlow):
        if "/chat/completions" not in flow.request.path:
            return
        # 每个请求独立上下文，永不串
        REQUEST_CTX[flow.id] = {"log_path": get_log_path()}

    def response(self, flow: http.HTTPFlow):
        # 取出并删除当前请求上下文
        ctx = REQUEST_CTX.pop(flow.id, None)
        if not ctx or not ctx["log_path"]:
            return

        req = flow.request
        resp = flow.response

        # 构建日志（只做最轻量操作，不阻塞）
        log_data = {
            "timestamp": datetime.now().isoformat(),
            "flow_id": flow.id,
            "url": req.url,
            "method": req.method,
            "request": {
                "headers": dict(req.headers),
                "body": self.safe_request_body(req.content, ctx['log_path']),
            },
            "response": {
                "status": resp.status_code,
                "headers": dict(resp.headers),
                "body": resp.text,  # mitm 已自动拼完所有流式chunk
            },
            "is_stream": self.is_stream(flow),
        }

        # 【关键】扔进队列就不管，异步解析 + 异步写文件
        log_queue.put((ctx["log_path"], log_data))

    # ======================== 工具 ========================
    def is_stream(self, flow):
        ct = flow.response.headers.get("content-type", "")
        if "text/event-stream" in ct:
            return True
        try:
            return flow.request.json().get("stream") is True
        except:
            return False

    def safe_request_body(self, b, log_path: str):
        try:
            content = json.loads(b)
            if Path(log_path).exists():
                content['messages'] = content['messages'][2:][-4:]
            return content
        except:
            return b.decode("utf-8", "ignore")[:3000]


addons = [LLMCapture()]