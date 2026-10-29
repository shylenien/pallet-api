from fastapi import FastAPI, HTTPException, BackgroundTasks, WebSocket, WebSocketDisconnect, Query
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
import pandas as pd
from pathlib import Path
import threading
from typing import Optional, List

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "data.xlsx"
DOCS_FILE = BASE_DIR / "README.md"
excel_lock = threading.Lock()

try:
    api_description = DOCS_FILE.read_text(encoding="utf-8")
except FileNotFoundError:
    api_description = "Hệ thống Điều Phối Robot Kho Hàng & Excel Quantity API"

app = FastAPI(title="🤖 Hệ Thống Điều Phối Robot", description=api_description, version="8.0.0")


class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept(); self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections: self.active_connections.remove(websocket)

    async def broadcast(self, message: str):
        for connection in self.active_connections:
            try:
                await connection.send_text(message)
            except Exception:
                pass


manager = ConnectionManager()
DIST_COL, DIST_ROW, TOP_CORRIDOR_Y, BOTTOM_CORRIDOR_Y = 25, 10, 0, 70

STOPS_CONFIG = [
    {"stop_id": "S1", "x": 0, "y": 0, "lane": 1, "lane_name": "Làn 1", "targets": [(1, "L")]},
    {"stop_id": "S2", "x": 0, "y": 10, "lane": 1, "lane_name": "Làn 1", "targets": [(2, "L")]},
    {"stop_id": "S3", "x": 0, "y": 20, "lane": 1, "lane_name": "Làn 1", "targets": [(3, "L")]},
    {"stop_id": "S4", "x": 0, "y": 30, "lane": 1, "lane_name": "Làn 1", "targets": [(4, "L")]},
    {"stop_id": "S8", "x": 25, "y": 0, "lane": 2, "lane_name": "Làn 2", "targets": [(1, "R"), (9, "L")]},
    {"stop_id": "S7", "x": 25, "y": 10, "lane": 2, "lane_name": "Làn 2", "targets": [(2, "R"), (10, "L")]},
    {"stop_id": "S6", "x": 25, "y": 20, "lane": 2, "lane_name": "Làn 2", "targets": [(3, "R"), (11, "L")]},
    {"stop_id": "S5", "x": 25, "y": 30, "lane": 2, "lane_name": "Làn 2", "targets": [(4, "R"), (12, "L")]},
    {"stop_id": "S9", "x": 50, "y": 0, "lane": 3, "lane_name": "Làn 3", "targets": [(9, "R"), (17, "L")]},
    {"stop_id": "S10", "x": 50, "y": 10, "lane": 3, "lane_name": "Làn 3", "targets": [(10, "R"), (18, "L")]},
    {"stop_id": "S11", "x": 50, "y": 20, "lane": 3, "lane_name": "Làn 3", "targets": [(11, "R"), (19, "L")]},
    {"stop_id": "S12", "x": 50, "y": 30, "lane": 3, "lane_name": "Làn 3", "targets": [(12, "R"), (20, "L")]},
    {"stop_id": "S16", "x": 75, "y": 0, "lane": 4, "lane_name": "Làn 4", "targets": [(17, "R")]},
    {"stop_id": "S15", "x": 75, "y": 10, "lane": 4, "lane_name": "Làn 4", "targets": [(18, "R")]},
    {"stop_id": "S14", "x": 75, "y": 20, "lane": 4, "lane_name": "Làn 4", "targets": [(19, "R")]},
    {"stop_id": "S13", "x": 75, "y": 30, "lane": 4, "lane_name": "Làn 4", "targets": [(20, "R")]},
    {"stop_id": "S17", "x": 0, "y": 40, "lane": 1, "lane_name": "Làn 1", "targets": [(5, "L")]},
    {"stop_id": "S18", "x": 0, "y": 50, "lane": 1, "lane_name": "Làn 1", "targets": [(6, "L")]},
    {"stop_id": "S19", "x": 0, "y": 60, "lane": 1, "lane_name": "Làn 1", "targets": [(7, "L")]},
    {"stop_id": "S20", "x": 0, "y": 70, "lane": 1, "lane_name": "Làn 1", "targets": [(8, "L")]},
    {"stop_id": "S24", "x": 25, "y": 40, "lane": 2, "lane_name": "Làn 2", "targets": [(5, "R"), (13, "L")]},
    {"stop_id": "S23", "x": 25, "y": 50, "lane": 2, "lane_name": "Làn 2", "targets": [(6, "R"), (14, "L")]},
    {"stop_id": "S22", "x": 25, "y": 60, "lane": 2, "lane_name": "Làn 2", "targets": [(7, "R"), (15, "L")]},
    {"stop_id": "S21", "x": 25, "y": 70, "lane": 2, "lane_name": "Làn 2", "targets": [(8, "R"), (16, "L")]},
    {"stop_id": "S25", "x": 50, "y": 40, "lane": 3, "lane_name": "Làn 3", "targets": [(13, "R"), (21, "L")]},
    {"stop_id": "S26", "x": 50, "y": 50, "lane": 3, "lane_name": "Làn 3", "targets": [(14, "R"), (22, "L")]},
    {"stop_id": "S27", "x": 50, "y": 60, "lane": 3, "lane_name": "Làn 3", "targets": [(15, "R"), (23, "L")]},
    {"stop_id": "S28", "x": 50, "y": 70, "lane": 3, "lane_name": "Làn 3", "targets": [(16, "R"), (24, "L")]},
    {"stop_id": "S32", "x": 75, "y": 40, "lane": 4, "lane_name": "Làn 4", "targets": [(21, "R")]},
    {"stop_id": "S31", "x": 75, "y": 50, "lane": 4, "lane_name": "Làn 4", "targets": [(22, "R")]},
    {"stop_id": "S30", "x": 75, "y": 60, "lane": 4, "lane_name": "Làn 4", "targets": [(23, "R")]},
    {"stop_id": "S29", "x": 75, "y": 70, "lane": 4, "lane_name": "Làn 4", "targets": [(24, "R")]},
]

MODEL_CONFIG = {
    "1": {"max": 36, "bpl": 6},
    "2": {"max": 36, "bpl": 6},
    "3": {"max": 36, "bpl": 6},
}

global_route_data = {"output_string": "", "route_steps": [], "update_details": []}

# 🔥 BỘ BIẾN TRẠNG THÁI (STATE MACHINE) ĐIỀU PHỐI AMR <-> COBOT
execution_state = {
    "current_step_index": 0,  # Bước hiện tại trong route_steps
    "amr_status": "IDLE",  # IDLE | MOVING | ARRIVED | DONE_ALL
    "remaining_boxes": 0  # Số thùng còn phải thả tại trạm hiện tại
}


def get_model_cfg(model_name: str) -> dict:
    m = str(model_name).strip()
    return MODEL_CONFIG.get(m, {"max": 36, "bpl": 6})


def parse_slots_safely(val):
    if pd.isna(val) or not str(val).strip(): return set()
    parts = str(val).split(',')
    res = set()
    for p in parts:
        try:
            num = int(float(p.strip()))
            if num > 0: res.add(num)
        except ValueError:
            pass
    return res


class ModelInput(BaseModel):
    models_string: str
    start_x: int
    start_y: int
    save_to_db: bool = False  # Mặc định False để Cobot PUT từng thùng lên giao diện theo thời gian thực


class QuantityUpdate(BaseModel): new_quantity: int


class RecordUpdate(BaseModel):
    stt: int
    so_pallet: int
    model: str
    so_luong: int
    added_slot: Optional[int] = None


def clean_excel_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if 'STT Pallet' not in df.columns:
        if 'STT' in df.columns:
            df['STT Pallet'] = df['STT']
        elif 'Số Pallet' in df.columns:
            df['STT Pallet'] = df['Số Pallet']

    df = df.dropna(subset=['STT Pallet']).copy()
    for col in ['STT Pallet', 'Model', 'Số lượng', 'Số tầng', 'Số thùng/tầng', 'Tổng số thùng']:
        if col not in df.columns: df[col] = 0
    for col in ['STT Pallet', 'Số lượng']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    if 'Data_Slots' not in df.columns: df['Data_Slots'] = ""
    df['Data_Slots'] = df['Data_Slots'].fillna("").astype(str)

    if 'STT' in df.columns: df = df.drop(columns=['STT'])
    if 'Số Pallet' in df.columns: df = df.drop(columns=['Số Pallet'])
    return df


@app.get("/", response_class=HTMLResponse, tags=["Giao diện Web"])
def index():
    html_path = BASE_DIR / "index.html"
    if html_path.exists(): return FileResponse(html_path)
    return "<h1>File index.html không tồn tại</h1>"


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    icon_path = BASE_DIR / "favicon.ico"
    if icon_path.exists(): return FileResponse(icon_path)
    from fastapi.responses import Response
    return Response(content=b"", media_type="image/x-icon")


@app.get("/api/latest_route", tags=["Giao diện Web"])
def get_latest_route():
    return global_route_data


@app.get("/api/rows", tags=["Giao diện Web"])
def get_rows():
    with excel_lock:
        if not EXCEL_FILE.exists(): raise HTTPException(status_code=500, detail="Lỗi Excel")
        try:
            return clean_excel_dataframe(pd.read_excel(EXCEL_FILE)).to_dict(orient="records")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/pallets/status", tags=["REST API Ngoài"])
def get_pallets_status(stt: List[int] = Query(None, description="Danh sách STT Pallet")):
    with excel_lock:
        if not EXCEL_FILE.exists(): raise HTTPException(status_code=500)
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            if not stt: return []
            return df[df['STT Pallet'].isin(stt)].fillna("").to_dict(orient="records")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/process_models", tags=["Giao diện Web"])
def process_models(payload: ModelInput, background_tasks: BackgroundTasks):
    global global_route_data, execution_state

    if not payload.models_string: raise HTTPException(status_code=400, detail="Chuỗi rỗng")
    try:
        input_models = [int(x.strip()) for x in payload.models_string.replace('{', '').replace('}', '').split(';') if
                        x.strip()]
        input_models.reverse()
    except ValueError:
        raise HTTPException(status_code=400, detail="Sai định dạng đầu vào.")

    blocks = []
    for m in input_models:
        m_str = str(m)
        if not blocks or blocks[-1]['model'] != m_str:
            blocks.append({'model': m_str, 'count': 1})
        else:
            blocks[-1]['count'] += 1

    with excel_lock:
        if not EXCEL_FILE.exists(): raise HTTPException(status_code=500, detail="Lỗi Excel")
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            available_models = df['Model'].astype(str).str.strip().unique().tolist()
            invalid_models = [m for m in [str(x) for x in input_models] if m not in available_models]
            if invalid_models: raise HTTPException(status_code=400,
                                                   detail=f"Model không tồn tại: {set(invalid_models)}")

            virtual_counts = {row['STT Pallet']: row['Số lượng'] for _, row in df.iterrows()}
            log_pallet_counts = virtual_counts.copy()
            pallets_involved = set()

            virtual_states = {}
            for p in virtual_counts:
                mask = df['STT Pallet'] == p
                state_str = df[mask]['Data_Slots'].iloc[0]
                state = parse_slots_safely(state_str)
                if len(state) < virtual_counts[p]: state = set(range(1, virtual_counts[p] + 1))
                virtual_states[p] = state

            sequential_tasks = []
            for block in blocks:
                m_str, rem = block['model'], block['count']
                cfg = get_model_cfg(m_str)
                while rem > 0:
                    valid_p = df[df['Model'].astype(str).str.strip() == m_str]
                    best_p = next((p for p in sorted(valid_p['STT Pallet']) if virtual_counts[p] < cfg['max']), None)
                    if best_p is None: raise HTTPException(status_code=400, detail=f"Đã đầy kho chứa Model {m_str}!")
                    avail = cfg['max'] - virtual_counts[best_p]
                    drop_qty = min(rem, avail)
                    sequential_tasks.append({"model": m_str, "p_stt": best_p, "count": drop_qty})

                    pallets_involved.add(best_p)
                    virtual_counts[best_p] += drop_qty
                    rem -= drop_qty

            curr_x, curr_y = payload.start_x, payload.start_y

            def get_real_distance(x1, y1, x2, y2):
                if x1 == x2: return abs(y1 - y2)
                return min(abs(y1 - TOP_CORRIDOR_Y) + abs(x1 - x2) + abs(y2 - TOP_CORRIDOR_Y),
                           abs(y1 - BOTTOM_CORRIDOR_Y) + abs(x1 - x2) + abs(y2 - BOTTOM_CORRIDOR_Y))

            route_steps, output_tokens, update_log, step_counter = [], [], [], 0

            involved_list = sorted(list(pallets_involved))
            query_str = "&amp;".join([f"stt={p}" for p in involved_list])
            api_endpoint_html = f"GET /api/pallets/status?{query_str}"

            update_log.append(
                "<b style='color:#fbbf24; font-size:14px; margin-top: 8px; display: block;'>--- [SYSTEM PRE-CHECK] TRẠNG THÁI PALLET ĐÍCH TẠI HIỆN TRƯỜNG ---</b>")
            update_log.append(
                f"<span style='color:#94a3b8; font-size:12px; font-family:monospace; display: block; margin-bottom: 4px;'>📡 Server gọi API: <code style='color:#38bdf8'>{api_endpoint_html}</code> (Trích xuất cục bộ mục tiêu)</span>")

            for idx, p_id in enumerate(involved_list):
                curr_q = log_pallet_counts[p_id]
                model_n = df.loc[df['STT Pallet'] == p_id, 'Model'].iloc[0]

                layer = (curr_q // 6) + 1 if curr_q < 36 else 6
                boxes_in_layer = curr_q % 6
                if curr_q > 0 and curr_q % 6 == 0:
                    boxes_in_layer = 6
                    layer = curr_q // 6

                slots_left = 6 - boxes_in_layer

                if curr_q == 0:
                    msg = "Trống hoàn toàn - Bắt đầu xếp Tầng 1 ➔ Còn trống 6 lỗ."
                elif boxes_in_layer == 6:
                    msg = f"Tầng {layer} đã đầy (6/6) - Chuẩn bị sang Tầng {layer + 1} ➔ Còn trống 6 lỗ."
                else:
                    msg = f"Đang xếp Tầng {layer} - Đã có {boxes_in_layer}/6 thùng ➔ Còn trống {slots_left} lỗ."

                prefix = "└─" if idx == len(involved_list) - 1 else "├─"
                update_log.append(
                    f"<span style='color:#fcd34d; font-size:13px; font-family:monospace; display: block;'> {prefix} 🔍 Pallet {p_id} (Model {model_n}): {msg}</span>")
            update_log.append(
                "<b style='color:#fbbf24; margin-bottom: 8px; display: block;'>-----------------------------------------------------------------</b>")

            for task in sequential_tasks:
                p_stt, rem_drop, model = task["p_stt"], task["count"], task["model"]
                cfg = get_model_cfg(model)
                bpl = cfg['bpl']
                max_layers = cfg['max'] // bpl + (1 if cfg['max'] % bpl else 0)

                while rem_drop > 0:
                    state = virtual_states[p_stt]
                    active_layer = -1
                    for L in range(max_layers):
                        if len([s for s in state if L * bpl + 1 <= s <= L * bpl + bpl]) < bpl:
                            active_layer = L;
                            break
                    if active_layer == -1: break

                    offset = active_layer * bpl
                    L_slots = [offset + i for i in range(1, bpl + 1, 2)]
                    R_slots = [offset + i for i in range(2, bpl + 1, 2)]

                    avail_L = sorted([s for s in L_slots if s not in state], reverse=True)
                    avail_R = sorted([s for s in R_slots if s not in state], reverse=True)
                    if not avail_L and not avail_R: break

                    valid_stops = []
                    for stop in STOPS_CONFIG:
                        for tgt_p, side in stop["targets"]:
                            if tgt_p == p_stt:
                                avail_side = avail_L if side == 'L' else avail_R
                                if avail_side:
                                    valid_stops.append({"stop": stop, "side": side, "avail": avail_side,
                                                        "dist": get_real_distance(curr_x, curr_y, stop["x"],
                                                                                  stop["y"])})

                    if not valid_stops: break
                    valid_stops.sort(key=lambda s: s["dist"])
                    best_stop = valid_stops[0]

                    drop_count = min(rem_drop, len(best_stop["avail"]))
                    slots_to_fill = best_stop["avail"][:drop_count]
                    state.update(slots_to_fill)
                    rem_drop -= drop_count

                    stop_cfg, side = best_stop["stop"], best_stop["side"]
                    curr_x, curr_y = stop_cfg["x"], stop_cfg["y"]
                    step_counter += 1
                    token_str = f"{p_stt}{side}"
                    output_tokens.append(token_str)
                    slots_relative = [s - offset for s in slots_to_fill]
                    tang_logical = active_layer + 1
                    slot_desc = [f"T{tang_logical}-Ô{s}" for s in slots_relative]

                    route_steps.append({
                        "step_num": step_counter, "stop_id": stop_cfg["stop_id"], "lane": stop_cfg["lane"],
                        "lane_name": stop_cfg["lane_name"], "token": token_str, "is_dual": False,
                        "actions": [{"pallet_stt": p_stt, "side": side, "side_label": f"{p_stt}{side}", "model": model,
                                     "count": drop_count, "slots": slot_desc, "relative_slots": slots_relative}]
                    })

            merged_route = []
            for step in route_steps:
                if not merged_route:
                    merged_route.append(step)
                else:
                    prev_step = merged_route[-1]
                    if (step["stop_id"] == prev_step["stop_id"] and step["token"] == prev_step["token"] and
                            step["actions"][0]["model"] == prev_step["actions"][0]["model"]):
                        existing_act = prev_step["actions"][0];
                        new_act = step["actions"][0]
                        existing_act["count"] += new_act["count"]
                        existing_act["slots"].extend(new_act["slots"])
                        existing_act["relative_slots"].extend(new_act["relative_slots"])
                    else:
                        merged_route.append(step)

            final_route_steps, final_output_tokens = [], []
            for idx, step in enumerate(merged_route):
                step_num = idx + 1;
                step["step_num"] = step_num
                final_output_tokens.append(step["token"])
                act = step["actions"][0]
                side_vn = "Trái" if act["side"] == "L" else "Phải"
                p_stt = act['pallet_stt']

                update_log.append(
                    f"<b style='color:#4ade80;'>Bước {step_num} [{step['lane_name']}] - Trạm {act['side_label']} ({side_vn}): Xếp {act['count']} thùng [{', '.join(act['slots'])}]</b>")
                update_log.append(
                    f"<span style='color:#cbd5e1; font-size:12px; font-family:monospace; display: block;'>&nbsp;&nbsp;├─ 🚐 AMR lấy lệnh &amp; báo đỗ: <code style='color:#fbbf24'>GET /api/amr/next_step</code> ➔ <code style='color:#fbbf24'>POST /api/amr/arrived</code> (Tại {act['side_label']})</span>")

                for i in range(act["count"]):
                    slot_name = act["slots"][i]
                    log_pallet_counts[p_stt] += 1
                    current_qty = log_pallet_counts[p_stt]

                    current_layer = (current_qty - 1) // 6 + 1 if current_qty > 0 else 0
                    boxes_in_layer = (current_qty - 1) % 6 + 1 if current_qty > 0 else 0

                    update_log.append(
                        f"<span style='color:#94a3b8; font-size:12px; font-family:monospace; display: block;'>&nbsp;&nbsp;├─ 🤖 Cobot nhận việc (Lần {i + 1}): <code style='color:#38bdf8'>GET /api/cobot/current_task</code> (Trả về {act['side_label']})</span>")

                    prefix = "└─" if i == act["count"] - 1 else "├─"
                    update_log.append(
                        f"<span style='color:#94a3b8; font-size:12px; font-family:monospace; display: block;'>&nbsp;&nbsp;{prefix} 📝 Báo cáo: <code style='color:#f472b6'>POST /records/quantity</code> (Nhét vào {slot_name}) ➔ <b style='color:#fcd34d'>Pallet {p_stt} đạt Tầng {current_layer} ({boxes_in_layer}/6) - Tổng: {current_qty}/36</b></span>")

                final_route_steps.append(step)

            global_route_data.update({
                "output_string": "{" + "; ".join(final_output_tokens) + "}",
                "route_steps": final_route_steps,
                "update_details": update_log
            })

            # 🔥 KHỞI TẠO TRẠNG THÁI CHO AMR & COBOT BẮT ĐẦU CHẠY BƯỚC 1
            if final_route_steps:
                execution_state["current_step_index"] = 0
                execution_state["amr_status"] = "IDLE"
                execution_state["remaining_boxes"] = final_route_steps[0]["actions"][0]["count"]
            else:
                execution_state["amr_status"] = "DONE_ALL"
                execution_state["remaining_boxes"] = 0

            if payload.save_to_db:
                for p_stt, state in virtual_states.items():
                    mask = df['STT Pallet'] == p_stt
                    row_model = str(df.loc[mask, 'Model'].iloc[0]).strip()
                    cfg = get_model_cfg(row_model)
                    new_qty = len(state)
                    df.loc[mask, 'Số lượng'] = new_qty
                    df.loc[mask, 'Tổng số thùng'] = new_qty
                    df.loc[mask, 'Số tầng'] = 0 if new_qty == 0 else (new_qty - 1) // cfg['bpl'] + 1
                    df.loc[mask, 'Số thùng/tầng'] = 0 if new_qty == 0 else (new_qty % cfg['bpl'] or cfg['bpl'])
                    df.loc[mask, 'Data_Slots'] = ",".join(map(str, sorted(list(state))))
                try:
                    df.to_excel(EXCEL_FILE, index=False)
                except PermissionError:
                    raise HTTPException(status_code=400, detail="Vui lòng tắt Excel.")

            background_tasks.add_task(manager.broadcast, "update")
            return {"success": True, "output_string": global_route_data["output_string"],
                    "route_steps": final_route_steps, "update_details": update_log}
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


# =========================================================================
# 🔥 CÁC API GIAO TIẾP (HANDSHAKE) GIỮA AMR VÀ COBOT
# =========================================================================

@app.get("/api/amr/next_step", tags=["Giao tiếp AMR & Cobot"])
def amr_get_next_step():
    steps = global_route_data.get("route_steps", [])
    idx = execution_state["current_step_index"]

    if not steps or idx >= len(steps) or execution_state["amr_status"] == "DONE_ALL":
        execution_state["amr_status"] = "DONE_ALL"
        return {"status": "idle", "message": "Không có lệnh hoặc đã hoàn thành toàn bộ lộ trình."}

    if execution_state["amr_status"] == "ARRIVED" and execution_state["remaining_boxes"] > 0:
        return {"status": "wait_cobot",
                "message": f"Đang chờ Cobot thả nốt {execution_state['remaining_boxes']} thùng."}

    current_step = steps[idx]
    act = current_step["actions"][0]
    execution_state["amr_status"] = "MOVING"

    return {
        "status": "move",
        "step_num": current_step["step_num"],
        "stop_id": current_step["stop_id"],
        "lane_name": current_step["lane_name"],
        "token": current_step["token"],
        "pallet_stt": act["pallet_stt"],
        "side": act["side"],
        "boxes_to_drop": execution_state["remaining_boxes"]
    }


@app.post("/api/amr/arrived", tags=["Giao tiếp AMR & Cobot"])
def amr_report_arrived():
    steps = global_route_data.get("route_steps", [])
    idx = execution_state["current_step_index"]

    if not steps or idx >= len(steps):
        raise HTTPException(status_code=400, detail="Không có lộ trình nào đang chạy.")

    execution_state["amr_status"] = "ARRIVED"
    current_step = steps[idx]
    return {
        "success": True,
        "token": current_step["token"],
        "message": f"Server đã ghi nhận AMR đỗ tại {current_step['token']}. Đã mở khóa cho Cobot!"
    }


@app.get("/api/cobot/current_task", tags=["Giao tiếp AMR & Cobot"])
def cobot_get_current_task():
    steps = global_route_data.get("route_steps", [])
    idx = execution_state["current_step_index"]

    if not steps or idx >= len(steps) or execution_state["amr_status"] == "DONE_ALL":
        return {"status": "idle", "message": "Hệ thống đang nghỉ, không có đơn hàng."}

    current_step = steps[idx]
    if execution_state["amr_status"] != "ARRIVED":
        return {
            "status": "waiting_amr",
            "target_token": current_step["token"],
            "message": f"Đang chờ AMR di chuyển tới trạm {current_step['token']}..."
        }

    act = current_step["actions"][0]
    task_info = get_cobot_task(pallet=act["pallet_stt"], side=act["side"])
    task_info["remaining_boxes_at_stop"] = execution_state["remaining_boxes"]
    return task_info


# =========================================================================

@app.delete("/api/reset_all", tags=["Giao diện Web"])
def reset_all(background_tasks: BackgroundTasks):
    with excel_lock:
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            for col in ['Số lượng', 'Tổng số thùng', 'Số tầng', 'Số thùng/tầng']: df[col] = 0
            df['Data_Slots'] = ""
            df.to_excel(EXCEL_FILE, index=False)
            global_route_data.update({"output_string": "", "route_steps": [], "update_details": []})
            execution_state.update({"current_step_index": 0, "amr_status": "IDLE", "remaining_boxes": 0})
            background_tasks.add_task(manager.broadcast, "update")
            return {"success": True}
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


@app.put("/api/update_row_quantity/{stt}", tags=["Giao diện Web"])
def update_row_quantity(stt: int, payload: QuantityUpdate, background_tasks: BackgroundTasks):
    qty = payload.new_quantity
    with excel_lock:
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            mask = df['STT Pallet'] == stt
            row_model = str(df.loc[mask, 'Model'].iloc[0]).strip()
            cfg = get_model_cfg(row_model)
            df.loc[mask, 'Số lượng'] = df.loc[mask, 'Tổng số thùng'] = qty
            df.loc[mask, 'Số tầng'] = 0 if qty == 0 else (qty - 1) // cfg['bpl'] + 1
            df.loc[mask, 'Số thùng/tầng'] = 0 if qty == 0 else (qty % cfg['bpl'] or cfg['bpl'])
            df.loc[mask, 'Data_Slots'] = ",".join(map(str, range(1, qty + 1))) if qty > 0 else ""
            df.to_excel(EXCEL_FILE, index=False)
            background_tasks.add_task(manager.broadcast, "update")
            return {"success": True}
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


@app.delete("/api/delete_row/{stt}", tags=["Giao diện Web"])
def delete_row(stt: int, background_tasks: BackgroundTasks):
    return update_row_quantity(stt, QuantityUpdate(new_quantity=0), background_tasks)


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True: await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)


@app.get("/health", tags=["REST API Ngoài"])
def health_check(): return {"status": "ok"}


@app.get("/records", tags=["REST API Ngoài"])
def get_records(model: Optional[str] = Query(None), so_pallet: Optional[int] = Query(None)):
    df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
    if model: df = df[df['Model'].astype(str).str.strip() == model.strip()]
    if so_pallet is not None: df = df[df['STT Pallet'] == so_pallet]
    return df.fillna("").to_dict(orient="records")


@app.get("/records/lookup", tags=["REST API Ngoài"])
def lookup_record(stt: int, so_pallet: int, model: str):
    df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
    mask = (df['STT Pallet'] == stt) & (df['Model'].astype(str).str.strip() == model.strip())
    result = df[mask]
    if result.empty: raise HTTPException(status_code=404, detail="Không tìm thấy")
    if len(result) > 1: raise HTTPException(status_code=409, detail="Trùng lặp")
    return result.fillna("").to_dict(orient="records")[0]


@app.post("/records/quantity", tags=["REST API Ngoài"])
def update_quantity(payload: RecordUpdate, background_tasks: BackgroundTasks):
    with excel_lock:
        if not EXCEL_FILE.exists(): raise HTTPException(status_code=500)
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            mask = (df['STT Pallet'] == payload.stt) & (df['Model'].astype(str).str.strip() == payload.model.strip())
            matching_indices = df[mask].index
            if len(matching_indices) == 0: raise HTTPException(status_code=404)

            target_index = matching_indices[0]
            new_qty = payload.so_luong
            cfg = get_model_cfg(payload.model)

            if payload.added_slot is not None:
                current_val = df.loc[target_index, 'Data_Slots']
                state = parse_slots_safely(current_val)
                state.add(payload.added_slot)
                df.loc[target_index, 'Data_Slots'] = ",".join(map(str, sorted(list(state))))
                df.loc[target_index, 'Số lượng'] = len(state)
                df.loc[target_index, 'Tổng số thùng'] = len(state)
            else:
                df.loc[target_index, 'Số lượng'] = new_qty
                df.loc[target_index, 'Tổng số thùng'] = new_qty
                df.loc[target_index, 'Data_Slots'] = ",".join(map(str, range(1, new_qty + 1))) if new_qty > 0 else ""

            final_qty = df.loc[target_index, 'Số lượng']
            df.loc[target_index, 'Số tầng'] = 0 if final_qty == 0 else (final_qty - 1) // cfg['bpl'] + 1
            df.loc[target_index, 'Số thùng/tầng'] = 0 if final_qty == 0 else (final_qty % cfg['bpl'] or cfg['bpl'])

            df.to_excel(EXCEL_FILE, index=False)

            # 🔥 CẬP NHẬT TRẠNG THÁI HANDSHAKE SAU KHI COBOT THẢ XONG 1 THÙNG
            if execution_state["amr_status"] == "ARRIVED" and execution_state["remaining_boxes"] > 0:
                execution_state["remaining_boxes"] -= 1
                if execution_state["remaining_boxes"] == 0:
                    execution_state["current_step_index"] += 1
                    steps = global_route_data.get("route_steps", [])
                    if execution_state["current_step_index"] < len(steps):
                        next_act = steps[execution_state["current_step_index"]]["actions"][0]
                        execution_state["remaining_boxes"] = next_act["count"]
                        execution_state["amr_status"] = "IDLE"  # Giải phóng cho AMR chạy sang trạm kế tiếp
                    else:
                        execution_state["amr_status"] = "DONE_ALL"

            background_tasks.add_task(manager.broadcast, "update")
            return {"success": True}
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/cobot/task", tags=["REST API Ngoài"])
def get_cobot_task(pallet: int = Query(...), side: str = Query(...)):
    side = side.strip().upper()
    if side not in ['L', 'R']: raise HTTPException(status_code=400, detail="Tham số 'side' sai.")
    with excel_lock:
        if not EXCEL_FILE.exists(): raise HTTPException(status_code=500, detail="Không tìm thấy Excel.")
        try:
            df = clean_excel_dataframe(pd.read_excel(EXCEL_FILE))
            mask = df['STT Pallet'] == pallet
            if not mask.any(): raise HTTPException(status_code=404, detail="Không tìm thấy Pallet.")
            row = df[mask].iloc[0]
            current_model = str(row['Model']).strip()
            cfg = get_model_cfg(current_model)

            state = parse_slots_safely(row['Data_Slots'])

            active_layer = -1
            max_layers = cfg['max'] // cfg['bpl'] + (1 if cfg['max'] % cfg['bpl'] else 0)
            for L in range(max_layers):
                if len([s for s in state if L * cfg['bpl'] + 1 <= s <= L * cfg['bpl'] + cfg['bpl']]) < cfg['bpl']:
                    active_layer = L;
                    break

            if active_layer == -1: return {"status": "full", "message": f"Pallet {pallet} đã đầy ({cfg['max']} thùng).",
                                           "data": None}
            tang_logical = active_layer + 1
            offset = active_layer * cfg['bpl']
            target_slots = [offset + i for i in range(1, cfg['bpl'] + 1, 2)] if side == 'L' else [offset + i for i in
                                                                                                  range(2,
                                                                                                        cfg['bpl'] + 1,
                                                                                                        2)]
            avail_slots_absolute = sorted([s for s in target_slots if s not in state], reverse=True)
            if not avail_slots_absolute: return {"status": "side_full",
                                                 "message": f"Mặt {side} của Tầng {tang_logical} đã được xếp đầy.",
                                                 "data": {"pallet_stt": pallet, "model": current_model, "side": side,
                                                          "current_layer": tang_logical, "empty_slots_relative": []}}

            avail_slots_relative = [s - offset for s in avail_slots_absolute]
            next_slot = avail_slots_relative[0]
            pose_id_str = f"P{pallet}-M{current_model}-T{tang_logical}-O{next_slot}-{side}"
            return {
                "status": "ready", "message": "Pallet sẵn sàng.",
                "data": {"pallet_stt": pallet, "model": current_model, "side": side, "current_layer": tang_logical,
                         "empty_slots_relative": avail_slots_relative, "next_target_slot": next_slot,
                         "absolute_slot": offset + next_slot, "pose_id": pose_id_str}
            }
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))