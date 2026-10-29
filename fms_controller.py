import requests
import time
import threading
import os
import sys

API_URL = "http://127.0.0.1:8000"


def print_amr(msg): print(f"\033[93m🚐 [AMR API LOG]\033[0m {msg}")


def print_cobot(msg): print(f"\033[96m🤖 [COBOT API LOG]\033[0m {msg}")


def print_sys(msg): print(f"\033[92m⚙️ [SYSTEM]\033[0m {msg}")


def print_err(msg): print(f"\033[91m❌ [LỖI]\033[0m {msg}")


def check_server_ready():
    try:
        res = requests.get(f"{API_URL}/api/amr/next_step", timeout=3)
        if res.status_code != 200:
            print_err(f"Server chưa sẵn sàng, mã lỗi: {res.status_code}")
            return False
        print_sys("Kết nối Server app.py thành công! Bắt đầu giám sát giao tiếp...")
        return True
    except requests.exceptions.ConnectionError:
        print_err(f"Không thể kết nối tới {API_URL}. Bạn chưa chạy app.py?")
        return False


# =========================================================================
# LUỒNG 1: MÔ PHỎNG CON XE AMR GIAO TIẾP VỚI SERVER
# =========================================================================
def amr_client_simulation():
    last_status = None
    while True:
        try:
            # 1. AMR GỌI LỆNH XEM MÌNH PHẢI ĐI ĐÂU
            # Request: GET /api/amr/next_step
            res = requests.get(f"{API_URL}/api/amr/next_step")
            data = res.json()
            status = data.get("status")

            if status != last_status:
                print_amr(f"Gửi GET /api/amr/next_step ➔ Server trả về: {data}")
                last_status = status

            if status == "move":
                token = data["token"]
                print("\n" + "═" * 70)
                print_amr(f"🚀 [HÀNH ĐỘNG] AMR nhận lệnh đến trạm [{token}]. Đang cho xe lăn bánh...")

                # Làm chậm lại 4 giây để bạn kịp nhìn rõ trạng thái
                time.sleep(4.0)

                # 2. AMR TỚI NƠI VÀ GỌI LỆNH BÁO CÁO ĐÃ ĐỖ
                # Request: POST /api/amr/arrived
                print_amr(f"🏁 [HÀNH ĐỘNG] AMR đã đỗ cứng tại [{token}]. Đang gọi: POST /api/amr/arrived")
                res_arr = requests.post(f"{API_URL}/api/amr/arrived")
                print_amr(f"   ↳ Server xác nhận: {res_arr.json()}")
                print("═" * 70 + "\n")
                last_status = "ARRIVED_SENT"

            elif status == "wait_cobot":
                # AMR đang đứng chờ Cobot gắp hàng xong
                time.sleep(1.0)

            elif status == "idle":
                time.sleep(2.0)

        except Exception as e:
            print_err(f"Lỗi ở luồng AMR: {e}")
            time.sleep(2)


# =========================================================================
# LUỒNG 2: MÔ PHỎNG CÁNH TAY COBOT GIAO TIẾP VỚI SERVER
# =========================================================================
def cobot_client_simulation():
    last_msg = None
    while True:
        try:
            # 1. COBOT LIÊN TỤC HỎI XEM ĐÃ ĐƯỢC PHÉP LÀM VIỆC CHƯA
            # Request: GET /api/cobot/current_task
            res = requests.get(f"{API_URL}/api/cobot/current_task")
            data = res.json()
            status = data.get("status")

            if status == "waiting_amr":
                target = data.get("target_token")
                msg = f"Đang chờ AMR chạy tới [{target}]..."
                if msg != last_msg:
                    print_cobot(f"Gửi GET /api/cobot/current_task ➔ Server: {data}")
                    last_msg = msg
                time.sleep(1.5)  # Hỏi thăm định kỳ mỗi 1.5 giây

            elif status == "ready":
                last_msg = None
                task_data = data["data"]
                pose_id = task_data["pose_id"]
                pallet_id = task_data["pallet_stt"]
                model_hang = task_data["model"]
                abs_slot = task_data["absolute_slot"]
                rem = data.get("remaining_boxes_at_stop", 1)

                print("\n" + "─" * 70)
                print_cobot(f"🎯 [HÀNH ĐỘNG] Nhận tín hiệu READY từ Server! Tọa độ gắp thả: {pose_id}")
                print_cobot(f"   ↳ (Còn phải thả {rem} thùng tại trạm này). Đang vung tay gắp thả...")

                # Làm chậm lại 3 giây để mô phỏng Cobot quắp hàng
                time.sleep(3.0)

                # 2. THẢ XONG, COBOT BÁO CÁO LÊN SERVER ĐỂ LẤP LỖ TRỐNG VÀ TRỪ SỐ LƯỢNG
                # Request: PUT /records/quantity
                put_payload = {
                    "stt": pallet_id,
                    "so_pallet": pallet_id,
                    "model": model_hang,
                    "so_luong": 0,
                    "added_slot": abs_slot
                }
                print_cobot(f"📝 [HÀNH ĐỘNG] Thả xong! Đang gọi: POST /records/quantity (Chốt ô số {abs_slot})")
                res_post = requests.post(f"{API_URL}/records/quantity", json=put_payload)
                print_cobot(f"   ↳ Server cập nhật kho thành công: {res_post.json()}")
                print("─" * 70 + "\n")

                time.sleep(1.0)

            elif status == "idle":
                time.sleep(2.0)

        except Exception as e:
            print_err(f"Lỗi ở luồng Cobot: {e}")
            time.sleep(2)


if __name__ == "__main__":
    os.system('color')
    print("🔍 MỞ RỘNG GIÁM SÁT GIAO TIẾP API (CHẾ ĐỘ CHẬM - SLOW MOTION)")
    if not check_server_ready():
        sys.exit(1)

    t_amr = threading.Thread(target=amr_client_simulation, daemon=True)
    t_cobot = threading.Thread(target=cobot_client_simulation, daemon=True)

    t_amr.start()
    t_cobot.start()

    while True:
        time.sleep(2)