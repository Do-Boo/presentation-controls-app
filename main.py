import sys
import os
import pyautogui
import websockets
import asyncio
import json

class RemoteControl:
    def __init__(self):
        # 마우스 설정
        pyautogui.MINIMUM_DURATION = 0
        pyautogui.MINIMUM_SLEEP = 0
        pyautogui.PAUSE = 0
        self.is_presenting = False
        print("Remote Control Server Started")

    async def start_server(self, websocket, path):
        print("Client Connected")
        try:
            async for message in websocket:
                data = json.loads(message)
                await self.handle_command(data)
        except Exception as e:
            print(f"Error: {e}")

    async def handle_command(self, data):
        try:
            command_type = data.get('type', '')
            print(f"Received command: {command_type}")

            if command_type == 'command':
                # 프레젠테이션 제어 명령
                command = data.get('command', '')
                if command == 'NEXT_SLIDE':
                    pyautogui.press('right')
                elif command == 'PREV_SLIDE':
                    pyautogui.press('left')
                elif command == 'START_PRESENTATION':
                    pyautogui.press('f5')
                    self.is_presenting = True
                elif command == 'END_PRESENTATION':
                    pyautogui.press('esc')
                    self.is_presenting = False
                elif command == 'TOGGLE_BLACK_SCREEN':
                    pyautogui.press('b')
                elif command == 'TOGGLE_WHITE_SCREEN':
                    pyautogui.press('w')
                
            elif command_type == 'laser':
                # 레이저 포인터 (마우스 이동)
                x = float(data.get('x', 0))
                y = float(data.get('y', 0))
                screen_width, screen_height = pyautogui.size()
                target_x = int(x * screen_width)
                target_y = int(y * screen_height)
                pyautogui.moveTo(target_x, target_y)

            elif command_type == 'mouse_move':
                # 일반 마우스 이동
                dx = data.get('dx', 0)
                dy = data.get('dy', 0)
                current_x, current_y = pyautogui.position()
                pyautogui.moveTo(current_x + dx, current_y + dy)
            
            elif command_type == 'mouse_click':
                # 마우스 클릭
                button = data.get('button', 'left')
                pyautogui.click(button=button)
            
            elif command_type == 'key':
                # 키보드 입력
                key = data.get('key', '')
                if key:
                    if isinstance(key, list):
                        pyautogui.hotkey(*key)  # 단축키
                    else:
                        pyautogui.press(key)    # 단일 키

            # PowerPoint 특수 명령
            elif command_type == 'presentation_command':
                command = data.get('command', '')
                if command == 'START_FROM_BEGIN':
                    pyautogui.hotkey('shift', 'f5')  # 처음부터 시작
                elif command == 'START_FROM_CURRENT':
                    pyautogui.press('f5')            # 현재 슬라이드부터
                elif command == 'GOTO_SLIDE':
                    slide_number = data.get('slide', '1')
                    # 슬라이드 번호 입력 후 엔터
                    pyautogui.write(slide_number)
                    pyautogui.press('enter')

        except Exception as e:
            print(f"Error executing command: {e}")

    def handle_presentation_key(self, key):
        """프레젠테이션 관련 특수 키 처리"""
        presentation_keys = {
            'next': 'right',
            'prev': 'left',
            'start': 'f5',
            'end': 'esc',
            'black': 'b',
            'white': 'w',
        }
        if key in presentation_keys:
            pyautogui.press(presentation_keys[key])

async def main():
    try:
        if len(sys.argv) < 2:
            print("Error: Connection code is required")
            print("Usage: python main.py <connection_code>")
            return

        connection_code = sys.argv[1]
        print(f"Received connection code: {connection_code}")  # 디버깅용 로그 추가
        
        host = "sibar.roma.com"
        port = 1828
        
        print(f"Connecting to Remote Control Server at {host}:{port} with code {connection_code}")
        
        remote_control = RemoteControl()
        
        uri = f"ws://{host}:{port}"
        async with websockets.connect(uri) as websocket:
            print(f"WebSocket connected, registering with code: {connection_code}")
            # 연결 코드를 사용하여 PC 등록
            await websocket.send(json.dumps({
                "type": "register_pc",
                "code": connection_code
            }))
            
            while True:
                try:
                    message = await websocket.recv()
                    data = json.loads(message)
                    await remote_control.handle_command(data)
                except websockets.exceptions.ConnectionClosed:
                    print("Connection to server closed")
                    break
                except Exception as e:
                    print(f"Error in message handling: {e}")
                    break
    except Exception as e:
        print(f"Main function error: {e}")
        raise

if __name__ == "__main__":
    try:
        print(f"Starting with arguments: {sys.argv}")  # 디버깅용 로그 추가
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nServer stopped by user")
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)
