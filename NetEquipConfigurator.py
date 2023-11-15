import pandas as pd
import serial
import serial.tools.list_ports
import time
import os

def list_serial_ports():
    ports = serial.tools.list_ports.comports()
    available_ports = [port.device for port in ports]
    return available_ports

def select_serial_port(available_ports):
    print("사용 가능한 시리얼 포트: ", available_ports)
    while True:
        choice = input("연결할 시리얼 포트 번호를 입력하세요 (예: COM1 또는 1): ").upper()
        if choice.startswith("COM"):
            selected_port = choice
        else:
            selected_port = f"COM{choice}"

        if selected_port in available_ports:
            return selected_port
        else:
            print("잘못된 포트 번호입니다. 다시 시도하세요.")

def read_device_configs(file_path):
    return pd.read_excel(file_path)

def apply_config(serial_connection, config_commands):
    # 여러 줄의 명령어를 개행 문자로 분리하여 전송
    commands = config_commands.split('\n')
    for command in commands:
        if command:  # 빈 줄은 무시
            serial_connection.write((command + '\n').encode())
            time.sleep(1)
            # 장비로부터의 응답 읽기 및 출력
            while serial_connection.inWaiting() > 0:
                response = serial_connection.read(serial_connection.inWaiting())
                print(response.decode('utf-8'), end='')

def main():
    script_dir = os.path.dirname(__file__)  # 스크립트 파일의 디렉토리
    file_path = os.path.join(script_dir, "device_configs.xlsx")

    device_configs = read_device_configs(file_path)

    total_rows = len(device_configs)
    for index, device_config in enumerate(device_configs.iterrows(), start=1):
        _, device_config = device_config  # enumerate로 인한 unpacking
        available_ports = list_serial_ports()
        port = select_serial_port(available_ports)
        with serial.Serial(port, device_config['Baudrate'], timeout=1) as ser:
            # 시리얼 연결 후 최초로 엔터 키 입력
            ser.write(b'\r\n')
            time.sleep(1)
            apply_config(ser, device_config['Config'])
            print("\n설정 완료.")

            # 마지막 행이 아니라면 사용자 입력 대기
            if index < total_rows:
                print("다음 장비를 연결하고 엔터 키를 누르세요.")
                input()
            else:
                print("모든 장비 설정이 완료되었습니다. 프로그램을 종료합니다.")
                break

if __name__ == "__main__":
    main()

