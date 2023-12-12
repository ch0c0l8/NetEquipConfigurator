import pandas as pd
import serial
import serial.tools.list_ports
import paramiko
import telnetlib
import time
import os

def list_serial_ports():
    ports = serial.tools.list_ports.comports()
    available_ports = [port.device for port in ports]
    return available_ports

def select_serial_port(available_ports):
    while True:
        ports = serial.tools.list_ports.comports()
        available_ports = [port.device for port in ports]
        print("사용 가능한 시리얼 포트: ", available_ports)
        choice = input("연결할 시리얼 포트 번호를 입력하세요 (예: COM1 또는 1): ").upper()
        if choice.startswith("COM"):
            selected_port = choice
        else:
            selected_port = f"COM{choice}"

        if selected_port in available_ports:
            return selected_port
        else:
            print("잘못된 포트 번호입니다. 다시 시도하세요.")

def clear_buffer(channel):
    while channel.recv_ready():
        channel.recv(1024)

def ssh_connect(ip, port, username, password):
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(ip, port, username=username, password=password)
        channel = client.invoke_shell()
        clear_buffer(channel)
        return client, channel
    except Exception as e:
        print(f"SSH 연결 실패: {e}")
        return None, None

def telnet_connect(ip, port):
    try:
        return telnetlib.Telnet(ip, port=port)
    except Exception as e:
        print(f"Telnet 연결 실패: {e}")
        return None

def read_device_configs(file_path):
    return pd.read_excel(file_path)

def apply_config(connection, config_commands, connection_type):
    commands = config_commands.split('\n')
    if connection_type == 'SERIAL':
        for command in commands:
            if command:
                connection.write((command + '\n').encode())
                time.sleep(1)
                while connection.inWaiting() > 0:
                    response = connection.read(connection.inWaiting())
                    print(response.decode('utf-8'), end='\n')
    elif connection_type == 'SSH':
        channel = connection
        for command in commands:
            if command:
                channel.send(command + "\n")
                time.sleep(1)
                while channel.recv_ready():
                    response = channel.recv(9999).decode('utf-8')
                    print(response, end='\n')
    elif connection_type == 'TELNET':
        for command in commands:
            if command:
                connection.write((command + '\n').encode('ascii'))
                time.sleep(1)
                response = connection.read_very_eager()
                print(response.decode('ascii'), end='\n')

def main():
    print("엑셀 파일 데이터 읽어오는 중...")
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, "device_configs.xlsx") 
    device_configs = read_device_configs(file_path)
    print("엑셀 파일 데이터 읽기 성공")
    input("프로그램을 시작하려면 엔터 키를 누르세요...")
    for index, device_config in enumerate(device_configs.iterrows(), start=1):
        _, device_config = device_config
        connection_type = device_config['ConnectionType'].upper()

        if connection_type == 'SERIAL':
            while True:  # 시리얼 포트 연결 시도 루프
                try:
                    ports = serial.tools.list_ports.comports()
                    available_ports = [port.device for port in ports]
                    comport = device_config['COMPort']
                    baudrate = device_config['Baudrate']
                    data_bits = device_config['DataBits']
                    stop_bits = device_config['StopBits']
                    parity = device_config['Parity'].upper()

                    # Parity 값 처리
                    if parity == 'NONE':
                        parity = serial.PARITY_NONE
                    elif parity == 'ODD':
                        parity = serial.PARITY_ODD
                    elif parity == 'EVEN':
                        parity = serial.PARITY_EVEN
                    elif parity == 'MARK':
                        parity = serial.PARITY_MARK
                    elif parity == 'SPACE':
                        parity = serial.PARITY_SPACE
                    # 추가적인 parity 옵션 처리가 필요하면 여기에 추가

                    if comport not in available_ports:
                        print(f"포트 {comport} 사용 불가. 사용 가능한 포트를 선택하세요.")
                        comport = select_serial_port(available_ports)

                    print(f'접속 정보: {connection_type}, COMPort={comport}, Baudrate={baudrate}, DataBits={data_bits}, Parity={parity}, StopBits={stop_bits}')

                    with serial.Serial(comport, baudrate, bytesize=data_bits, parity=parity, stopbits=stop_bits, timeout=3) as ser:
                        ser.write(b'\r\n')
                        time.sleep(1)
                        apply_config(ser, device_config['Config'], connection_type)
                    break  # 성공적으로 연결되면 루프 탈출

                except serial.SerialException as e:
                    print(f"시리얼 포트 연결 실패: {e}")
                    while True:
                        retry = input("다시 시도하시겠습니까? (y/n): ").lower()
                        if retry == 'y':
                            break  # 다시 시도
                        elif retry == 'n':
                            return  # 프로그램 종료
                        else:
                            print("잘못된 입력입니다. 'y' 또는 'n'을 입력하세요.")

        elif connection_type == 'SSH':
            ip = str(device_config['IP'])
            username = str(device_config['Username'])
            password = str(device_config['Password'])
            port = device_config.get('Port')
            if pd.isna(port):
                port = 22
            else:
                port = int(port)
            client, channel = ssh_connect(ip, port, username, password)

            print(f'접속 정보: {connection_type}, IP={ip}, Port={port}, Username={username}, Password={password}')

            if client and channel:
                apply_config(channel, device_config['Config'], connection_type)
                channel.close()
                client.close()
            else:
                print("SSH 연결 실패")

        elif connection_type == 'TELNET':
            ip = str(device_config['IP'])
            port = device_config.get('Port')
            if pd.isna(port):
                port = 23
            else:
                port = int(port)
            telnet = telnet_connect(ip, port)

            print(f'접속 정보: {connection_type}, IP={ip}, Port={port}')

            if telnet:
                apply_config(telnet, device_config['Config'], connection_type)
                telnet.close()
            else:
                print("TELNET 연결 실패")
        
        else:
            print(f"알 수 없는 ConnectionType: '{connection_type}'")

        if index < len(device_configs):
            print("다음 장비를 연결하고 엔터 키를 누르세요.")
            input()
        else:
            print("모든 장비 설정이 완료되었습니다. 프로그램을 종료합니다.")
            break

if __name__ == "__main__":
    main()
