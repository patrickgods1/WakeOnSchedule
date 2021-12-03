import pandas as pd
import struct
import socket
import datetime
import time
from config import BROADCAST_IP, DEFAULT_PORT, hwAddress


def create_magic_packet(macaddress):
    """
    Create a magic packet.
    A magic packet is a packet that can be used with the for wake on lan
    protocol to wake up a computer. The packet is constructed from the
    mac address given as a parameter.
    Args:
        macaddress (str): the mac address that should be parsed into a
            magic packet.
    """
    if len(macaddress) == 12:
        pass
    elif len(macaddress) == 17:
        sep = macaddress[2]
        macaddress = macaddress.replace(sep, '')
    else:
        raise ValueError('Incorrect MAC address format')

    # Pad the synchronization stream
    data = b'FFFFFFFFFFFF' + (macaddress * 16).encode()
    send_data = b''

    # Split up the hex values in pack
    for i in range(0, len(data), 2):
        send_data += struct.pack(b'B', int(data[i: i + 2], 16))
    return send_data


def send_magic_packet(*macs, **kwargs):
    """
    Wake up computers having any of the given mac addresses.
    Wake on lan must be enabled on the host device.
    Args:
        macs (str): One or more macaddresses of machines to wake.
    Keyword Args:
        ip_address (str): the ip address of the host to send the magic packet
                     to (default "255.255.255.255")
        port (int): the port of the host to send the magic packet to
               (default 9)
    """
    packets = []
    ip = kwargs.pop('ip_address', BROADCAST_IP)
    port = kwargs.pop('port', DEFAULT_PORT)
    for k in kwargs:
        raise TypeError('send_magic_packet() got an unexpected keyword '
                        'argument {!r}'.format(k))

    for mac in macs:
        packet = create_magic_packet(mac)
        packets.append(packet)

    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    sock.connect((ip, port))
    for packet in packets:
        sock.send(packet)
    sock.close()


def main():
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    # Read in courses from Excel
    # 1     B   Date
    # 3     D   Type
    # 4     E   Start Time
    # 6     G   End Time
    # 9     J   Section Number
    # 11    L   Section Title
    # 12    M   Instructor
    # 13    N   Building
    # 15    P   Room
    # 16    Q   Configuration
    # 17    R   Technology
    # 18    S   Section Size
    # 20    U   Notes
    # 22    W   Approval Status

    scheduleFile = f'SectionScheduleDailySummary (2).xls'
    schedule = pd.read_excel(scheduleFile, header=6, skipfooter=1, usecols=[1,4,6,15,22])
    schedule = schedule.loc[(schedule['Approval Status'] == 'Final Approval')]
    schedule['Start Time'] = pd.to_datetime(schedule['Start Time']).dt.strftime('%H:%M')
    schedule['End Time'] = pd.to_datetime(schedule['End Time']).dt.strftime('%H:%M')
    schedule['Date'] = pd.to_datetime(schedule['Date'], errors = 'coerce',format = '%Y-%m-%d')
    schedule['Date'] = schedule['Date'].dt.strftime('%Y-%m-%d')
    schedule.sort_values(by=['Date','Start Time', 'Room'], inplace=True)

    for _, row in schedule[schedule['Date'] >= datetime.datetime.now().strftime('%Y-%m-%d')].iterrows():
        try:
            difference = (datetime.datetime.strptime(f"{row['Date']} {row['Start Time']}", '%Y-%m-%d %H:%M') - datetime.timedelta(minutes=30) - datetime.datetime.now()).total_seconds()
            if difference < 0.5:
                print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Waking {row['Room']}, which started at {row['Start Time']}.")
                send_magic_packet(*hwAddress[row['Room']], ip_address='169.229.232.127', port=9)
            else:
                print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Sleeping for {difference/60:.2f} minutes. {row['Room']} starts at {row['Start Time']}.")
                time.sleep(difference)
                print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Waking {row['Room']}, which starts at {row['Start Time']}.")
                send_magic_packet(*hwAddress[row['Room']], ip_address='169.229.232.127', port=9)

        except KeyboardInterrupt:
            print('Exiting......')
            break


if __name__ == '__main__':
    main()
