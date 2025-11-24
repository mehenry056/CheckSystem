import datetime
import socket
from contextlib import nullcontext

import paramiko
import pandas as pd
import os

from pandas.core.dtypes.missing import nan_checker

import FormatOutPut
from FormatOutPut import txt_to_excel


# 读取目标服务器信息
# 读取服务器信息
def read_servers(file_path):
    df = pd.read_excel(file_path)
    servers = []
    for _, row in df.iterrows():
        servers.append({
            'module': row["Module"],
            'ip': row["IP"],a
            'username': row["Username"],
            'password': row["Password"]
        })
    return servers


# 通过跳板机连接到目标服务器
def connect_via_jump(jump_host, jump_user, jump_key_path, target_host, target_user, target_password, target_key_path):
    try:
        # 创建跳板机连接
        jump_client = paramiko.SSHClient()
        jump_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        jump_key = paramiko.RSAKey.from_private_key_file(jump_key_path)
        print("try to connect via jump via jump key")
        jump_client.connect(hostname=jump_host, username=jump_user, pkey=jump_key)
        print("connected via jump via jump key--------------------------")

        # 创建目标服务器连接
        transport = jump_client.get_transport()
        dest_addr = (target_host, 22)
        local_addr = ('localhost', 22)
        channel = transport.open_channel("direct-tcpip", dest_addr, local_addr)
        print("Opened the channel")

        target_client = paramiko.SSHClient()
        target_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        target_key = paramiko.RSAKey.from_private_key_file(target_key_path)
        if str(target_user) != 'nan':
            print("Ready to check: " + target_host + " with username and password")
            target_client.connect(hostname=target_host, username=target_user, password=target_password, pkey=target_key,
                                  sock=channel, timeout=3)
            print(target_host + " has been checked")
        else:
            print("Ready to check:" + target_host)
            target_client.connect(hostname=target_host, username="sbersh", pkey=target_key, sock=channel, timeout=3)
            print(target_host + " has been checked")

        return target_client
    except socket.timeout:
        print("SSH连接超时")
        return None
    except Exception as e:
        print(f"连接失败: {str(e)}")
        return None


# 执行远程命令
def execute_command(ssh, command):
    try:
        stdin, stdout, stderr = ssh.exec_command(command)
        output = stdout.read().decode()
        error = stderr.read().decode()
        if error:
            return f"错误信息: {error}"
        else:
            return output
    except Exception as e:
        return f"命令执行失败: {str(e)}"


# 主函数
def main():
    # 跳板机配置
    jump_host = '100.76.129.19'
    jump_user = 'coldbutter'
    jump_key_path = os.path.expanduser('~/.ssh/id_rsa')

    # 目标服务器信息文件路径
    servers_file = 'servers.xlsx'

    # 读取目标服务器信息
    servers = read_servers(servers_file)
    filepath = "C:\python\sys-check-output"
    now = datetime.datetime.now()
    outputfilename: str = now.strftime("%Y%m%d-%H%M%S")

    txtfilename = os.path.join(filepath, outputfilename + ".txt")
    excelfilename = os.path.join(filepath, outputfilename + ".xlsx")

    # 打开输出文件
    with open(txtfilename, 'w') as out_file:
        # 遍历目标服务器
        for server in servers:
            target_host = server['ip']
            target_user = server['username']
            target_password = server['password']
            target_key_path = os.path.expanduser('~/.ssh/id_rsa')  # 假设目标服务器使用相同的私钥

            # 通过跳板机连接目标服务器, 某个服务器连不上，就跳过
            try:
                target_client = connect_via_jump(jump_host, jump_user, jump_key_path, target_host, target_user, target_password, target_key_path)

                if target_client:
                    # 执行命令
                    result = execute_command(target_client, 'df -h')
                    out_file.write(f"服务器: {target_host}\n")
                    out_file.write(result + '\n\n')

                    # 关闭连接
                    target_client.close()
            except Exception as e:
                print("遇到的异常：" + e)
                out_file.write(f"服务器: {target_host}\n")
                out_file.write("连接遇到了异常，连接下一个服务器" + '\n\n')
                continue
            except ValueError as e:
                print("遇到的异常：" + e)
                out_file.write(f"服务器: {target_host}\n")
                out_file.write("连接遇到了异常，连接下一个服务器" + '\n\n')
                continue

    print("结果已保存到 output.txt")
    FormatOutPut.txt_to_excel(txtfilename, excelfilename)
    if os.path.exists(os.path.join(os.getcwd(), txtfilename)):
        os.remove(os.path.join(os.getcwd(), txtfilename))
        print(f"文件{txtfilename} 已经删除。")

if __name__ == '__main__':
    main()
