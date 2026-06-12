# -*- coding: utf-8 -*-
"""部署报表到腾讯云 Ubuntu 服务器"""

import paramiko
import os

SERVER_IP = "114.132.62.244"
SERVER_USER = "ubuntu"
SERVER_PASSWORD = "2983801136Aa!"
REMOTE_DIR = "/home/ubuntu/reports"

LOCAL_FILES = [
    "AI质检工单准确率明细表.html",
    "工单质检规则准确率分析报表.html",
]


def deploy():
    print("=" * 55)
    print(f"部署到 {SERVER_IP} (Ubuntu 22.04)")
    print("=" * 55)

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SERVER_IP, username=SERVER_USER, password=SERVER_PASSWORD, timeout=10)
    print("[1] 已连接")

    sftp = ssh.open_sftp()

    # 创建目录
    ssh.exec_command(f"mkdir -p {REMOTE_DIR}")
    print(f"[2] 目录已创建: {REMOTE_DIR}")

    # 上传文件
    for f in LOCAL_FILES:
        if os.path.exists(f):
            sftp.put(f, f"{REMOTE_DIR}/{f}")
            print(f"[3] 已上传: {f}")

    # 停旧进程,重新启动HTTP服务
    ssh.exec_command("pkill -f 'http.server 8088' 2>/dev/null; sleep 1")
    ssh.exec_command(f"cd {REMOTE_DIR} && nohup python3 -m http.server 8088 > /tmp/http.log 2>&1 &")
    print("[4] HTTP服务已启动 (端口8088)")

    # 验证
    stdin, stdout, stderr = ssh.exec_command("curl -s -o /dev/null -w '%{http_code}' http://localhost:8088/")
    code = stdout.read().decode().strip()
    print(f"[5] 本地验证: HTTP {code}")

    sftp.close()
    ssh.close()

    print()
    print("=" * 55)
    print("部署成功!")
    print()
    print(f"  报表地址:")
    print(f"  http://{SERVER_IP}:8088/AI质检工单准确率明细表.html")
    print(f"  http://{SERVER_IP}:8088/工单质检规则准确率分析报表.html")
    print()
    print("  请确保腾讯云安全组已开放 8088 端口(TCP)")
    print("  控制台: 安全组 -> 入站规则 -> 添加 TCP:8088")
    print("=" * 55)


if __name__ == '__main__':
    deploy()
