import paramiko, os

ip, user, pw = "114.132.62.244", "ubuntu", "2983801136Aa!"

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(ip, username=user, password=pw, timeout=10)
print("[1] 已连接")

sftp = ssh.open_sftp()
ssh.exec_command("mkdir -p /home/ubuntu/reports2")
sftp.put("AI质检工单准确率明细表_8089.html", "/home/ubuntu/reports2/index.html")
print("[2] 已上传")

ssh.exec_command("pkill -f 'http.server 8089' 2>/dev/null; pkill -f 'server8089' 2>/dev/null; sleep 1")
ssh.exec_command("cd /home/ubuntu/reports2 && nohup python3 /home/ubuntu/server8089.py > /tmp/http8089.log 2>&1 &")
print("[3] 已启动 (端口8089)")

stdin, stdout, stderr = ssh.exec_command("curl -s -o /dev/null -w '%{http_code}' http://localhost:8089/")
print(f"[4] 验证: HTTP {stdout.read().decode().strip()}")

sftp.close(); ssh.close()
print(f"\n访问: http://{ip}:8089/质检类型准确率分析.html")
print("请开放安全组 TCP:8089")
