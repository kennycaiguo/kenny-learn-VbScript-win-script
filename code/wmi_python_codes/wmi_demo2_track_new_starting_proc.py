import wmi
"""这个程序可以检测新创建的进程,并且把他的名字显示在控制台上面"""

w = wmi.WMI()

# watch_for函数的参数种类: “creation”, “deletion”, “modification” or “operation”
process_watcher = w.Win32_Process.watch_for("creation")

while True:
    new_proc = process_watcher()
    print(new_proc.Caption)