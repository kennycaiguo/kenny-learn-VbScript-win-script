import wmi

wmiObj = wmi.WMI()

for os in wmiObj.Win32_OperatingSystem():  # 相当于"select * from Win32_OperatingSystem"
    print(os.Caption)  # Microsoft Windows 11 专业工作站版
    print(os.BootDevice)  # \Device\HarddiskVolume1
    print(os.BuildNumber)  # 26100
    print(os.CodeSet)  # 936
    print(os.CountryCode)  # 86,中国的country code
    print(os.CSDVersion)  # 获取不到
    print(os.CreationClassName)  # Win32_OperatingSystem
    print(os.Description)  # Win32_OperatingSystem
    print(os.Manufacturer)  # Microsoft Corporation
