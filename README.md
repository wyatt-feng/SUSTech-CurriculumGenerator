# SUSTech iCal课表生成工具

这个脚本可以帮你把从[南科大教务系统](https://tis.sustech.edu.cn)下载的xlsx格式的课表转化为可以导入日历软件通用的iCal（.ics）格式的课表。

## 为什么要生成iCal格式课表
可以导入设备自带日历软件，从而与系统深度整合，例如基本上不占用系统资源实现提醒功能，或通过Siri等语音助手管理课程安排。

<img src="images/Siri-Integration.png" width = "50%" />

## 使用方式
1. 首先在教务系统主界面左侧点击课表，下载xlsx格式课表。
2. 若电脑存在[Python](https://www.python.org)环境及[openpyxl](https://openpyxl.readthedocs.io/en/stable/)，则下载[CurriculumGenerator](https://github.com/dazhi0619/CurriculumGenerator)目录下的两个.py文件。
   如果您不知道我在说什么，请[在此处](https://github.com/dazhi0619/CurriculumGenerator/releases/)下载打包好的可执行文件。
3. 运行
   ```
   python3 CurriculumGenerator.py <xlsx课表文件名> <学期开始日期> <学期结束日期>
   ```
   若您下载的是可执行文件，请在下载文件夹中按住shift右键单击文件夹空白处，选择“在此处打开powershell窗口”，在弹出的窗口中输入
   ```
   .\CurriculumGenerator <xlsx课表文件名> <学期开始日期> <学期结束日期>
   ```
4. 将生成好的`课表.ics`导入日历软件。通常情况下直接打开即可。对于iPhone和iPad，请将此文件AirDrop到您的设备上，或设法通过Safari浏览器打开此文件。
