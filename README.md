# AutoTest

这是一个命令行工具
```
>AutoTest.exe -h
usage: AutoTest.exe [-h] [-f FILE] [-o [LOG]]

optional arguments:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  input file(.xlsx)
  -o [LOG], --log [LOG] output logfile(.txt)
```
# Build Exe

1. install pyinstaller
```
pip install pyinstaller
```
2. Pack
```
pyinstaller -F ./AutoTest.py
```

# Dependent

```
pip install chardet
```
# 注意

excel 需要设置: 文件 -> 选项 -> 校对 -> 自动更正选项 -> 键入时自动套用格式 : 去掉 "internet及网络路径替换为超链接"的勾选

demo 的excel 中有批注,如果excel版本较低的话,可能导致批注无法显示

# 后续设计的思考

几个模块
1. excel解析和数据分析模块
2. 串口模块: 使用串口和域socket实现
3. 其他模块: 使用如tcp/udp + 域socket实现,还可以支持


# http url 测试支持
在CMD.xlsx 中, 如果要支持http,则端口部分,设置端口参数为
```
vport=VS-1616H2A_HTTP
config=http://192.168.20.141:80
```
注意config的末尾不要加`/`

在CMD中,'send'写为如下格式,这是一个json格式的配置,各项参数值可以通过postman的 `Code snippet -> Python - Requests`生成.
```
{
	"methed":"GET",
	"url" : "/test.cgx?cmd=VID?",
	"headers":{},
	"payload":{}
}
```
在postman中,`http://192.168.20.141/test.cgx?cmd=VID?`请求可转换为
```
import requests

url = "http://192.168.20.141/test.cgx?cmd=VID?"

payload={}
headers = {}

response = requests.request("GET", url, headers=headers, data=payload)

print(response.text)
```
将url,payload,headers以json形式填写到表格里即可

http 命令的测试,接收部分的判断是根据 待接收部分是否包含在接收数据中进行的.所以可以将接收的数据的头尾去掉,进行匹配

