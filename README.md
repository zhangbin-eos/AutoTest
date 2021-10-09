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

## Dependent

```
pip install chardet
```
# 使用说明
## 填写命令表格
1. 参考demo中提供的表格
2. 首先填写要用到的接口,目前支持了串口和http
3. 然后填写需要发送的数据

## 远传接收日志
脚本内部并未实现远程控制的功能,但是通过一些第三方工具可以简单的实现发送日志到远传服务器的功能,如下
```bash
# @Author: zhangbin.eos@foxmail.com
# @Date:   2021-09-30 10:06:27
# @Last Modified by:   zhangbin.eos@foxmail.com
# @Last Modified time: 2021-10-09 13:48:10
./AutoTest.exe -f CMD.xlsx -o logfile.log  > /dev/null &

trap "kill  $(jobs -p)" SIGINT

# 只发送错误日志
# tail -f logfile.log | grep -E "(ERROR)|(loop)" | logger.exe --udp -n 192.168.20.35 -P 8080 --rfc5424 --prio-prefix

tail -f logfile.log | logger.exe --udp -n 192.168.20.35 -P 8080 --prio-prefix
```
其中发送远程日志是通过`logger`程序实现的,之所以不使用`python`中的`logging`模块中的`socket句柄`,是因为`socket`的读写可能会影响到测试流程,造成测试过程意外的阻塞

`logger`的使用方法参考其help选项说明,列举的很详细

除了`logger`外,还有其他很多的工具,如在linux中的`nc`命令等,可以实现数据的转发,只要把这些工具拼接起来就行.

## http url 测试支持
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

对于一个post命令,使用postman转换后,得到
```
import requests

url = "http://192.168.20.196/cgi-bin/ligline.cgi"

payload = "{\"cmd\":\"video_info\"}"
headers = {
  'content-type': 'application/json;charset=UTF-8',
  'accept': 'application/json, text/plain, */*'
}

response = requests.request("POST", url, headers=headers, data=payload)

print(response.text)

```
对于的excel中的send数据,如下,注意上方的单引号修改为如下的双引号
```
{
	"methed":"POST",
	"url" : "/cgi-bin/ligline.cgi",
	"headers":{
		"'content-type": "application/json;charset=UTF-8",
		"accept": "application/json, text/plain, */*"
	},
	"payload":"{\"cmd\":\"video_info\"}"
}

```
# 注意

excel 需要设置: 文件 -> 选项 -> 校对 -> 自动更正选项 -> 键入时自动套用格式 : 去掉 "internet及网络路径替换为超链接"的勾选

demo 的excel 中有批注,如果excel版本较低的话,可能导致批注无法显示

目前已知的一个问题是在git的bash终端中,`AutoTest.exe`运行时不会在标准输出中打印数据,在程序关闭后,数据

# 后续设计的思考

几个模块
1. excel解析和数据分析模块
2. 串口模块: 使用串口和域socket实现
3. 其他模块: 使用如tcp/udp + 域socket实现,还可以支持

