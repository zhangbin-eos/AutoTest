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



