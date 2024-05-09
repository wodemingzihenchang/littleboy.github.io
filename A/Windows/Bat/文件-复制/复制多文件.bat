@echo off
for /f %%i in (
nameList.txt
) do copy 02-草图.md %%i.md
 
::循环nameList.txt里面的字符
:: do mkdir %%i && copy test.docx %%i\%%i.docx && start %%i\%%i.docx
:: do copy 复制的文件 新名字
::pause
