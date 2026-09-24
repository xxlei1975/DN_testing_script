这里可以放「便携版 dig」，实现免安装使用。

怎么用：
1. 找一台已经装好 BIND 的电脑（例如通过  winget install ISC.Bind  安装的），
   打开目录：
      C:\Users\<用户名>\AppData\Local\Microsoft\WinGet\Packages\ISC.Bind_Microsoft.Winget.Source_8wekyb3d8bbwe\
2. 把该目录里**所有文件**（dig.exe 以及它依赖的 *.dll）复制到本文件夹，
   最终确保存在：
      tools\bind\dig.exe
3. 之后运行 run_dns_test.bat 或 domain_dns_test.py 时，程序会自动使用这里的 dig，
   不需要管理员权限、不需要安装。

也可以用自己下载的 BIND 9 压缩包解压后的 bin 目录内容。
