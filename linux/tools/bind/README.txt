这里可以放「便携版 dig」，实现免安装使用（适合没有 root/sudo、
或不允许用包管理器安装的机器）。

怎么用：
1. 复制一个可用的 dig 到本目录，最终确保存在可执行文件：

       linux/tools/bind/dig

   程序会自动优先使用它（查找顺序：PATH → tools/bind/dig → /usr/bin/dig 等）。

2. 如果 dig 还有依赖的动态库（一般发行版的 dig 只依赖 libc / libcrypto），
   可以用下面命令确认还能带哪些库：

       ldd /usr/bin/dig

   把缺失的 .so 也一并放到本目录，然后用 LD_LIBRARY_PATH 启动，例如：

       LD_LIBRARY_PATH="$PWD/tools/bind" python3 domain_dns_test.py

3. 别忘了可执行权限：

       chmod +x tools/bind/dig

更简单的做法（推荐）：在 Debian/Ubuntu 上直接安装 dnsutils（或 bind9-dnsutils）：

    sudo apt-get install -y dnsutils

RHEL/CentOS 用 bind-utils，Alpine 用 bind-tools。
