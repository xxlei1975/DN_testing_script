# DN_testing_script —— 非洲域名批量 DNS 解析 / 运营商简化 / 时延丢包测试

从 Excel（`非洲域名测试列表.xlsx`）读取域名，对表里配置的每个 DNS 服务器执行
`dig @<DNS服务器IP> <域名> +short`，把 **CNAME 链路**、**A 记录 IP**、
**运营商简化结果（CT / CU / CM）**、**ping 平均时延与丢包率**写回各自的列，
并在数据区全部处理完后重算每个 DNS 块下方的**表尾统计**。

程序不依赖硬编码的行列号：表头行、DNS 服务器 IP 行、数据区、表尾统计位置、
各块的列（CNAME / A / IP归属 / 时延 / 丢包率 / 首IP归属 / 时延改善 / 丢包改善）
都是**自动探测**出来的，表结构小改（改 IP、加减行）不用改代码。

## 两个版本

| 目录 | 平台 | 一键入口 | 环境部署 | 说明文档 |
| --- | --- | --- | --- | --- |
| [`windows/`](windows/) | Windows 10/11 | 双击 `run_dns_test.bat` | `scripts\setup_env.ps1`（winget 装 Python / BIND，支持离线 wheels） | [windows/README.md](windows/README.md)、[部署到其他Windows电脑.md](windows/部署到其他Windows电脑.md) |
| [`linux/`](linux/) | Linux（Debian/Ubuntu、RHEL/CentOS、Alpine…） | `./run_dns_test.sh` | `scripts/setup_env.sh`（apt/dnf/yum/apk/zypper/pacman 自动识别） | [linux/README.md](linux/README.md) |

两个版本的功能、表格结构规则、统计口径、命令行参数**完全一致**，只有和操作系统相关的
部分不同，便于两个平台上跑出来的结果互相对拍：

| 差异点 | Windows 版 | Linux 版 |
| --- | --- | --- |
| ping 命令 | `ping -n <报文数> -w <毫秒> <IP>` | `ping -c <报文数> -W <秒> -i <间隔秒> <IP>` |
| ping 输出解析 | 中文/英文 Windows 摘要（`平均 = 23ms`、`(0% 丢失)`） | iputils / busybox 摘要（`rtt min/avg/max/mdev`、`N% packet loss`），同时保留对 Windows 输出的解析 |
| dig 查找顺序 | PATH → `tools\bind\dig.exe` → WinGet/BIND 安装目录 | PATH → `tools/bind/dig` → `/usr/bin/dig` 等常见位置 |
| 控制台解码 | OEM 代码页（cp936）→ ANSI → UTF-8 | 系统区域编码 → UTF-8 |
| 子进程 | 传 `creationflags` 隐藏黑窗 | 不传（Linux 无此参数） |
| 离线依赖 | `wheels/`（Windows x64 / Python 3.11~3.14） | 默认走 PyPI；若放了与本机匹配的 `wheels/` 会优先离线安装 |

## 快速开始

### Windows

```
双击 windows\run_dns_test.bat
```

缺 Python / 依赖 / dig 时会自动安装（详见 `windows/部署到其他Windows电脑.md`），
参数可以原样传给 bat，例如 `run_dns_test.bat --limit 5 --dry-run`。

### Linux

```bash
cd linux
./run_dns_test.sh --limit 5 --dry-run     # 先试跑 5 个域名（不写文件）
./run_dns_test.sh                          # 全量运行，写出 *_result.xlsx
```

`run_dns_test.sh` 会自动调用 `scripts/setup_env.sh` 准备环境（Python 3、venv、
依赖、`dig`），首次运行可能需要 sudo/root 权限装包。也可以单独执行：

```bash
bash scripts/setup_env.sh            # 只准备环境
bash scripts/setup_env.sh --no-install   # 只体检，不动系统
```

## 常用参数（两版通用）

| 参数 | 说明 |
| --- | --- |
| `-i/--input` | 输入 Excel，默认 `非洲域名测试列表.xlsx` |
| `-o/--output` | 输出 Excel，默认在输入文件名后加 `_result` |
| `--inplace` | 直接回写原文件（建议先备份） |
| `--limit N` | 只处理前 N 个域名（试跑，默认不写表尾统计） |
| `--dry-run` | 只打印结果预览，不写 Excel |
| `--workers N` | dig 并发线程数，默认 8 |
| `--no-ip-info` | 跳过 ip-api 归属地查询（归属列留空，同时跳过表尾统计） |
| `--use-dig-cache` | 复用上次 dig 结果（试跑更快，数据非实时） |
| `--no-ping` | 跳过 ping（不写时延/丢包率） |
| `--ping-count N` / `--ping-timeout MS` / `--ping-workers N` | ping 报文数 / 单包超时 / 并发数 |
| `--ping-interval S` | 仅 Linux：报文间隔秒数，默认 1（与 Windows 一致，调小可加速） |
| `--no-ping-cache` / `--no-ping-retry` | 不用 ping 缓存 / 丢包率不为 0% 时不重测 |
| `--no-autofit-row-height` | 不按内容自动调整行高 |
| `--keep-duplicates` | 保留重复域名行（默认自动删除多余行并备份到 `backup/`） |
| `--force-stats` | 配合 `--limit` 时也写入表尾统计 |

## 结果与校验

- 运行结果写入 `*_result.xlsx`（`--inplace` 时写回原文件）。
- 表尾 11 项统计：无效解析行数 / 非三大运营商行数 / 三大运营商行数 /
  仅CT / 仅CU / 仅CM / CT+CU / CT+CM / CU+CM / CT+CU+CM / 首IP归属为CM行数。
- 单元测试（两版各自目录下）：

  ```bash
  # Linux
  cd linux && .venv/bin/python -m unittest discover -s tests
  # Windows
  cd windows && .venv\Scripts\python.exe -m unittest discover -s tests
  ```

- 结果文件校验（列格式、逐行对应、改善列、表尾统计重算比对）：

  ```bash
  python tests/verify_result.py 非洲域名测试列表_result.xlsx
  ```

## 运行产生的目录

| 目录/文件 | 说明 |
| --- | --- |
| `cache/` | IP 归属地缓存、dig 缓存、ping 缓存、已解析的解释器路径（`python_path.txt`，含本机计算机名，**属于单机缓存**） |
| `backup/` | 自动删除重复域名行之前的带时间戳备份 |
| `.venv/` | 自动创建的解释器环境 |

> `cache/`（尤其 `python_path.txt`）和 `.venv/` 都写死了本机的绝对路径，
> **拷贝给别的电脑时不要带**：新机器上会被自动识别并忽略/重建。
> 仓库里已经用 `.gitignore` 排除了它们。

## 数据来源

IP 归属地来自 <https://ip-api.com/>（免费版，仅支持 HTTP，带限速与重试，结果本地缓存）。
