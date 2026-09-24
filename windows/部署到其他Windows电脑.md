# 在其他 Windows 电脑上运行本程序

本文档说明如何把 `Africa_DN_Testing` 这个文件夹拿到**另一台 Windows 电脑**上跑起来，
重点解决「那台机器可能没有 Python、没有 openpyxl/requests、没有 dig」的问题。

三种方式，按推荐程度排序：

| 方式 | 适用场景 | 需要手动做的事 | 是否需要外网 |
| --- | --- | --- | --- |
| **A. 一键脚本（推荐）** | 目标机器能上网 | 复制文件夹 → 双击 `run_dns_test.bat` | 需要（下载 Python/依赖/dig） |
| **B. 离线部署** | 目标机器在内网、不能上网 | 复制文件夹（含 `wheels\`）→ 双击 `run_dns_test.bat` | 不需要 |
| **C. 打包成 exe** | 完全不希望目标机器出现 Python 安装过程 | 在有网机器上打包，再拷贝 exe | 打包时需要 |

---

## 一、方式 A：一键自动（推荐）

### 前置要求

- Windows 10 / 11（64 位）。Windows 7/8 缺少 winget，请直接看方式 B。
- 目标机器**能访问外网**（或能访问公司内网 pip 镜像）。
- 安装 dig（BIND）时会弹出 **UAC 授权框**，需要点「是」；如果目标机器没有管理员权限，
  请改用方式 B（便携版 dig 不需要安装）。

### 操作步骤

**第 1 步：复制整个文件夹**

把 `Africa_DN_Testing` 整个文件夹复制到目标电脑（U 盘 / 共享盘 / 压缩包均可）。
必须一起带走的文件：

```
Africa_DN_Testing\
├─ domain_dns_test.py          主程序
├─ requirements.txt            依赖清单
├─ run_dns_test.bat            双击运行入口
├─ scripts\setup_env.ps1       自动装环境的脚本
├─ wheels\                     离线依赖包（可选，见方式 B）
├─ 非洲域名测试列表.xlsx        输入数据
└─ tools\bind\                 便携版 dig 放这里（可选）
```

**必须不要带的目录（重要）**

| 目录 / 文件 | 为什么不能带 |
| --- | --- |
| `.venv\` | 虚拟环境里写死了**原电脑**的 Python 路径（`pyvenv.cfg` 的 `home` 和解释器路径）。拷到新机器上是个“跑不起来”的壳，依赖也装不进去。程序会自动把它删掉重建 |
| `cache\` | 里面的 `python_path.txt` 记录的是**原电脑**的解释器绝对路径（例如 `D:\Tech Doc\...\.venv\Scripts\python.exe`）。带过去后新机器上这个路径根本不存在 |
| `__pycache__\`、`backup\`、`*_result.xlsx` | 无用产物，可一并排除 |

> `cache\python_path.txt` 现在是**绑定机器**的：文件里同时记录了写入它的计算机名，
> `run_dns_test.bat` 只有在“计算机名相同 + 解释器确实存在 + 能 import 依赖”时才会复用，
> 否则忽略并重新检测。即使误拷了 `cache\` 也不会再出问题（见第四节）。

**第 2 步：双击 `run_dns_test.bat`**

脚本会先自检，缺什么装什么，然后自动开始跑测试。整个过程不需要手动敲命令。

**第 3 步：看脚本都自动做了什么**

脚本 `scripts\setup_env.ps1` 按下面的顺序自动完成：

| 步骤 | 检测顺序 | 缺失时的自动处理 |
| --- | --- | --- |
| 1. Python 解释器 | 项目内 `.venv` → `PATH` 里的 `python` → `py` 启动器 → `%LOCALAPPDATA%\Programs\Python\Python3x` → `C:\Program Files\Python3x` → `C:\Python3x`（并自动排除微软商店的 0 字节假 `python.exe`） | ① `winget install Python.Python.3.13`（依次尝试 3.13 / 3.12 / 3.11，用户级安装，不需要管理员）<br>② winget 不可用时，自动从 python.org 下载最新的稳定版 `python-<版本>-amd64.exe`，用 `/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1` 静默安装 |
| 2. 虚拟环境 | 项目内 `.venv\Scripts\python.exe` | 用上一步的解释器 `python -m venv .venv` 自动创建（依赖装在 venv 里，不污染系统 Python） |
| 3. Python 依赖 | 执行 `python -c "import openpyxl, requests"` | ① 若存在 `wheels\` 且里面有 `.whl`，用 `pip install --no-index --find-links wheels -r requirements.txt` **离线安装**（完全不联网）<br>② 否则依次尝试：默认 PyPI → 清华镜像 → 阿里云镜像 |
| 4. dig 命令 | `PATH` → 项目内 `tools\bind\dig.exe` → `%LOCALAPPDATA%\Microsoft\WinGet\Packages\ISC.Bind*` → `C:\Program Files\ISC BIND*` | `winget install ISC.Bind --exact --silent`（会弹 UAC）。装不上时，主程序会提示改用便携版 dig |
| 5. 记录解释器路径 | — | 把最终解释器路径写入 `cache\python_path.txt`（ANSI 编码，同时记录**本机计算机名**），下次运行直接复用，启动更快。该文件属于单机缓存，**不要复制到其他电脑** |
| 6. 启动方式 | — | `run_dns_test.bat` 不依赖 `%PATH%` 里的 `powershell.exe`，而是用 `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe` 的绝对路径调用（`%PATH%` 被改坏的机器也能跑） |

**第 4 步（建议先验证再全量跑）**

双击运行等价于命令行执行 `python domain_dns_test.py`，所以也可以这样试跑：

```bat
run_dns_test.bat --limit 2 --dry-run     :: 只查前 2 个域名，只打印不写文件
```

确认没问题后再双击 `run_dns_test.bat` 全量运行，结果写入
`非洲域名测试列表_result.xlsx`（原文件不会被修改）。

### 只检查、不安装

想先看看这台机器缺什么，可以用（不会改动系统）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup_env.ps1 -DryRun
```

---

## 二、方式 B：离线 / 内网部署

适用于目标机器不能上外网、或者没有管理员权限安装 dig 的情况。

**第 1 步：在能上网的机器上准备离线依赖**

本项目已经带了一个 `wheels\` 目录（含 Windows x64、Python 3.11~3.14 的 wheel）。
如果你改了 `requirements.txt`，或有别的 Python 版本需求，重新生成一次即可：

```powershell
# 在当前 Python 版本上下载（最简单）
python -m pip download -r requirements.txt -d wheels --only-binary=:all:

# 指定目标机器是 Python 3.11~3.14 / Windows x64 时
python -m pip download -r requirements.txt -d wheels --only-binary=:all: `
       --platform win_amd64 --python-version 3.13
```

`setup_env.ps1` 一旦发现 `wheels\` 里有 `.whl`，就会优先离线安装，不访问网络。

**第 2 步：准备便携版 dig（免安装）**

在一台已经能用的机器上找到 BIND 目录（`dig.exe` 和它依赖的一堆 DLL 都在里面），例如：

```
C:\Users\<用户名>\AppData\Local\Microsoft\WinGet\Packages\ISC.Bind_Microsoft.Winget.Source_8wekyb3d8bbwe\
```

把**整个文件夹**的内容复制到目标机器的 `Africa_DN_Testing\tools\bind\` 下
（确保 `tools\bind\dig.exe` 存在）。主程序和 `setup_env.ps1` 都会自动查找这个目录，无需安装、无需管理员。

**第 3 步：在新机器上直接双击 `run_dns_test.bat`**

如果那台机器上**连 Python 都没有**、又完全不能上网，请在能上网的机器上用方式 C 打包 exe。

---

## 三、方式 C：打包成 exe（目标机器无需 Python）

思路：把 Python 程序和依赖打包成一个 exe，dig 一起带走。

```powershell
# 在能上网的机器上执行一次
python -m pip install pyinstaller
python -m pyinstaller --onefile --name domain_dns_test `
       --add-data ".;." domain_dns_test.py
```

- 生成 `dist\domain_dns_test.exe`。
- **dig 不会被自动打包**：把 `tools\bind\`（便携版 dig 目录）和 exe、`非洲域名测试列表.xlsx`
  放在同一个文件夹里一起拷过去，主程序会自动找到 `tools\bind\dig.exe`。
- 运行：`domain_dns_test.exe --limit 5 --dry-run`（参数与脚本版一致）。

> 注意：`--onefile` 打包的 exe 首次启动会解压依赖，稍慢；如果目标机器装杀毒软件，
> 未签名的 exe 可能被拦截，需加白名单。

---

## 四、常见问题

| 现象 | 原因 / 解决办法 |
| --- | --- |
| `'powershell.exe' is not recognized as an internal or external command` | 目标机器的 `%PATH%` 里少了 `…\WindowsPowerShell\v1.0`（或被改坏）。现已修复：`run_dns_test.bat` 会先把 `Windows\System32` 等目录补回 `%PATH%`，再用绝对路径调用 PowerShell，不再依赖环境变量 |
| 运行时提示 `The system cannot find the path specified.`，最后 `[FAILED] exit code 3` | 旧版本会误用 `cache\python_path.txt` 里**别的电脑**的解释器路径。现已修复：缓存文件带计算机名，且使用前会校验“文件存在 + 能 import openpyxl, requests”，不符合就自动忽略并重新检测 |
| `[WARN] setup_env.ps1 returned . Trying to continue anyway.` | 旧版本在 `if (...)` 块里读 `%errorlevel%`，被 cmd 提前展开成空值。现已改为子过程调用，能拿到真实退出码 |
| `winget` 找不到 | Windows 10 1809 以下或未安装「应用安装程序」。手动装 Python：到 <https://www.python.org/downloads/windows/> 下载 `python-3.x.y-amd64.exe`，安装时勾选 **Add Python to PATH**；或用 `python-3.x.y-amd64.exe /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1`。 |
| pip 装依赖很慢 / 超时 | 脚本会自动重试清华、阿里云镜像。也可手动指定：`powershell -File scripts\setup_env.ps1 -IndexUrl https://pypi.tuna.tsinghua.edu.cn/simple`；公司代理环境可先设置 `set HTTPS_PROXY=http://<代理>:<端口>` 再运行。 |
| 提示 `[FAIL] dig NOT ready` | 没有管理员权限装不了 BIND。按方式 B 准备 `tools\bind\` 便携版即可。 |
| 脚本一闪而过 / PowerShell 报执行策略错误 | `run_dns_test.bat` 已用 `-ExecutionPolicy Bypass` 启动 PowerShell，正常不会遇到；若被组策略强制限制，请用管理员执行一次 `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`。 |
| 报错 `无法写入 ...xlsx` | 目标文件正被 Excel 打开。关闭 Excel，或改用 `-o` 指定另一个输出文件名。 |
| 结果里某个 DNS 服务器那一列是空的 | 说明这台机器的网络访问不到那个 DNS 服务器（例如只能从 CMI 非洲办公网查询的智能 DNS），程序按设计跳过了，不是程序错误。 |
| 中文显示乱码 | 请用 `run_dns_test.bat` 启动（会自动设置好编码）；在 PowerShell 里手动跑程序时中文提示是正常的，只有把输出重定向进文件才可能出现编码问题。 |

---

## 五、部署后快速验证清单

0. 确认目标机器上**没有**从原机器拷来的 `.venv\`、`cache\`（有 `tools\bind\` 则保留）。
1. `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup_env.ps1 -DryRun`
   → 应看到 Python / 依赖 / dig 三项都 `OK`（或给出将要安装的计划）。
2. `run_dns_test.bat --limit 2 --dry-run`
   → 能看到 2 个域名在 3 个 DNS 服务器上的解析预览。
3. `python -m unittest discover -s tests`（若目标机有完整项目文件）
   → 31 个单元测试通过。
4. `run_dns_test.bat` 全量运行
   → 生成 `非洲域名测试列表_result.xlsx`，`tests\verify_result.py` 校验通过。

---

## 六、常用运行命令速查

> 前提：先 `cd` 到项目文件夹（或者直接用完整路径调用 bat，bat 内部会自动切到自己的目录）。

**① 什么都不用记：双击 `run_dns_test.bat`**（首次运行会自动装环境，之后直接跑）。

**② 命令行方式（推荐用这一条）**

```cmd
:: cmd / 双击打开的命令提示符
cd /d "D:\Tech Doc\AI工具\VScode Program\Africa_DN_Testing"
run_dns_test.bat
```

```powershell
# PowerShell（注意前面要加 .\）
cd "D:\Tech Doc\AI工具\VScode Program\Africa_DN_Testing"
.\run_dns_test.bat
```

**③ 带参数运行**（参数与 `domain_dns_test.py` 完全一致，原样透传）

```cmd
run_dns_test.bat --limit 5 --dry-run     :: 只查前 5 个域名，只预览不写文件
run_dns_test.bat --limit 2               :: 只查前 2 个域名，会写结果文件
run_dns_test.bat                          :: 全量运行，写入 非洲域名测试列表_result.xlsx
run_dns_test.bat --inplace                :: 直接回写原文件（先关掉 Excel）
run_dns_test.bat -o D:\out\结果.xlsx       :: 指定输出文件
run_dns_test.bat --workers 4              :: 并发线程数（默认 8）
```

也可以从别的目录直接调用（bat 会自己切换工作目录）：

```cmd
"D:\Tech Doc\AI工具\VScode Program\Africa_DN_Testing\run_dns_test.bat" --limit 5 --dry-run
```

**④ 只想装环境、先不跑程序**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup_env.ps1            # 检查并安装
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\setup_env.ps1 -DryRun    # 只看缺什么，什么都不装
```

**⑤ 环境已就绪时，跳过 bat 直接跑 Python**

```cmd
.venv\Scripts\python.exe domain_dns_test.py --limit 5 --dry-run
```

（`cache\python_path.txt` 里记录的是上次使用的解释器路径与写入它的计算机名，可用 `type cache\python_path.txt` 查看；
换了电脑或解释器已不存在时，bat 会自动忽略这份缓存并重新检测。如果依赖装在系统 Python 里，也可以直接用 `python domain_dns_test.py`。）

**⑥ 辅助命令**

```cmd
.venv\Scripts\python.exe -m unittest discover -s tests        :: 单元测试
.venv\Scripts\python.exe tests\verify_result.py               :: 校验结果文件
.venv\Scripts\python.exe tests\list_duplicate_domains.py      :: 列出重复域名
.venv\Scripts\python.exe tests\remove_duplicate_domains.py    :: 演练删除重复行（加 --apply 才真删）
```

---

## 七、`setup_env.ps1` 参数一览

| 参数 | 说明 |
| --- | --- |
| `-DryRun` | 只检查并打印将要执行的动作，不安装、不写文件 |
| `-NoInstall` | 只检查，不安装（缺什么就报什么） |
| `-FakeMissingPython` / `-FakeMissingDeps` / `-FakeMissingDig` | 假装某个组件缺失，配合 `-DryRun` 用来验证自动安装分支 |
| `-VenvPath <路径>` | 指定虚拟环境位置（默认 `<项目>\.venv`） |
| `-IndexUrl <地址>` | 指定 pip 镜像（默认 PyPI，失败后自动试清华/阿里云） |

> 备注：脚本里的提示文字刻意使用英文纯 ASCII。Windows PowerShell 5.1 在没有
> UTF-8 BOM 时按 ANSI 读取 `.ps1`，中文会变乱码；同样地 `cmd.exe` 读取含中文的
> `.bat` 也可能错位。中文交互全部由 `domain_dns_test.py` 输出，因此不影响使用。
