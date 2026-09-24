# 非洲域名测试（Linux 版）

从 `非洲域名测试列表.xlsx` 读取域名，对表头上一行里的每个 DNS 服务器执行

```
dig @<DNS服务器IP> <域名> +short
```

把 **CNAME 链路**写入「CNAME」列、**A 记录 IP** 写入「A」列（每个 IP 一行），
用 <https://ip-api.com/> 把每个 IP 的运营商简化成 **CT / CU / CM** 写入「IP归属」列，
对每个 IP 做 ping 把**平均时延(ms) 与丢包率(%)** 写入「时延」「丢包率」列，
数据区全部处理完后重算每个 DNS 块下方的**表尾统计**。

本目录是 **Linux 版本**；Windows 版本在仓库的 [`../windows/`](../windows/) 下，
两个版本功能、参数、统计口径完全一致（差异见根目录 [README.md](../README.md)）。

## 1. 目录内容

| 文件 | 说明 |
| --- | --- |
| `domain_dns_test.py` | 主程序（Python 3.9+） |
| `run_dns_test.sh` | 一键运行：缺 Python / 依赖 / dig 时会自动调用 `scripts/setup_env.sh` 装上 |
| `scripts/setup_env.sh` | 环境自动部署脚本（apt / dnf / yum / apk / zypper / pacman 自动识别），可单独运行 |
| `tests/test_domain_dns_test.py` | 单元测试（unittest，87 个用例） |
| `tests/verify_result.py` | 结果文件校验脚本（列格式、域名块对应、改善列、表尾统计重算比对；两种格式都支持） |
| `tests/list_duplicate_domains.py` | 列出「域名」列中的重复域名及其行号/来源 |
| `tests/remove_duplicate_domains.py` | 删除重复域名的后续出现行（默认只演练，`--apply` 才写入并自动备份） |
| `requirements.txt` | 依赖清单（openpyxl、requests） |
| `tools/bind/` | 可选：把便携版 `dig` 放这里实现免安装（见该目录下的 README.txt） |
| `cache/` | 运行后自动生成：IP 归属地缓存、dig 缓存、ping 结果缓存、已解析的解释器路径（`python_path.txt`，含本机主机名，**属于单机缓存，不要拷给别的机器**） |
| `backup/` | 自动删除重复域名行之前，原文件的带时间戳备份 |
| `*_result.xlsx` | 运行后自动生成的带结果文件 |

> 脚本里的提示文字刻意使用英文 ASCII：这样 `LANG`/`LC_ALL` 未设置（POSIX locale）时
> 也不会出现乱码。中文提示一律由 Python 程序输出。
>
> **拷到其他机器时，`.venv/` 和 `cache/` 两个目录一定不要带**（它们写死了本机的绝对路径）：
> `.venv` 会在新机器上自动重建；`cache/python_path.txt` 绑定了主机名，不符时程序会自动忽略。

## 2. 环境准备

**推荐做法：什么都不用准备，直接运行 `./run_dns_test.sh`** —— 脚本会自动检测并安装
Python 3、pip 依赖和 `dig`，装完自动开始运行（装系统包时需要 root 或 sudo，
否则会打印出需要你手动执行的命令）。

手动准备（可选）：

1. **Python 3.9+**（本机验证环境为 3.12），并确保 `venv` 可用：

   ```bash
   sudo apt-get install -y python3 python3-venv python3-pip    # Debian/Ubuntu
   sudo dnf install -y python3 python3-pip                      # RHEL/CentOS/Fedora
   sudo apk add python3 py3-pip py3-virtualenv                  # Alpine
   ```

2. **dig 命令**（`bind9-dnsutils` / `dnsutils` / `bind-utils` / `bind-tools`）：

   ```bash
   sudo apt-get install -y dnsutils        # Debian ≤12 / Ubuntu ≤23.10
   sudo apt-get install -y bind9-dnsutils  # Debian 13 / Ubuntu 24.04+
   sudo dnf install -y bind-utils          # RHEL/CentOS/Fedora
   sudo apk add bind-tools                 # Alpine
   ```

   程序会自动在 `PATH`、`tools/bind/`、`/usr/bin`、`/usr/local/bin`、`/snap/bin`
   中查找 `dig`；也可以用 `--dig /path/to/dig` 手动指定。

3. **依赖**：

   ```bash
   python3 -m venv .venv
   .venv/bin/python -m pip install -r requirements.txt
   ```

### ping 权限（重要）

Linux 的 `ping` 需要发送原始 ICMP，某些系统上普通用户没有权限，此时
「时延」「丢包率」会变成 `-` 或空，程序仍会正常跑完。一次性修复方式二选一：

```bash
sudo setcap cap_net_raw+ep "$(command -v ping)"          # 给 ping 加能力位
sudo sysctl -w net.ipv4.ping_group_range='0 2147483647'  # 放开 ICMP socket 组范围
```

`scripts/setup_env.sh` 启动时会自动检查并提示这个问题。

## 3. 使用方式

```bash
./run_dns_test.sh                        # 全量运行，写出 非洲域名测试列表_result.xlsx
./run_dns_test.sh --limit 5 --dry-run    # 先看 5 个域名的效果（不写文件）
./run_dns_test.sh --no-ping              # 只做 dig + 运营商简化
./run_dns_test.sh --inplace              # 直接回写原文件
```

参数会原样传给 `domain_dns_test.py`；也可以直接调用 Python：

```bash
.venv/bin/python domain_dns_test.py --limit 5 --dry-run
```

| 参数 | 说明 |
| --- | --- |
| `-i/--input` | 输入 Excel，默认 `非洲域名测试列表.xlsx` |
| `-o/--output` | 输出 Excel，默认在输入文件名后加 `_result` |
| `--inplace` | 直接回写原文件（建议先备份；被占用/无权限时报错退出） |
| `--limit N` | 只处理前 N 个域名，用于试跑（默认不写表尾统计） |
| `--dry-run` | 只打印结果预览，不写 Excel |
| `--workers N` | dig 并发线程数，默认 8；设为 1 表示逐个串行 |
| `--sheet NAME` | 指定工作表，默认第一个 |
| `--dig PATH` | 指定 `dig` 可执行文件路径 |
| `--no-ip-info` | 不做归属地查询（IP归属/首IP归属 留空，同时跳过表尾统计） |
| `--no-ip-cache` | 不使用 IP 归属地本地缓存 |
| `--use-dig-cache` | 复用上次 dig 结果（试跑更快，但数据非实时） |
| `--no-ping` | 跳过 ping 测试（不写时延/丢包率） |
| `--ping-count N` | 每个 IP 发送的 ping 报文数，默认 10 |
| `--ping-timeout MS` | 单个报文的等待超时（毫秒），默认 4000；Linux 下换算成整数秒传给 `ping -W` |
| `--ping-interval S` | 报文间隔秒数（`ping -i`），默认 1（与 Windows 版一致，便于对拍）；试跑时可调小加速，非 root 下最小 0.2 |
| `--ping-workers N` | ping 并发数，默认 16 |
| `--no-ping-cache` | 不使用 ping 结果缓存（每次都真实测试） |
| `--no-ping-retry` | 丢包率不为 0% 时不重测（默认会重测一次并按第二次结果记录） |
| `--no-autofit-row-height` | 不按内容自动调整行高（默认会自动调整） |
| `--no-local-dns-rewrite` | 不把「DNS 块填的是本机公网 IP」自动改写成 `127.0.0.1` 查询（默认会自动改写） |
| `--local-ip IP` | 手动指定本机出口公网 IP（跳过自动探测） |
| `--split-ip-rows` | 每个 IP 占一行：解析出多个 IP 时按 IP 数拆行（见第 4 节），**默认开启** |
| `--no-split-ip-rows` | 关闭拆行，恢复成「一个域名一行」（每格多行、靠换行对齐） |
| `--keep-duplicates` | 保留重复域名行（默认自动删除多余行并备份） |
| `--force-stats` | 配合 `--limit` 使用时也写入表尾统计 |

> `--limit` 试跑时默认**不写表尾统计**（半量数据会误导读表的人），需要时可加 `--force-stats`。

## 4. 表格结构与写入规则

程序不依赖硬编码的行列号，会按下面的规则自动探测：

- 「域名」列：表头行及以上、文字为 `域名` 的列（本例为 `C` 列，标题跨多行合并也能识别）。
- 表头行：含 `CNAME` 表头**最多的那一行**（本例第 6 行）。
- DNS 服务器 IP：表头行的**上一行**（本例第 5 行）；每个块取块内第一个合法 IP，
  写在合并单元格里也能识别。
- 数据区：表头行下一行 ~ 第一个统计标签（`无效解析行数`）所在行的上一行（本例第 7~182 行）。

本表的实际布局：

| DNS 服务器（第 5 行） | CNAME | A | IP归属 | 时延 | 丢包率 | 首IP归属 | 时延改善 | 丢包改善 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `8.8.8.8`（公共DNS，基准块） | `D` | `E` | `F` | `G` | `H` | `I` | - | - |
| `223.119.133.15`（中国移动智能DNS） | `J` | `K` | `L` | `M` | `N` | `O` | `P` | `Q` |
| `118.145.108.65`（预留对比用） | `R` | `S` | `T` | `U` | `V` | `W` | - | - |

单元格内容示例（某域名在 `8.8.8.8` 下解析出 3 个 IP）：

```
E7: 113.200.1.133
    113.200.2.218
    119.249.103.3

F7: CU          ← IP归属，与 A 列逐行对应
    CU
    CU

G7: 55          ← 时延（整数毫秒）
    51
    53

H7: 0%          ← 丢包率（整数百分比）
    0%
    33%

I7: CU          ← 首IP归属（只看第一个 IP）
```

规则说明：

- `dig +short` 输出中，**能解析成 IP 的行算 A 记录**，其余行算 CNAME 记录
  （结尾的 `.` 会被去掉）；以 `;` 开头的诊断行不会写入单元格。
- `A` 列每行一个纯 IP；`IP归属` / `时延` / `丢包率` 与 `A` 列**逐行一一对应**。
- **该 DNS 下没有解析出任何 IP**（超时/报错/只有 CNAME 没有 A）时，这一行的
  CNAME/A/IP归属/时延/丢包率/首IP归属/改善列全部留空。
- 同一个 IP 只查询一次归属地、只 ping 一次，但结果会写入它出现的每个格子。
- 含中文的域名会自动转成 punycode 后再查询。
- 表内重复域名只保留第一次出现，多余的行会被自动删除（见第 8 节）。
- **两个块即使配了同一个 DNS 服务器 IP，也各自独立写入与统计**（不会互相覆盖）。
- **DNS 块填的是本机自己的公网 IP 时**，程序会自动改用 `127.0.0.1` 去执行 dig
  （云主机通常无法用自身公网 IP 回环访问自己，直接查会超时）；
  **结果表里的 DNS 服务器 IP 保持原样不变**。想关掉用 `--no-local-dns-rewrite`，
  或手动指定本机公网 IP：`--local-ip <IP>`。
- **默认是「每个 IP 一行」（`--split-ip-rows`）**：把解析出多个 IP 的域名拆成多行：
  - **每域一列**（序号 / 域名等前置列 、CNAME、首IP归属、时延改善、丢包改善）
    纵向**合并**成一块，值写在合并区首行；
  - **每个 IP 一列**（A / IP归属 / 时延 / 丢包率）逐行写入，行数不足时留空。
  - 三个块的 IP 数不同时按**最大值**拆行（例如 5/5/20 就拆 20 行，前两块后 15 行空白）。
  - 这样每格只有一个 IP 的归属文字，不靠换行对齐；归属列不开自动换行，
    长文字保持原列宽、超出的部分被裁切。
  - ⚠️ **表尾统计仍按「域名」口径计算**（一个域名的多行算一行），11 项含义不变；
    统计区会随插入的行整体下移。
  - ⚠️ 行数会明显变多（实测 20 个域名 -> 139 行；全量 175 个域名预计 1000+ 行）。
    `tests/verify_result.py` 已支持两种格式（按「域名块」统计），可直接校验。
- 用 **`--no-split-ip-rows`** 可切回「一个域名一行」：此时 `IP归属` / `首IP归属` 列
  每个 IP 占一行（靠单元格里的显式换行，必须保持「自动换行」开启，
  Excel 才会渲染成多行）；这两列只按**显式换行**参与行高估算，
  很长的运营商文字不会把行高撞高，列宽也保持不变（见第 5.4 节）。

## 5. 运营商简化 / ping / 改善值

### 5.1 运营商简化（写入 `IP归属`、`首IP归属`）

判定范围是 ip-api 返回的 **`isp`、`org`、`as`、`asname` 四个字段**：
按 `isp` → `org` → `as` → `asname` 的顺序依次检查，**任意一个**字段包含下表关键字
就简化成 `CT` / `CU` / `CM`（取第一个命中的）。
不区分大小写，先把文字里的空白与标点去掉再做包含判断：

| 上述四个字段中任一字段包含 | 写入 |
| --- | --- |
| `chinanet`、`china telecom`、`wanbao` | `CT` |
| `china unicom`、`china169` | `CU` |
| `china mobile communications corporation`、`china mobile communications group` | `CM` |
| 四个字段都不含（如 `Hangzhou Alibaba Advertising Co`） | 原样写 `isp` 字段的文字 |
| 查询失败 / `isp` 为空 | `未知` |

- 之所以要看四个字段：ip-api 的 `isp` 字段偶发脏值，会把城市名当运营商返回
  （实测出现过 `Jinan,` / `Qingdao,`，且字段之间还会串），
  但同一条响应里的 `org`（`Chinanet SD`）与 `as` / `asname`
  （`AS58540 CHINATELECOM SHANDONG JINAN IDC`）是正确的 —— 只看 `isp` 会把这类
  电信 IP 误判成「非三大运营商」。
- 四个字段都没命中时**仍写 `isp` 原文**（与旧版行为一致）；简化只影响写入，
  `cache/ip_info_cache.json` 里保存四个字段的原始文字。
- `China Mobile Hong Kong Company Limited`（CMHK）不在 CM 规则内，会原样保留；
  如果也要算 CM，改 `ISP_SIMPLIFY_RULES` 一行即可。
- **升级提示**：旧版缓存的记录没有 `org`/`as`/`asname` 三个字段，升级后请删除
  `cache/ip_info_cache.json`（或加 `--no-ip-cache` 运行），
  否则这些 IP 会走缓存命中、继续沿用旧结果，看起来像「改了没生效」。

归属地数据来自 <https://ip-api.com/>（免费版，字段
`status,country,regionName,city,isp,org,as,asname`，语言 `zh-CN`）：
批量接口 15 次/分钟（每次最多 100 个 IP），单条 45 次/分钟，**只支持 HTTP**；
程序内置限速与重试，批量失败会自动降级为单条查询，结果缓存到 `cache/ip_info_cache.json`。

### 5.2 ping 测试（写入 `时延`、`丢包率`）

- 命令：`ping -c <报文数> -W <单包超时秒> -i <间隔秒> <IP>`，
  默认 10 个报文、单包超时 4000ms（换算成 `-W 4`）、间隔 1 秒。
- 解析 Linux/iputils 摘要：`rtt min/avg/max/mdev = 22.100/23.456/25.000/1.234 ms`
  取 `avg`；`10 packets transmitted, 10 received, 0% packet loss` 取丢包率。
  BusyBox 的 `round-trip min/avg/max = …` 写法同样支持；**摘要行缺失时**退回用每个
  回包的 `time=23.4 ms` 求平均，Windows 版的中英文摘要也能解析（便于对拍）。
- `时延` 写整数毫秒（四舍五入）；`丢包率` 写整数百分比（如 `0%`）。
- 全部丢包或测不到平均时延时，`时延` 写 `-`（便于区分「测了但全丢」与「没测」）。
- **丢包率不为 0% 时会自动重测一次**（同样的报文数、超时与间隔），并且**不论第二次结果
  如何都按第二次记录**（包括第二次仍然丢包、甚至全丢包的情况）；第一次没测出结果
  （超时/无法解析/权限不足）也会重测。控制台会列出哪些 IP 被重测及两次的对比；
  想关掉用 `--no-ping-retry`。
- 结果缓存在 `cache/ping_cache.json`（保存的是重测后的最终结果，并记录报文数；
  报文数不一致的旧缓存会自动作废）。需要重新实测时加 `--no-ping-cache`。
- 权限不足时不会静默失败：错误信息里会直接给出 `setcap` / `ping_group_range` 的修复命令。

### 5.3 时延改善 / 丢包改善

- 只有带这两个表头的块（中国移动智能DNS）才会写；对应「公共DNS」块作为基准。
- 条件：该行在智能DNS 块里解析出的**首IP** 的简化归属为 `CM`。
- 取值：`智能DNS 首IP 的值 − 公共DNS 块首IP 的值`（时延为整数毫秒、丢包为整数百分比，
  负值表示智能DNS 更快/更少丢包）。
- 任一侧没有解析结果或没有 ping 数据时留空。

### 5.4 行高自适应

- 写入完成后，程序会按每行内容估算需要多少行文字，自动设置该行行高，
  使单元格里的多行内容（多个 IP / 多个归属 / 多行时延）不需手工拖动就能完整显示。
- 估算方式：按列宽把文本换行后的总行数 × 15 磅 + 2 磅留白（全角字符按 2 个半角宽计），
  上限为 Excel 允许的 409 磅。
  **`IP归属` / `首IP归属` 列只按单元格里的显式换行计数**（不按列宽折算），
  所以很长的运营商文字不会把行高撞高，列宽也保持不变。
- 想保留自己的行高设置：加 `--no-autofit-row-height`。

## 6. 表尾统计

每个 DNS 块下方有 11 项统计：**标签写在块内 `A` 列，数值写在标签右边一格**
（本表为 `F` / `L` / `T` 列）。程序运行时会**先清空这些数值**，等所有域名都解析完再统一计算填写。

| 统计项 | 口径 |
| --- | --- |
| 无效解析行数 | 该 DNS 下没有解析出任何 IP 的行数 |
| 非三大运营商行数 | 所有 IP 的简化归属里都不含 CT/CU/CM 的行数（含「未知」） |
| 三大运营商行数 | 至少有一个 IP 属于 CT/CU/CM 的行数 |
| 仅CT / 仅CU / 仅CM 行数 | 只看三大运营商结果，全部为单一家 |
| CT+CU / CT+CM / CU+CM / CT+CU+CM 行数 | 只看三大运营商结果，同时包含对应几家、且不含其他家 |
| 首IP归属为CM行数 | `首IP归属` 列填 `CM` 的行数（同一行还有其他归属也算） |

- 加和关系：`无效解析 + 非三大 + 三大 = 数据行数`；七个组合之和 `= 三大运营商行数`。
- 非三大运营商的行不会计入任何组合；「未知」算非三大。
- 重复域名行会被删除（见第 8 节），所以统计不会把同一域名数两遍。

## 7. 单元测试与结果校验

```bash
cd linux
.venv/bin/python -m unittest discover -s tests     # 100 个用例
.venv/bin/python tests/verify_result.py 非洲域名测试列表_result.xlsx
```

单元测试覆盖：`dig +short` 输出解析（CNAME/IP/诊断行/IPv6/去重）、失败原因分类、
dig 路径查找、表结构自动探测（真实表 + 合成表）、运营商文字简化（含 `isp`/`org`/`as`/`asname`
四字段判定）、本机公网 IP 改写、ping 输出解析（Linux iputils / BusyBox / Windows 中英文 /
全丢包 / 部分丢包 / 退回逐包求平均）、Linux ping 命令拼装（`-W` 秒换算、`-i` 间隔）、
ping 失败原因提示（权限/不可达）、丢包重测策略、ping 缓存（含报文数校验）、
统计分类与加和关系、时延改善计算、行高自适应、**每个 IP 一行的拆行与域名块分组**、
重复域名删除与备份、单元格写入、命令行参数解析。

校验脚本会重新读取结果文件并独立复核（**两种写入格式都支持**）：A 列必须是纯 IP、
IP归属/时延/丢包率必须与 A 列的 IP 一一对应且格式正确、首IP归属必须等于该域名块的
第一个归属、改善列只在首IP归属为 CM 的域名块出现、**表尾统计与按表内实际内容重算的结果
完全一致**、重复域名为 0、表头与 DNS 服务器 IP 行未被改动。

> 校验脚本按**域名块**统计：拆行格式下同一个域名的多行会合并成一个块来算，口径与程序一致；
> 用 `--limit` 试跑时，数据区末尾没处理到的空白域名块会自动排除
> （否则「无效解析行数」会虚高）。

## 8. 重复域名处理

- 主程序默认**自动删除**重复域名的多余行（保留第一次出现），删除前把原文件
  带时间戳备份到 `backup/`，控制台会打印删除了哪一行。
- 需要保留重复行时加 `--keep-duplicates`（只警告，不删除）。
- 只想查看/单独处理重复：

  ```bash
  .venv/bin/python tests/list_duplicate_domains.py              # 列出重复域名、行号与来源
  .venv/bin/python tests/remove_duplicate_domains.py            # 演练：只列出将删除哪些行
  .venv/bin/python tests/remove_duplicate_domains.py --apply    # 实际删除并自动备份原文件
  ```

- 删除行后表尾统计区会整体上移；程序按标签文字动态定位，不会错位。
- 删除后「序号」列会保留原编号并出现空缺；如需重新连续编号，改脚本或另行处理。

## 9. 常见问题

| 现象 | 原因与处理 |
| --- | --- |
| `未找到 dig 命令` | 安装 `dnsutils`/`bind9-dnsutils`/`bind-utils`/`bind-tools`，或把 `dig` 放到 `tools/bind/`，或用 `--dig` 指定 |
| 「时延」全是 `-`、丢包率 100%，或错误里提到 `CAP_NET_RAW` | 当前用户没有发 ICMP 的权限，按第 2 节的 `setcap` / `ping_group_range` 处理 |
| `No module named openpyxl` | 运行 `bash scripts/setup_env.sh` 重新准备环境；直接调用时请用 `.venv/bin/python` |
| 依赖装不上（内网/离线） | 把匹配本机 Python 版本与架构的 `.whl` 放进 `wheels/`，`setup_env.sh` 会优先离线安装 |
| 出现乱码 | 终端请使用 UTF-8 locale（`export LANG=C.UTF-8`），或者 `export PYTHONIOENCODING=utf-8` 后再运行 |
| 从 Windows 拷过来的 `cache/`、`.venv/` | 直接删掉；`python_path.txt` 因主机名不符会被忽略，`.venv` 会被自动重建 |
| 表里 IP/行数被我改过 | 不用改代码：表头行、DNS IP 行、数据区、统计位置以及块内各列都是自动探测的 |
