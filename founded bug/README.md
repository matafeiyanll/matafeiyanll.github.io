# DBMS Bug List 维护手册

这个目录包含一个从 Excel 加载数据的 DBMS Bug List 网页。每个研究方法对应一份独立的 Excel，例如 `SQLess.xlsx`、`SchemaMorph.xlsx` 和 `FMU.xlsx`。

以后修改数据时，请以 `data/` 目录里的 Excel 为数据源。不要手工编辑 `data/bug-data.js`，它应该通过更新脚本自动生成。

## 最常用的操作

### 修改已有方法的数据

1. 打开 `data/` 目录中对应的 Excel，例如 `data/SchemaMorph.xlsx`。
2. 修改数据并保存 Excel。
3. 双击 `update-bug-data.cmd`。
4. 看到 `Updated data/bug-data.js with ... records.` 表示生成成功。
5. 刷新网页检查结果。
6. 如果网站使用 GitHub Pages，还需要提交并推送修改后的 Excel、`data/bug-data.js` 和 `data/workbooks.json`。

### 新增一个研究方法

以新增 `FMU` 为例：

1. 准备 `data/FMU.xlsx`。
2. 检查 Excel 的列名和字段值是否符合下方规范。
3. 在 `data/workbooks.json` 中登记 `FMU.xlsx`。
4. 在 `app.js` 顶部的默认清单中登记 `FMU.xlsx`。
5. 双击 `update-bug-data.cmd`。
6. 刷新网页，检查 FMU 标签、数量、筛选和下载链接。
7. 发布时提交 Excel、`workbooks.json`、`bug-data.js` 和 `app.js`。

## 目录和文件作用

| 文件 | 作用 | 是否手工修改 |
| --- | --- | --- |
| `data/<Method>.xlsx` | 每个研究方法的 Bug 数据源 | 是 |
| `data/workbooks.json` | 登记网页需要加载的 Excel | 新增或删除方法时修改 |
| `data/bug-data.js` | 网页无法直接读取 Excel 时使用的本地数据快照 | 否，由脚本生成 |
| `app.js` | 页面筛选、Excel 加载和默认方法清单 | 新增或删除方法时修改 |
| `index.html` | 页面结构 | 一般不需要修改 |
| `styles.css` | 页面样式 | 一般不需要修改 |
| `update-bug-data.cmd` | Windows 下的一键更新入口 | 否 |
| `update-bug-data.py` | 从所有 Excel 重新生成网页数据 | 否 |

## Excel 格式规范

每个 Excel 使用第一张工作表，建议将工作表命名为 `Bugs`。第一行必须包含以下 7 个列名，名称和空格都不能修改：

| Program | ID | Bug ID | Bug Link | Bug Type | Status | Method |
| --- | --- | --- | --- | --- | --- | --- |

字段要求：

- `Program`：数据库名称，例如 `MySQL`、`MariaDB`、`PostgreSQL`、`TiDB`。
- `ID`：Excel 内的原始编号，建议在每份工作簿中从 1 连续编号。
- `Bug ID`：Bug 的公开编号，例如 `#120158`、`MDEV-39553`。
- `Bug Link`：完整的 `https://` 链接。
- `Bug Type`：建议统一使用 `Logical Bug`、`Crash` 或 `Simplification`。
- `Status`：只使用 `Confirmed`、`Fixed`、`Waiting`、`Duplicated` 或 `Not a bug`。
- `Method`：当前研究方法名称，例如 FMU 的每一行都填写 `FMU`。

注意事项：

- 不要修改表头拼写。
- 不要在字段末尾留下空格，例如应写 `Fixed`，不能写 `Fixed `。
- 不要混用 `Crash` 和 `Crash Bug`，否则网页会显示两个不同的筛选项。
- `Program` 和 `Bug ID` 为空的行不会被导入。
- 网页第一列的 `No.` 是筛选后动态生成的，不使用 Excel 的 `ID` 作为网页序号。

## 修改已有 Excel 的完整流程

假设需要修改 SchemaMorph：

1. 修改 `data/SchemaMorph.xlsx`，不要修改 Downloads 或 Desktop 中的原始文件。
2. 保存并关闭 Excel，避免文件尚未写入完成或被锁定。
3. 双击 `update-bug-data.cmd`。
4. 脚本会读取 `data/workbooks.json` 中登记的所有 Excel。
5. 脚本会自动更新：
   - `data/bug-data.js`
   - `data/workbooks.json` 中每个方法的 `count`
6. 刷新网页确认状态、总数和筛选结果。

不需要手工修改 `index.html` 中的总数，页面运行后会根据实际数据自动计算。

## 新增方法的完整流程

### 1. 创建 Excel

可以复制一份现有的标准 Excel，删除旧数据后另存为：

```text
data/FMU.xlsx
```

填写数据时保持 7 个标准表头，并确保每行的 `Method` 都是 `FMU`。

### 2. 修改 workbooks.json

打开 `data/workbooks.json`，在数组最后添加：

```json
{
  "method": "FMU",
  "file": "FMU.xlsx",
  "count": 0
}
```

如果前面还有其他对象，记得在前一个对象结尾添加逗号。`count` 可以先写 `0`，更新脚本会自动替换成实际数量。

文件名必须和磁盘上的文件名完全一致。GitHub Pages 区分大小写，并且不能漏掉 `.xlsx`。

### 3. 修改 app.js

在 `app.js` 顶部的 `manifest` 数组中添加：

```js
{ method: "FMU", file: "FMU.xlsx" },
```

这里的文件名同样必须包含 `.xlsx`。

### 4. 生成网页数据

双击：

```text
update-bug-data.cmd
```

也可以在终端进入 `founded bug` 后运行：

```powershell
py -3 update-bug-data.py
```

如果电脑没有 `py` 命令，也可以使用：

```powershell
python update-bug-data.py
```

成功输出示例：

```text
FMU: 128 records
Updated data/bug-data.js with 392 records.
```

### 5. 检查页面

检查以下内容：

- 页面顶部出现新方法标签。
- 新方法的数量正确。
- Program、Bug Type、Status 和 Method 均可筛选。
- Bug Link 可以打开。
- Excel 下载区域出现新文件。
- 页面总数等于所有方法记录数之和。

## 删除一个方法

1. 从 `data/workbooks.json` 删除对应对象。
2. 从 `app.js` 顶部默认清单删除对应项。
3. 双击 `update-bug-data.cmd` 重新生成数据。
4. 确认网页正常后，再决定是否删除对应的 Excel。

删除 Excel 前一定要先取消登记；否则更新脚本会提示找不到工作簿并停止。

## 本地查看和 GitHub Pages

### 直接双击 index.html

浏览器通常不允许通过 `file://` 读取旁边的 Excel。此时页面主要使用 `data/bug-data.js`，所以修改 Excel 后必须先运行 `update-bug-data.cmd`。

### GitHub Pages

GitHub Pages 会读取仓库中已经提交的文件。因此本地修改后，需要提交并推送这些文件：

- 修改过的 `data/*.xlsx`
- `data/bug-data.js`
- `data/workbooks.json`
- 新增或删除方法时还包括 `app.js`

推送完成后等待 GitHub Pages 部署，再刷新网页。

## 常见问题

### 修改 Excel 后网页没有变化

依次检查：

1. 修改的是不是 `founded bug/data` 中的 Excel。
2. Excel 是否已经保存。
3. 是否运行了 `update-bug-data.cmd`。
4. `bug-data.js` 的修改时间是否已经更新。
5. 使用 GitHub Pages 时，相关文件是否已经提交并推送。
6. GitHub Pages 是否已经完成部署。

### 新方法没有出现在页面上

检查：

- Excel 是否位于 `data/`。
- `workbooks.json` 是否已经登记。
- `app.js` 是否已经登记。
- 文件名是否包含 `.xlsx`。
- Excel 的 `Method` 列是否填写正确。
- 是否重新运行了更新脚本。

### 页面出现两个相似的筛选项

通常是 Excel 中存在拼写或空格差异，例如：

- `Fixed` 和 `Fixed `
- `Crash` 和 `Crash Bug`
- `MySQL` 和 `mysql`

统一 Excel 中的值后，再运行更新脚本。

### 更新脚本运行失败

- 确认安装了 Python 3。
- 保存并关闭正在编辑的 Excel。
- 检查 `workbooks.json` 是否是合法 JSON。
- 检查登记的 Excel 文件是否真实存在。
- 检查所有 Excel 是否包含 7 个标准表头。

脚本发生错误时不会使用不完整的数据覆盖网页数据；修复错误后重新运行即可。

## 当前方法的 Bug Type 约定

- SQLess：`Simplification`
- QTRAN：`Logical Bug`
- ValScope：`Logical Bug`
- SmartFuzz：`Crash`
- FPMT：`Logical Bug`
- DMLScope：`Logical Bug`
- SchemaMorph：`Logical Bug`
- FMU：可同时包含 `Logical Bug` 和 `Crash`

## 发布前检查清单

- [ ] Excel 位于 `data/`，文件名正确。
- [ ] 7 个表头完全一致。
- [ ] `Method` 列与方法名称一致。
- [ ] Bug Type 和 Status 拼写统一、没有多余空格。
- [ ] 新方法已经登记到 `workbooks.json` 和 `app.js`。
- [ ] 已运行 `update-bug-data.cmd`。
- [ ] `bug-data.js` 和记录数已经更新。
- [ ] 网页的标签、筛选、链接和下载均正常。
- [ ] GitHub Pages 所需文件已经提交并推送。
