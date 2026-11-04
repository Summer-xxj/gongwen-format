# DOCX 程序化核验与修复要点（WorkBuddy 本地环境实战）

> 来源：2026-10-16 两轮实战——某地区合作方案（方案类文档）、某企业邀请函（红头函件类）。适用环境：WorkBuddy + workbuddy-local-office-edit（editor_sdk）+ WPS。纯文本核验之外，本文记录 Word/XML 层容易踩的坑。

## 0. 先分文档类型（决定核验深度）

**判断链：标题定文种 → 文种问版式 → 版式定规则。** 标题含「关于××的函」即文种为函，须向用户确认是否按信函格式发（不能默认通用格式）——信函格式首页不显示页码、红头为双红武文线（详见 format-rules.md §9）。

- **方案/报告类**：无主送机关、无落款日期，重点是标题层级、正文格式、数字字体。基础格式问题少。
- **红头函件类**（发文红头 + 标题 + 主送 + 正文 + 落款 + 附件）：要素齐全，坑密集——间距模型、软回车、落款对称、附件另面、标题断行、页码三件套、红双线都必须逐项核。**此类文档按本文件 §7–§13 全套执行。**

## 1. 字体存在性核查：先查 WPS，再谈下载

**顺序铁律：先确认本机有没有，再决定是否下载。** 用户说「我 WPS 里明明有」时，字体几乎一定已在硬盘上，只是没注册到系统字体库。

### 1.1 目录清单（按平台）

| 平台 | 路径 |
|---|---|
| Windows 系统 | `C:\Windows\Fonts`、`C:\Users\<user>\AppData\Local\Microsoft\Windows\Fonts` |
| **Windows WPS** | `C:\Users\<user>\AppData\Roaming\kingsoft\office6\docerFonts\` |
| macOS 系统 | `/System/Library/Fonts`、`/Library/Fonts`、`~/Library/Fonts` |
| **macOS WPS** | `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/office6/` 下：`docerFonts/`（方正系）、`data/fonts/extend_font/`、`data/fonts/online_ext_font_cache/`（后两者内容重复） |

公文字体全家桶通常就在 WPS 目录：方正小标宋简体、方正公文小标宋、方正小标宋_GBK、方正仿宋_GB2312、方正楷体_GB2312、方正黑体_GBK。

### 1.2 关键教训

- 只查系统目录会误判「字体缺失」——实际教训：报告方正小标宋缺失，用户指出 WPS 里明明有。
- **macOS 尤其容易误判**：WPS 容器里的字体不会进系统字体库，Chrome / Pages / Keynote 一律看不到；但 WPS 自己渲染正常。
- DOCX 里声明的字体 ≠ 已安装字体；用户默认用 WPS 打开时，docerFonts 内字体渲染正常，**可以直接写入使用**，不必降级替代。
- 注册表查询（reg.exe）可能被沙箱拦截，用 `ls 目录 | grep 字体名` 遍历即可。

### 1.3 读真实家族名，别信文件名

文件名和 font-family 经常不一致（`方正小标宋_GBK.ttf` → 家族名 `FZXiaoBiaoSong-B05`；`仿宋_GB2312.ttf` → `FangSong_GB2312`）。用 pymupdf 直接读 name 表：

```python
import pymupdf
print(pymupdf.Font(fontfile=path).name)   # 真实家族名
print(pymupdf.Font(fontfile=path).glyph_count)  # 顺带看字数覆盖
```

macOS 常态：`FZXiaoBiaoSong-B05`（方正小标宋_GBK，22034 字）有、`FangSong_GB2312`（7580 字）有、`HYKaiTiKW`（汉仪楷体）有；**楷体_GB2312 / 黑体_GB2312 通常没有**，用汉仪楷体 / 华文黑体顶替并声明。

### 1.4 不装系统也能用：CSS @font-face 直读 ttf

只需生成 PDF、不想改动用户字体库时的最优解——零安装、零 base64：

```css
@font-face { font-family:"GWTitle"; src:url("file:///…/方正小标宋_GBK.ttf"); }
@font-face { font-family:"GWBody";  src:url("file:///…/仿宋_GB2312.ttf"); }
```

```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless=new --no-sandbox --allow-file-access-from-files --no-pdf-header-footer \
  --run-all-compositor-stages-before-draw --virtual-time-budget=20000 \
  --print-to-pdf=out.pdf "file://…/template.html"
```

必须带 `--allow-file-access-from-files`，否则 file:// 页面加载 file:// 字体被 CORS 拦掉、静默回退。
路径含中文要 `urllib.parse.quote` 编码。

**数字/英文用 Times New Roman 的实现**（国标通行惯例：西文与相邻中文同字号）：利用逐字回退，把 Times 放在字体栈第一位——它没有 CJK 字形，中文自动落到下一个：

```css
@font-face { font-family:"GWTimes"; src:url("file:///System/Library/Fonts/Supplemental/Times%20New%20Roman.ttf"); }
body    { font-family:"GWTimes","GWBody";  }   /* 数字/英文→Times，中文→仿宋_GB2312 */
h1      { font-family:"GWTimes","GWTitle"; }
p.lead  { font-family:"GWTimes","GWKai";   }
```

校验是否真用上：`page.get_fonts(full=True)` 看 basefont 是否为 `FZXBSK--GBK1-0` / `FangSong_GB2312` / `HYKaiTiKW` / `TimesNewRomanPSMT`。二号标题 + 三号正文的 4 页成品约 140KB（内嵌子集）；若只有几十 KB 或体积明显异常，基本是回退了。
要长期全系统可用，再把 ttf 复制到 `~/Library/Fonts`。

**pymupdf 钤印页码时的西文**：页码数字也该是 Times，破折号仍用宋体，分段钤印再按宽度累加定位：

```python
parts=[("—","china-s"),(" ","tiro"),(str(n),"tiro"),(" ","tiro"),("—","china-s")]
ws=[pymupdf.get_text_length(t,fontname=f,fontsize=14) for t,f in parts]
x = 版心右边界 - 一字宽 - sum(ws)  if n%2 else  版心左边界 + 一字宽   # 单页右空一字/双页左空一字
for (t,f),w in zip(parts,ws):
    page.insert_text((x, y基线), t, fontname=f, fontsize=14); x+=w
```

`tiro` 是 pymupdf 内置 Times-Roman，无需字体文件。


## 2. 公文字体实践速查（通行实现，以用户模板为准）

| 要素 | 惯例 |
|---|---|
| 大标题 | 二号（22pt）方正小标宋简体，居中，**不加粗**；系统目录与 docerFonts 目录均确认无此字体时，黑体不加粗为常见替代 |
| 一级标题（一、） | 三号（16pt）黑体 |
| 二级标题（（一）/阶段标题） | 三号楷体 |
| 正文 | 三号仿宋(_GB2312)，首行缩进 2 字符 |
| 数字与英文 | Times New Roman，与相邻中文同字号（含日期、序号、条目编号、英文单词/机构英文名）；英文内部标点保持原文半角 |
| 行距 | 固定值 28–30pt（29pt 常用） |
| 页边距 | 上 37 / 下 35 / 左 28 / 右 26 mm（A4 + 版心 156×225 的标准实现） |

冲突处理：单位模板明确约定优先于本表；本表是通行惯例，不是国标原文强制项。

## 3. 自动编号陷阱（避免误报缺序号）

- Word/WPS 的「一、二、三、」层次序数常为**段落自动编号**（editor_sdk 读出 `is_numbering: true, numbering_lvl`；OOXML 为 `w:numPr`），纯文本提取和结构树预览**不显示编号前缀**。
- 实际教训：结构树里只见「现状分析」，被误报为缺「二、」序号，用户截图证明显示正常。
- 判定「缺序号」前必须：① 查段落 numbering 属性；② 或向用户要渲染截图。误报会导致用户手工补号，与自动编号叠加成双重编号。

## 4. 标题防孤行（排版建议，非国标条文）

> 口径：本条属通用排版建议（避免标题孤悬页底、标题与正文分离），GB/T 9704—2012 未作硬性规定；单位模板另有约定时从模板。

- 要求：标题不允许落在页面底部而正文翻到下一页。
- 正确机制：**「与下段同页」（keep_with_next）**，不是「段前分页」（page_break_before）——后者会强制每个标题另起一页，过度。
- 做法：对全文所有结构性标题（各层标题、阶段标题）统一设 keep_with_next=true；editor_sdk 可用 `doc_modify_paragraph` 传多个 ranges 一次批量完成。
- 验证：`doc_get_paragraph_property` 读回 keep_with_next=true。

## 5. 属性写入与读回的坑（editor_sdk）

- `doc_update_text_property` 传 `bold: false` 或 `clear_properties: ["bold"]` **可能不生效**；`doc_get_text_property` 读回 `bold: true` 也可能是**误报**（`w:bCs` 复杂文种加粗残留——只影响阿拉伯文等复杂脚本，不影响中文显示）。
- 实际教训：三种方式（bold:false / clear_properties / find_and_set）读回都显示 bold:true，落盘解 XML 后确认 `<w:b/>` 已移除、显示正常。
- **以落盘 XML 为准**：save_file 后解压 docx，检查目标 run 的 `<w:rPr>`（有无 `<w:b/>`、rFonts 值、sz 值）。
- `doc_replace_text` 整段替换会**保留原 run 属性**，无法靠重写文本清除格式。
- 段落标记的 rPr（pPr 内）残留旧字体名不影响正文显示，可忽略。

## 6. 必须人工/渲染核验的项（程序检查不可靠）

- ~~页码格式~~：已程序化为「页码三件套」核验修复（见 §11），仅需交付后目视确认渲染效果。
- 红双线与标志图的间距（程序按估算定位，可能与实际标志图位置差 1–3mm，需目视微调）。
- 实际分页位置（哪些标题真的落在页底——keep_with_next 已兜底，但仍需目视）。
- 印章位置与颜色、附件起始页。
- 建议交付时导出 PDF 逐页目视，或让用户在 WPS 里翻页确认。

## 7. 间距模型（红头函件核心，五轮实战的教训总集）

**国标公文的标准间距模型：段前段后 = 0，全靠统一行距；段与段的视觉间距 = 段内行距。**三个隐形杀手：

- **`Normal (Web)` 样式 + 网页粘贴来源**：样式自带 HTML 自动段距（`beforeAutospacing`/`afterAutospacing` 生效），段间比段内宽。**修复：每个正文段必须显式写 `beforeAutospacing="0" afterAutospacing="0"`**，光设 `w:line="560" w:lineRule="exact"` 不够。终检口径：全文所有 28pt-exact 段落 100% 带这两个属性，一个不漏（含主送段、空段、附件页段）。
- **文档网格（docGrid）**：`docGrid type="lines"` + 段落 `snapToGrid=1` 会把行框吸附到网格线（如 15.6pt 格距），28pt 行距被撑大变形。**修复：`doc_modify_section` 清除 `doc_grid_type`/`doc_grid_line_pitch`/`doc_grid_char_space`（传 null）。**
- **裸段落无 spacing**：分页符空段常是 `<w:p><w:r><w:br type="page"/></w:r></w:p>` 无 pPr，按默认单倍行距渲染出半行空隙。**修复：给它补 28pt-exact + autospacing=0。**

## 8. 隐藏字符陷阱（视觉多空行的元凶）

- **软回车 `&#13;`（Shift+Enter 残留）**：藏在 `<w:t>` 文本里（如「……有限公司：&#13;」「署名&#13;」），不算段落但渲染出半行空白——「某行后面莫名多空一行」九成是它。**修复：XML 直接 replace 删除；终检口径：全文 `&#13;` 计数 = 0。**
- 检测：`xml.count('&#13;')`；扫描 `<w:t>` 含 `&#13;` 的 run 即可定位。
- editor_sdk 层面 `doc_find` 不显示它、结构树也不算它——**必须在 XML 层扫**。

## 9. 落款版式（盖章 vs 不盖章，国标 7.3.5）

| 版式 | 成文日期 | 发文机关署名 |
|---|---|---|
| **盖章（单一机关）** | 右空四字 | 日期上方，**以日期为准居中**（署名对称包住日期） |
| 不盖章 | 右空二字 | 右空二字（下一行日期首字比署名首字右移二字） |

**中心对称的实现（血泪教训：不要用右缩进差值法）：**
- ❌ 错误做法：署名/日期都右对齐 + 各配不同右缩进（如 570/1280 twips）——依赖字宽估算，TNR 数字实际宽度与估算不符时就成了首字对齐，返工。
- ✅ 正确做法：**两段设完全相同的左缩进 + `jc="center"` 居中对齐**——中心线由排版引擎计算，与字宽无关，绝对对称。左缩进按落款区右移量定（示例：4360 twips = 218pt）。
- 先向用户确认盖章与否再选版式；盖章版用印时章下压署名和日期。

## 10. 红头函件专项核验顺序（收口清单）

1. 标题：小标宋、二号、居中、不加粗。**断行规则（国标 7.3.1，属格式项不是内容项）**：长标题（22 字上下）须**主动**按词意手动分行，上长下短梯形排布，「的函」不得孤尾两字——即使正文一字不动，标题断行也要调；实现方式：在词意边界拆段落/软换行，各行仍居中小标宋二号。
2. 主送机关：左顶格、三号仿宋、末尾全角冒号、**段尾无 `&#13;`**、与正文之间**不空行**（间距=行距，国标 7.3.3「正文编排于主送机关名称下一行」）。
3. 正文：首行缩进 2 字、28pt 固定行距、两端对齐、**全部段落 autospacing=0**。
4. 附件说明：正文下空一行（28pt 空段）、左空二字。
5. 落款：先确认盖章版式；盖章版用「同左缩进 + 居中」实现中心对称（§9）。**空段数量核对**：落款上方只空一行（28pt）、附件说明前只空一行——实际教训：原稿联系人/附件说明之间空了 3 行，视觉上被当成排版常规漏改；空段数量要逐个数，不是「有空段就对」。
6. 附件页：「附件：」顶格黑体（国标图示无冒号，保留与否问用户）+ 附件标题居中（版心第三行，惯例小标宋）+ 另面编排（**「附件：」前插分页符**，分页空段也要 28pt-exact）+ 正文仿宋_GB2312。
7. 页码三件套：按 §11 程序化核验与修复（格式 — N — / 宋体四号 / 奇偶分设；信函格式首页隐藏）。
8. 红双线：信函格式首页双红武文线，按 §12 程序化绘制；画完后位置需目视核（与标志图间距）。
9. 终检 XML：`&#13;`=0；28pt-exact 段 100% 带 autospacing=0；docGrid 已清；分页符位置正确（落款后、附件：前）。

## 11. 页码三件套（已程序化核验与修复）

> 国标 7.5：页码用 4 号半角宋体阿拉伯数字，数字左右各放一条一字线；单页码居右空一字，双页码居左空一字。信函格式首页不显示页码（10.1）。

- **三件套**：① 格式 `— N —`（一字线用全角破折号或 U+2014×1，两侧各空半角一空格）；② 字体宋体、四号（14pt）、非加粗；③ 奇偶分设——`titlePg`（信函首页隐藏：奇偶页分节属性 + 首页页脚留空）+ 单页页脚右对齐（右空一字≈缩进 280 twips）、双页页脚左对齐（左空一字）。
- **核验方式（XML）**：解包读 `word/footer*.xml` + `word/document.xml` 的 `sectPr`——查 `w:titlePg` 是否存在、`footerReference` 是否含 `first`/`default`/`even` 三类且指向正确 footer；footer 内 run 的 rFonts/sz/jc 逐一核对。
- **修复方式**：直接重写 footer XML——奇偶页脚各建 `<w:p>`，pPr 设 jc + 缩进，run 设宋体 14pt；sectPr 补 `<w:titlePg/>`（信函版）并把 first 页脚指向空 footer。editor_sdk 对页脚的支持不稳定，**页脚一律走 XML 手术**（§13）。
- 实际教训：某地函件原页码为「- 1 -」连字符格式且首页显示页码，重写两条 footer + titlePg 后收口。

## 12. 红色双线绘制（信函格式首页武文线，程序化 DrawingML）

> 国标 10.1 / 图 11：发文字号之下 4mm 处印一条上粗下细的红色双线；距下页边 20mm 处印一条上细下粗的红色双线；线长 170mm 居中。**仅首页有**。

- **实现**：4 条 DrawingML 锚定直线（`<w:drawing><wp:anchor>` + `a:ln`），锚定于首页正文，`behindDoc="1"`、`allowOverlap="1"`，不用页脚锚（避免随页码区偏移）。
- **实测参数（某地函件收口值，可直接套用）**：
  - 上双线（标志图/发文字号下方 4mm）：粗线 y≈71.2mm、宽度 2.25pt（上粗）+ 细线 y≈71.9mm、宽度 0.75pt（下细）→ **上粗下细**。
  - 下双线（距下页边 20mm，A4 页高 297mm → y≈277mm）：细线 y≈277.0mm、0.75pt + 粗线 y≈277.7mm、2.25pt → **上细下粗**。
  - 两条线组均 x 居中、总长 170mm；颜色 `FF0000`，`cap="flat"`，Y 轴用 `wp:positionV relativeFrom="page"` 以页为基准定位（比相对段落稳）。
- **替换旧图**：原文件常用整幅红头 PNG（含单线）——先删旧 PNG 所在 run，再画线；**红头文字部分保留原 PNG 或改为文本**，问用户。
- **位置微调**：上线与标志图的 4mm 间距是估算值，交付后让用户目视核对，偏 1–3mm 属正常，改 anchor 的 `posOffset` 即可。

## 13. XML 直接手术要领（editor_sdk 失效时的兜底）

- editor_sdk 对 `hint="eastAsia"` 的中英混排 run 改字体**不生效**；`doc_update_text_property` 的 bold:false/clear_properties 也常无效。此时解压 docx 直接改正则改 `word/document.xml`，再原子替换回包。
- **中西文分槽**：`w:rFonts` 的 `w:ascii`/`w:hAnsi` = 西文字体（TNR），`w:eastAsia` = 中文字体（仿宋_GB2312）。给含数字/英文的 run 只改 ascii/hAnsi 槽位、保留 eastAsia 槽位——这是 OOXML 原生机制，比按字符拆 run 稳。
- 纯中文 run 的西文槽位也要统一（否则夹全角数字时露出旧字体）；黑体/小标宋段按段保留。
- 写回：读旧包所有 entry → 只替换 document.xml → 写 .tmp → `os.replace` 原子替换。**文件被 WPS 打开时替换报 WinError 5**：轮询重试（10s × 6 次通常够），仍失败让用户关文件。
- 修改后用户必须**关闭 WPS 重新打开**才能看到新版；日期等文本被 run 分割时 `doc_find` 找不到完整串，按段提取纯文本定位。