V1.8.3 兼容与诊断检查记录 — 2026-11-21

现场证据：
用户的 V1.8.2 在开始导出后返回 HRESULT 0x8FE30C28。截图显示纸型来源为 Kyocera TASKalfa 2021 GX，接口预估打印区域 189.4 × 102.7 mm，四边留白各 5 mm。
旧日志未区分首次导出、复制、测量导出，不能据此确定具体失败调用，也不能认定京瓷打印机是根因。

本版调整：
- ExportAsFixedFormat 采用五个显式参数，保留 IgnorePrintAreas=false，省略 From/To 等不用的可选参数，避免传入 Missing.Value。
- 临时工作表采用单个 Before 参数复制，避免 Copy 的缺省占位参数。
- 中间排版优先使用可容纳内容的 A4/A3 等常见纸型；最终页面仍由实际打印边界及指定留白决定。不更改系统默认打印机。
- 记录实际自动化接口、软件版本、尝试纸型、每个导出/复制/标记/清理阶段的开始与完成。
- 异常包含失败步骤及原始 HRESULT；清理失败单独记录，不覆盖之前的导出错误。
- 导出接口返回后检查文件存在、大小和 PDF 文件头。

验证：
- C# 5 + .NET Framework 4.7.2 引用程序集本地编译通过。
- 35 项核心、27 项页边距、190 项 PDF 边界/布局及 20 项接口调用与异常路径检查通过，共 272 项。
- 新增测试使用内存对象替代 COM，实际执行测量流程；覆盖五参数导出、单参数复制、无文件/非 PDF 文件拒绝、纸型排序、恢复原工作表、原边框不变、导出失败后的清理，以及清理错误不覆盖 0x8FE30C28。
- 35 个依赖文件的 SHA-256 与锁定记录一致。

尚未验证：
- 用户 Windows Excel/WPS 上是否解决 0x8FE30C28。
- 真实 COM 调度、打印驱动和图形界面完整导出。
- 当前 V1.8.3 GitHub Windows 构建及 SignPath 签名。

结论：本版是未签名兼容诊断测试版，不能作为已验证修复的正式发行版。

官方接口参考：
https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Worksheet/member/ExportAsFixedFormat
https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.copy
https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.papersize
