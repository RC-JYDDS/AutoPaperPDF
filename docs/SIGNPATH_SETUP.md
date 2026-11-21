# SignPath 免费签名准备与手动操作

这份文件用于维护者手动操作，不会替维护者提交申请或发布软件。

## 1. 上传源码并验证构建

1. 解压 V1.8.5 源码上传包。能看到 `README.md`、`.github` 和 `src` 的这一层就是仓库根目录。包内的 `LICENSE` 与打包时公开仓库中的 MIT 许可证相同。
2. 使用 GitHub Desktop 时，先同步 `RC-JYDDS/AutoPaperPDF` 的 `main`，把解压内容复制到本地仓库根目录并合并同名文件夹。使用网页时，在仓库根目录选择 `Add file → Upload files`，上传解压后的内容。
3. 提交前检查 `.gitattributes`、`.gitignore`、`.github/workflows/build.yml` 和 `.github/workflows/sign.yml` 都在对应位置。`vendor/lib` 中的文件按原始字节校验。
4. 提交说明可填 `Update AutoPaperPDF to V1.8.5`。使用 GitHub Desktop 时，提交后推送到 GitHub。
5. 打开仓库的 `Actions`，查看 `Build and test`。必须取得实际成功结果；本地检查不能证明 GitHub 工作流已运行成功。
6. 成功后，在该次运行的 `Artifacts` 下载 `AutoPaperPDF-1.8.5-unsigned`，保存提交和运行的链接。在 Windows 上测试实际 Excel/WPS 导出、页边距和图片，再决定是否发布。

## 2. 准备公开发布及申请

根据基金会条款，项目需要公开的开源代码、可验证的构建过程、文档和实际发布。新建仓库并不自动满足全部条件，审核也不保证通过。

先完成实际构建与人工测试，再在 GitHub Releases 手动建立首个公开版本，附对应 EXE、源代码提交、使用说明和明确的“未签名”状态。不要把当前构建标成 SignPath 已签名。

申请前，确认 GitHub 双因素认证已开启，项目依赖许可证清晰，隐私说明和代码签名政策公开可见。代码签名政策中的团队角色必须符合实际情况。

随后前往 `https://signpath.org/apply`，按页面实际字段提供：

- 项目名：AutoPaperPDF
- 仓库：`https://github.com/RC-JYDDS/AutoPaperPDF`
- 许可证：MIT
- 维护者：AnAn / RC-JYDDS
- 项目说明、真实构建记录和已发布版本链接

不要虚构用户数量、项目历史、签名状态或独立审核人员。以基金会审核结果为准。

## 3. 获批后再配置签名

按基金会提供的指导配置 SignPath 项目、GitHub App 和 GitHub.com 受信任构建系统。配置必须关联此仓库，并要求 GitHub 托管 runner 和每次签名人工批准。

`signpath/artifact-configuration.xml` 提供 ZIP 包内 EXE 的配置模板；产品名限定为“通用表格 PDF 自动选纸工具”，文件名为 `AutoPaperPDF.exe`，版本由构建产生。若服务端校验或审核要求不同，以官方要求调整模板后再继续。

在 GitHub 的 `code-signing` Environment 中配置下列变量和 Secret：

| 名称 | 类型 | 值来源 |
| --- | --- | --- |
| SIGNPATH_ENABLED | Variable | 获批且配置完成后才填写 `true` |
| SIGNPATH_ORGANIZATION_ID | Variable | SignPath 中的 Organization ID |
| SIGNPATH_PROJECT_SLUG | Variable | 已创建项目的 slug |
| SIGNPATH_SIGNING_POLICY_SLUG | Variable | 要求人工批准的生产签名 policy slug |
| SIGNPATH_ARTIFACT_CONFIGURATION_SLUG | Variable | 已验证的 artifact configuration slug |
| SIGNPATH_API_TOKEN | Secret | 只有提交签名所需权限的 API token |

值必须来自实际 SignPath 配置，不是照抄示例名称。不要把 API token、PFX 或证书密码写入文件和提交记录。

配置完成后，在 `Actions → Request SignPath signing → Run workflow` 选择 `main` 手动启动。工作流会从本次提交重新构建并提交待签 EXE；在 SignPath 审查并批准此次签名。成功后下载 `AutoPaperPDF-1.8.5-signed` 构建产物，并按发布检查结果手动更新 Release。

工作流不自动创建 Release。生产签名验证要求 Windows 信任链有效且证书主体包含 `SignPath Foundation`，自签证书不能通过该检查。

## 参考与限制

- 官方申请与条件：<https://signpath.org/apply>、<https://signpath.org/terms>
- GitHub 集成：<https://docs.signpath.io/trusted-build-systems/github>
- 签名配置示例：<https://docs.signpath.io/artifact-configuration/examples>

这是一条待审核的免费开源签名路径，不是购买证书的替代凭据。签名证明文件来源及完整性，不保证 SmartScreen 对所有新版本都不提示。
