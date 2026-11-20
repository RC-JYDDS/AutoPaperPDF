# Code signing policy / 代码签名政策

## 当前状态

AutoPaperPDF 尚处于 SignPath Foundation 申请准备阶段。此文件描述计划采用的流程，不表示已获批准、已安装受信任证书或已有正式签名版本。普通构建产物保持未签名。

## 计划采用的发布规则

1. 源码、依赖版本、构建脚本和版本元数据在本公开仓库中维护。
2. 发布程序从明确的 Git 提交在 GitHub 托管的 Windows runner 上构建，运行测试，并先上传为 GitHub Actions 构建产物。
3. 获得 SignPath Foundation 批准并配置后，由维护者手动发起签名请求。每次生产签名均须在 SignPath 中人工审核批准；工作流本身不能替代该审核。
4. 只对属于本项目、具有匹配产品名和版本信息的 EXE 签名；不签署无关第三方软件。嵌入的开源依赖保留其原有许可证声明。
5. 发布前检查返回文件的 Authenticode 信任状态及证书主体，再由维护者手动发布并标明版本、源码提交和签名状态。

## 维护者与访问控制

目前公开列出的维护者是 AnAn（GitHub：RC-JYDDS）。尚未指定独立的代码审查者或签名批准人，申请时会如实说明并按基金会要求配置角色，不虚构多人审查。

启用生产签名前，参与者须按 SignPath 的要求开启账户双因素认证；API 令牌仅保存在 GitHub Actions Secret 中，证书私钥不进入仓库。

## 获批后的更新

获批后，本页和发布页应更新实际服务、证书主体、团队角色及 SignPath 要求的致谢。采用基金会证书时，Windows 显示的证书发布者由基金会证书决定，不能通过 EXE 的作者字段改成 AnAn。

参考：[SignPath Foundation terms](https://signpath.org/terms)、[GitHub integration](https://docs.signpath.io/trusted-build-systems/github)。
