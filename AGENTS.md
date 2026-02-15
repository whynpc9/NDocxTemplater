# Rules

- 每次新增一个 feature，都必须同时完成以下事项：
	1. 新增对应测试用例并验证通过；
	2. 新增相应的 example；
	3. 更新相关文档（至少包含 `README.md`）。
- 每次新增特性或修复 bug，都必须执行版本与发布规则：
	1. 根据修改内容按语义化版本（SemVer）更新版本号（建议维护在 `src/NDocxTemplater/NDocxTemplater.csproj` 的 `<Version>`）：
	   - 不兼容变更：升级主版本号（MAJOR）；
	   - 向后兼容的新特性：升级次版本号（MINOR）；
	   - 向后兼容的问题修复：升级修订号（PATCH）。
	2. 版本变更后创建并推送对应 tag（格式：`vX.Y.Z`）。
	3. 通过 tag 触发 NuGet 发布工作流，并确认发布成功。
