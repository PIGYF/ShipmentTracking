# Shipment Tracking

公司物流 shipment 查询原型。当前阶段只实现 DSV public tracking 的拉取验证：

- 从 Excel 读取 `运单` 和 `货代`
- 对 `货代 = DSV` 的行，用 House Bill number 查询 myDSV public tracking
- 先调用 public search 拿 `randomIdentifier`
- 再调用 shipment detail 接口拿 milestones
- 输出 JSON，后续再接 Excel 回写/刷新逻辑

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 运行

```powershell
python -m shipment_tracking.cli --excel "C:\Users\Administrator\Desktop\test.xlsx" --output exports\dsv_results.json
```

也可以直接查单号：

```powershell
python -m shipment_tracking.cli --dsv RTM0344100 --output exports\one.json
```

## 数据安全

不要把 `.env`、API key、账号、密码提交到 GitHub。真实凭据只放本机 `.env` 或 GitHub Actions Secrets。

当前 DSV public tracking 流程不需要 Developer Portal key；如果后续改用正式 DSV API，再按 `.env.example` 补充配置。
