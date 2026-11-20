# 遥控割草机客户爬取 - 状态

- 项目：gosom/google-maps-scraper，已下载到 src/，已用 Go 1.27.1 本地编译为 google-maps-scraper（41MB）。
- 浏览器：Playwright Chromium 已手动装好（官方下载器在 90% 处断线，改用 curl 断点续传 + INSTALLATION_COMPLETE 标记解决）。
- 验证跑（2026-11-04）：柏林 2 条搜索，36 条结果，电话 36/36，网站 35/36。文件：out/validation.csv（原始）、out/validation_clean.csv（去重中文表）。
- 欧洲第一批（进行中）：queries_europe_batch1.txt，31 条搜索，覆盖柏林、慕尼黑、巴黎、伦敦、罗马、马德里、阿姆斯特丹、华沙、维也纳；类别：割草机经销商、园林机械经销商、园林公司、设备租赁、农机经销商、建筑公司。参数：-depth 1 -c 2 -resume，输出 out/europe_batch1.csv，日志 out/europe_batch1.log。
- 注意：不带代理时浏览器直连失败，必须用代理跑；不要把代理凭证写进文件或发给用户。
- 欧洲第一批（已暂停 2026-11-04）：327 条已保存，31 条搜索中 20 条已完成，断点续跑文件 out/europe_batch1.csv.resume.json 已生成，之后可用 -resume 续跑，不用重跑已完成部分。
- 美国 1000 条（进行中 2026-11-04）：queries_us_1000.txt，70 条搜索，10 个城市（洛杉矶、休斯敦、达拉斯、迈阿密、奥兰多、亚特兰大、芝加哥、凤凰城、丹佛、西雅图）×6 类 + 建筑、农机、树木服务补充，参数 -depth 1 -c 2 -resume，输出 out/us_1000.csv，日志 out/us_1000.log，过滤脚本 us_clean.py。
- 下一步：美国跑完后去重、过滤无联系方式和不相关类别，按客户类型分类统计，再交表；澳大利亚先不管。
