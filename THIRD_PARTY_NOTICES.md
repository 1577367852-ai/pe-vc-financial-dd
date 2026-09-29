# 来源、实现与许可边界

本原型的Python代码、测试和说明为本次独立编写；没有复制候选仓库脚本或已有private-equity-investment-dd原文。没有安装或打包其他Skill。本仓库暂未附代码复用许可证；公开可见不等于授权复用。若后续开放复用，需另行确定许可证。

使用的是现有Python环境中的第三方库及系统工具，未将其复制进原型或打包；运行依赖各自许可。

仅供复核的候选静态证据（本轮不再联网搜索）：

- alirezarezvani/claude-skills，提交19392f7a08264ed00486a251f5b2098321771f94，MIT；finance/skills/financial-analyst/scripts/ratio_calculator.py:20、165；forecast_builder.py:184、317；budget_variance_analyzer.py:82；dcf_valuation.py:150。上述仅为先前静态发现，本原型未运行原脚本。
- FinRobot，提交2d059b2e53743a0644152100b836087007a2c104，根LICENSE为Apache-2.0；本模块未复用其清单。
- CaseMark/skills，提交afe48c44914d753297766404378e871d1f88d7a8，根LICENSE为Apache-2.0；只参考尽调方法，不复制其文档。
- CSlawyer1985/financial-statement-analyst，提交6dd6ba1505f190839cae79756a2007d13a76e418，README:964声明MIT，但先前树审查未见独立许可；未复制。
- CSlawyer1985/legal-skillhub，提交0228abc54e52e52e41fc56bc5f2fd761938b3d8c，目标private-equity目录许可未明确，return-modeling.md:185—196示例存在口径不一致；未复制。
- noahnan-max/private-equity-investment-dd-skill，提交18b31eec147cbd84bab1de066c3454420a39af4f，LICENSE为私人/内部使用限定；保留原安装，不修改、不复制进原型。

候选证据是指定版本的静态审查，不代表之后版本仍存在同样问题，也不代表候选仓库整体不安全。
