def capabilities(script_execution=True):
    return dict(calculation_status='AVAILABLE_NOT_YET_EXECUTED' if script_execution else 'NOT_EXECUTED',
                full_acceptance=False,
                note='能力存在不等于校验通过。无执行环境时仅整理材料和展示公式；不得宣称程序验证。')
