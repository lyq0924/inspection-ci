import re

with open('generate_security_report_test_cases.py', 'r', encoding='utf-8') as f:
    content = f.read()

strs_all = re.findall(r"'((?:[^'\\]|\\.)*?)'", content)
issues = []
in_cases = False
case_idx = 0

check_keywords = ['找到该', '找到"', '找到已', '找到刚', '找到历史', '找到执行', '在列表中找到']

for i, s in enumerate(strs_all):
    if s.startswith('TC-'):
        case_idx = i
        in_cases = True
    if in_cases and (i - case_idx) == 7:
        steps = s
        results = strs_all[i+1] if i+1 < len(strs_all) else ''
        step_lines = [l for l in steps.split('\n') if re.match(r'^\d+\.', l.strip())]
        result_lines = [l for l in results.split('\n') if re.match(r'^\d+\.', l.strip())]
        tc_id = strs_all[case_idx]
        if len(step_lines) != len(result_lines):
            issues.append(f'{tc_id}: 步骤{len(step_lines)}条 != 结果{len(result_lines)}条')
        for sl in step_lines:
            for kw in check_keywords:
                if kw in sl:
                    issues.append(f'{tc_id}: 步骤包含不规范描述 "{kw}"')
                    break
            if '查看或点击' in sl:
                issues.append(f'{tc_id}: 步骤包含模糊描述 "查看或点击"')
        in_cases = False

if issues:
    print('发现问题：')
    for issue in issues:
        print(f'  - {issue}')
else:
    print('未发现步骤/结果数量不匹配问题')
