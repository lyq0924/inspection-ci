import re

def check_test_cases():
    with open('generate_security_report_test_cases.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # 提取 test_cases 列表内容
    match = re.search(r'test_cases = \[(.*?)\]', content, re.DOTALL)
    if not match:
        print("未找到 test_cases")
        return

    # 提取每条用例（简化处理，按行查找）
    lines = content.split('\n')
    issues = []
    current_case = None
    in_case = False
    case_lines = []

    for i, line in enumerate(lines):
        if "['TC-" in line:
            in_case = True
            case_lines = [line]
        elif in_case:
            case_lines.append(line)
            if line.strip().endswith("'-', '-', '-', '-']") or line.strip().endswith("'-']"):
                # 完整的一条用例
                case_text = '\n'.join(case_lines)
                # 提取用例编号
                tc_match = re.search(r"'TC-(\d+)'", case_text)
                if tc_match:
                    tc_id = f"TC-{tc_match.group(1)}"
                    # 提取测试步骤和预期结果
                    # 用例格式：['TC-xxx', '模块', '二级', '名称', '测试点', '级别', '前置', '步骤', '结果', ...]
                    # 找到步骤和结果（第8和第9个字符串元素，索引7和8）
                    parts = []
                    # 简单按 ', ' 分割不行，因为内容里有逗号
                    # 改用正则提取被单引号包裹的字符串
                    strs = re.findall(r"'((?:[^'\\]|\\.)*?)'", case_text)
                    if len(strs) >= 9:
                        steps = strs[7]
                        results = strs[8]
                        step_count = len([s for s in steps.split('\n') if s.strip().startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'))])
                        result_count = len([s for s in results.split('\n') if s.strip().startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'))])
                        if step_count != result_count:
                            issues.append(f"{tc_id}: 步骤数({step_count}) != 结果数({result_count})")
                        # 检查"找到"开头的步骤
                        if '找到' in steps and '找到该' in steps:
                            issues.append(f"{tc_id}: 步骤以'找到'开头，不够规范")
                        # 检查步骤和结果语义对应
                        step_lines = [s.strip() for s in steps.split('\n') if s.strip().startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'))]
                        result_lines = [s.strip() for s in results.split('\n') if s.strip().startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.'))]
                        for idx, (sl, rl) in enumerate(zip(step_lines, result_lines), 1):
                            # 检查是否有"查看或点击"模糊描述
                            if '查看或点击' in sl:
                                issues.append(f"{tc_id} 步骤{idx}: 包含'查看或点击'模糊描述")
                in_case = False
                case_lines = []

    if issues:
        print("发现以下问题：")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("未发现明显问题")

if __name__ == '__main__':
    check_test_cases()
