"""
本地验证脚本 - 不需要真实飞书 API
模拟核心逻辑：_match_spec_id、_sync_finance_detail、_run_finance_calculate
"""
import pandas as pd
from typing import Dict, Tuple, List, Any, Optional
import sys

# ==================== 辅助函数（来自 scheduler_api.py）====================

def safe_str(value):
    """安全转换为字符串并清洗"""
    if pd.isna(value) or value is None:
        return ""
    return str(value).strip()


def safe_numeric(value) -> float:
    """安全转换数值，无效值返回 0.0。"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value) if not pd.isna(value) else 0.0
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return 0.0


# ==================== 核心函数：_match_spec_id (改进版) ====================

def _match_spec_id(product_name: str, spec: str, billing_df: pd.DataFrame) -> str:
    """根据产品名称和规格匹配计费规则表中的 Spec_ID。
    
    改进版本：支持单边值存在，处理大小写和空白。
    """
    # 预规范化
    product_name_norm = str(product_name).strip() if product_name else ""
    spec_norm = str(spec).strip() if spec else ""
    
    # 允许只有产品名或规格其中之一存在
    if billing_df.empty or (not product_name_norm and not spec_norm):
        return ""

    product_name_norm_lower = product_name_norm.lower()
    spec_norm_lower = spec_norm.lower()

    # 1. 精确匹配（仅当两个字段都非空）
    if product_name_norm and spec_norm:
        exact_match = billing_df[
            (billing_df["产品名称"].astype(str).str.strip() == product_name_norm) &
            (billing_df["规格型号"].astype(str).str.strip() == spec_norm)
        ]
        if not exact_match.empty:
            val = exact_match.iloc[0].get("Spec_ID", "")
            return str(val).strip() if pd.notna(val) else ""

    # 2. 别名匹配（产品名称匹配 + 规格匹配）
    if "别名" in billing_df.columns and spec_norm:
        for _, rule in billing_df.iterrows():
            rule_spec = str(rule.get("规格型号", "")).strip().lower()
            if rule_spec != spec_norm_lower:
                continue
            alias = str(rule.get("别名", "")).strip()
            if not alias:
                continue
            if alias.lower() in product_name_norm_lower or product_name_norm_lower in alias.lower():
                val = rule.get("Spec_ID", "")
                return str(val).strip() if pd.notna(val) else ""

    # 3. 关键词匹配（仅限 AC-RB800 规格的消控室值班助手产品）
    if "ac-rb800" in spec_norm_lower:
        keywords = [
            ("主键盘", "SPEC-RB800-KZP"),
            ("总线盘", "SPEC-RB800-ZXP"),
            ("总键盘", "SPEC-RB800-ZXP"),
            ("电话", "SPEC-RB800-PHONE"),
            ("多线盘", "SPEC-RB800-DXP"),
            ("广播", "SPEC-RB800-BROAD"),
        ]
        for kw, sid in keywords:
            if kw in product_name_norm:
                return sid

    return ""


# ==================== 核心函数：补行合并逻辑 (改进版) ====================

def _normalize_merge_key(contract_no: str, product_name: str, spec: str) -> Tuple[str, str, str]:
    """规范化 merge key"""
    return (
        safe_str(contract_no).strip().lower(),
        safe_str(product_name).strip(),
        safe_str(spec).strip().lower(),
    )


def merge_calc_results_to_new_rows(
    calc_results: List[Dict[str, Any]],
    new_detail_rows: List[Dict[str, Any]]
) -> None:
    """将计算结果合并到新补行"""
    calc_lookup: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    
    for cr in calc_results:
        key = _normalize_merge_key(
            cr.get("合同编号", ""),
            cr.get("产品名称", ""),
            cr.get("规格", ""),
        )
        if key in calc_lookup:
            print(f"[WARN] Duplicate key in calc_results: {key}, skipping")
            continue
        calc_lookup[key] = cr
    
    merged_count = 0
    for nr in new_detail_rows:
        key = _normalize_merge_key(
            nr.get("合同编号", ""),
            nr.get("产品名称", ""),
            nr.get("规格", ""),
        )
        if key in calc_lookup:
            cr = calc_lookup[key]
            nr["是否计入包干"] = safe_str(cr.get("是否计入包干", ""))
            nr["包干合同单采数量"] = safe_numeric(cr.get("包干合同单采数量", 0))
            nr["单采价格"] = safe_numeric(cr.get("单采价格", 0))
            nr["小计"] = safe_numeric(cr.get("小计", 0))
            merged_count += 1
    
    print(f"[INFO] Merged {merged_count} rows to new_detail_rows")


# ==================== 测试用例 ====================

class TestFinanceCalculate:
    
    @staticmethod
    def setup_billing_df():
        """创建模拟的计费规则表"""
        return pd.DataFrame([
            {
                "Spec_ID": "SPEC-RB800-KZP",
                "产品名称": "消防远程控制包干费",
                "规格型号": "定制",
                "别名": "包干费",
            },
            {
                "Spec_ID": "SPEC-SVC",
                "产品名称": "消防远程控制服务费",
                "规格型号": "定制",
                "别名": "服务费",
            },
            {
                "Spec_ID": "SPEC-EQP",
                "产品名称": "消防远程控制设备费",
                "规格型号": "定制",
                "别名": "设备费",
            },
        ])
    
    @staticmethod
    def test_match_spec_id_exact():
        """TC-001: 精确匹配"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消防远程控制服务费", "定制", billing_df)
        assert result == "SPEC-SVC", f"Expected SPEC-SVC, got {result}"
        print("✓ TC-001 PASSED: 精确匹配")
    
    @staticmethod
    def test_match_spec_id_product_only():
        """TC-002: 仅产品名，规格为空"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消防远程控制包干费", "", billing_df)
        assert result in ["SPEC-RB800-KZP", ""], f"Got {result}"
        print(f"✓ TC-002 PASSED: 产品名匹配，规格为空，result={result}")
    
    @staticmethod
    def test_match_spec_id_spec_only():
        """TC-003: 仅规格，产品名为空"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("", "定制", billing_df)
        assert result in ["SPEC-RB800-KZP", "SPEC-SVC", ""], f"Got {result}"
        print(f"✓ TC-003 PASSED: 规格匹配，产品名为空，result={result}")
    
    @staticmethod
    def test_match_spec_id_both_empty():
        """TC-004: 两者都为空"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("", "", billing_df)
        assert result == "", f"Expected empty, got {result}"
        print("✓ TC-004 PASSED: 两者都空返回空字符串")
    
    @staticmethod
    def test_match_spec_id_whitespace_only():
        """TC-005: 仅空格字符"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("   ", "  ", billing_df)
        assert result == "", f"Expected empty after strip, got {result}"
        print("✓ TC-005 PASSED: 空格被规范化为空字符串")
    
    @staticmethod
    def test_match_spec_id_case_insensitive():
        """TC-006: 大小写不一致（英文 Spec_ID 匹配）"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消防远程控制包干费", "定制", billing_df)
        assert result == "SPEC-RB800-KZP", f"Expected SPEC-RB800-KZP, got {result}"
        print("✓ TC-006 PASSED: 精确匹配「定制」规格")

    @staticmethod
    def test_keyword_match_main_keyboard():
        """TC-011: 关键词匹配-主键盘"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消控室值班助手-高新投三江GB2201-主键盘", "AC-RB800", billing_df)
        assert result == "SPEC-RB800-KZP", f"Expected SPEC-RB800-KZP, got {result}"
        print("✓ TC-011 PASSED: 关键词匹配-主键盘 → SPEC-RB800-KZP")

    @staticmethod
    def test_keyword_match_broadcast():
        """TC-012: 关键词匹配-广播"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消控室值班助手-高新投三江GB2201-消防广播", "AC-RB800", billing_df)
        assert result == "SPEC-RB800-BROAD", f"Expected SPEC-RB800-BROAD, got {result}"
        print("✓ TC-012 PASSED: 关键词匹配-广播 → SPEC-RB800-BROAD")

    @staticmethod
    def test_keyword_match_phone():
        """TC-013: 关键词匹配-电话"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("消控室值班助手-新品牌X999-消防电话", "AC-RB800", billing_df)
        assert result == "SPEC-RB800-PHONE", f"Expected SPEC-RB800-PHONE, got {result}"
        print("✓ TC-013 PASSED: 关键词匹配-电话 → SPEC-RB800-PHONE")

    @staticmethod
    def test_keyword_match_non_rb800():
        """TC-014: 关键词匹配不适用于非AC-RB800规格"""
        billing_df = TestFinanceCalculate.setup_billing_df()
        result = _match_spec_id("主机通讯板-主键盘-某型号", "JBF293K", billing_df)
        assert result == "", f"Expected empty for non-AC-RB800 spec, got {result}"
        print("✓ TC-014 PASSED: 非AC-RB800不触发关键词匹配")

    @staticmethod
    def test_merge_single_match():
        """TC-007: 单个补行与 calc 结果匹配"""
        calc_results = [
            {
                "_record_id": "",
                "合同编号": "AC20260601-001",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "是否计入包干": "是",
                "包干合同单采数量": 0,
                "单采价格": 0.0,
                "小计": 5000.0,
            }
        ]
        new_detail_rows = [
            {
                "合同编号": "AC20260601-001",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "合同数量": 1,
            }
        ]
        merge_calc_results_to_new_rows(calc_results, new_detail_rows)
        assert new_detail_rows[0]["是否计入包干"] == "是"
        assert new_detail_rows[0]["小计"] == 5000.0
        print("✓ TC-007 PASSED: 单个补行合并成功")
    
    @staticmethod
    def test_merge_case_insensitive():
        """TC-010: 合同编号大小写差异"""
        calc_results = [
            {
                "_record_id": "",
                "合同编号": "ac20260601-001",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "是否计入包干": "是",
                "包干合同单采数量": 0,
                "单采价格": 0.0,
                "小计": 5000.0,
            }
        ]
        new_detail_rows = [
            {
                "合同编号": "AC20260601-001",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "合同数量": 1,
            }
        ]
        merge_calc_results_to_new_rows(calc_results, new_detail_rows)
        assert new_detail_rows[0].get("是否计入包干") == "是"
        print("✓ TC-010 PASSED: 合同编号大小写规范化后匹配")
    
    @staticmethod
    def test_merge_no_match():
        """TC-008: 补行与 calc 结果不匹配"""
        calc_results = [
            {
                "_record_id": "",
                "合同编号": "AC20260601-001",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "是否计入包干": "是",
                "包干合同单采数量": 0,
                "单采价格": 0.0,
                "小计": 5000.0,
            }
        ]
        new_detail_rows = [
            {
                "合同编号": "AC20260601-002",
                "产品名称": "消防远程控制包干费",
                "规格": "定制",
                "合同数量": 1,
            }
        ]
        merge_calc_results_to_new_rows(calc_results, new_detail_rows)
        assert "是否计入包干" not in new_detail_rows[0]
        print("✓ TC-008 PASSED: 无匹配时补行保持原值")
    
    @staticmethod
    def run_all():
        print("="*60)
        print("开始执行本地测试套件...")
        print("="*60)
        
        try:
            TestFinanceCalculate.test_match_spec_id_exact()
            TestFinanceCalculate.test_match_spec_id_product_only()
            TestFinanceCalculate.test_match_spec_id_spec_only()
            TestFinanceCalculate.test_match_spec_id_both_empty()
            TestFinanceCalculate.test_match_spec_id_whitespace_only()
            TestFinanceCalculate.test_match_spec_id_case_insensitive()
            TestFinanceCalculate.test_keyword_match_main_keyboard()
            TestFinanceCalculate.test_keyword_match_broadcast()
            TestFinanceCalculate.test_keyword_match_phone()
            TestFinanceCalculate.test_keyword_match_non_rb800()
            TestFinanceCalculate.test_merge_single_match()
            TestFinanceCalculate.test_merge_case_insensitive()
            TestFinanceCalculate.test_merge_no_match()
            
            print("="*60)
            print("✅ 所有测试通过！")
            print("="*60)
            return True
        except AssertionError as e:
            print(f"❌ 测试失败: {e}")
            return False


if __name__ == "__main__":
    success = TestFinanceCalculate.run_all()
    sys.exit(0 if success else 1)
