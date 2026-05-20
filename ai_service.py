from openai import OpenAI
import json
import os

# ----- 自动加载 .env 文件 -----
def _load_dotenv():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.isfile(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value

_load_dotenv()

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    raise ValueError("请先设置环境变量 DEEPSEEK_API_KEY")

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com"
)

def analyze_order_risk(order_info: dict) -> dict:
    """分析单个订单的风险和给出建议。失败时返回空值，不中断排单流程。"""
    prompt = f"""你是供应链排单专家。请分析以下订单，只输出最关键的风险和一句建议。

订单信息：
{json.dumps(order_info, ensure_ascii=False, indent=2)}

规则：
1. 只关注真正重要的问题：严重缺货、交期过长、紧急订单处理不当、RB800远程控制项目交期、特殊订单优先级
2. 正常订单直接说"无明显风险"
3. 建议必须一句话，直击要点
4. risk 取值：无明显风险 / 库存风险 / 交期风险 / 高优先级风险

必须只返回 JSON（不要 markdown 代码块）：
{{"risk":"库存风险","advice":"建议优先备货SKU xxx，当前缺口数量较大"}}"""

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            timeout=30,
        )
        result_text = response.choices[0].message.content.strip()
        if result_text.startswith("```"):
            result_text = result_text.split("\n", 1)[1]
            if result_text.endswith("```"):
                result_text = result_text[:-3]
        return json.loads(result_text)
    except Exception as e:
        print(f"[AI分析] 调用失败: {e}")
        return {"risk": "分析失败", "advice": ""}
