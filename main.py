"""
PPT Agent - 自然语言驱动的 PPT 自动化工具
作者：使用 AI 编程工具（Cursor/Codex）辅助开发
"""
import os
import sys

# 把当前目录加入路径，方便导入 agent 模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from intent_parser import IntentParser
from ppt_executor import PPTExecutor
from context_manager import ContextManager
from ppt_tools import create_sample_files


def print_banner():
    print("=" * 50)
    print("🤖  PPT Agent  -  自然语言驱动 PPT 自动化")
    print("=" * 50)
    print("支持指令示例：")
    print('  • "把所有标题改成红色 24 号字"')
    print('  • "把第 3 页标题改为 Q3 财报"')
    print('  • "批量把 旧公司名 替换为 新公司名"')
    print('  • "从 data.xlsx 读取数据填充到第 2 页"')
    print('  • 输入 "退出" 结束程序')
    print("-" * 50)


def main():
    print_banner()

    # 初始化核心组件
    context = ContextManager()
    parser = IntentParser()
    executor = PPTExecutor()

    # 如果没有示例文件，自动创建
    if not os.path.exists("examples/sample.pptx"):
        print("📦 正在生成示例文件...")
        create_sample_files()
        print("✅ 示例文件已生成：examples/sample.pptx 和 examples/data.xlsx\n")

    while True:
        user_input = input("\n📝 请输入指令: ").strip()
        if user_input in ["退出", "quit", "exit", "q"]:
            print("\n👋 感谢使用 PPT Agent，再见！")
            print(f"📊 本次会话共执行 {len(context.history)} 个操作")
            break

        if not user_input:
            continue

        # 1. 解析用户意图
        intent = parser.parse(user_input)
        print(f"🔍 意图解析: {intent}")

        # 2. 确认 PPT 文件路径
        if not context.ppt_path:
            default_path = "examples/sample.pptx"
            use_default = input(f"📂 使用示例文件 [{default_path}]? (回车确认/输入其他路径): ").strip()
            ppt_path = use_default if use_default else default_path
            if not os.path.exists(ppt_path):
                print(f"❌ 文件不存在: {ppt_path}")
                continue
            context.set_ppt(ppt_path)

        # 3. 执行操作
        try:
            result = executor.execute(intent, context)
            print(f"✅ {result}")
            context.log_action(intent)
        except Exception as e:
            print(f"❌ 执行失败: {e}")
            print("💡 提示: 请检查 PPT 文件是否被其他程序占用")


if __name__ == "__main__":
    main()
