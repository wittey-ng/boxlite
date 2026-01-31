#!/usr/bin/env python3
"""
DocGeneratorBox Example - Document Generation from E-commerce Data

Demonstrates document generation capabilities:
- Generate PowerPoint presentations
- Create Excel data reports
- Produce data visualization charts
- Generate Word analysis reports
- Integration with vector databases (optional)
- AI-powered content generation (optional, requires API key)

Note: This example works with simulated data. For vector database features,
      configure the vector_db_* options and ensure your database is running.
"""

import asyncio
import os
from pathlib import Path
import boxlite


async def example_basic_ppt():
    """Example 1: Generate a basic PowerPoint presentation."""
    print("\n=== Example 1: Generate PowerPoint Presentation ===")

    async with boxlite.DocGeneratorBox() as docgen:
        print("✓ DocGeneratorBox ready")

        print("\nGenerating PowerPoint presentation...")

        slides = [
            {
                "title": "市场概况",
                "content": {
                    "text": "本季度电商数据分析显示：\n\n"
                           "• 总销售额增长 25%\n"
                           "• 用户活跃度提升 18%\n"
                           "• 平均订单价值 ¥350"
                }
            },
            {
                "title": "热门商品类目",
                "content": {
                    "text": "1. 电子产品 - 35%\n"
                           "2. 服装配饰 - 28%\n"
                           "3. 家居用品 - 20%\n"
                           "4. 图书音像 - 17%"
                }
            }
        ]

        result = await docgen.generate_ppt(
            title="电商数据分析报告",
            slides_data=slides,
            output_path="/workspace/ecommerce_report.pptx",
            template="business"
        )
        print(f"✓ {result}")


async def example_excel_report():
    """Example 2: Generate Excel data report."""
    print("\n\n=== Example 2: Generate Excel Report ===")

    async with boxlite.DocGeneratorBox() as docgen:
        print("✓ DocGeneratorBox ready")

        print("\nGenerating Excel report...")

        sheets = {
            "商品数据": [
                {"商品名称": "智能手机", "价格": 2999, "销量": 1200, "评分": 4.5},
                {"商品名称": "笔记本电脑", "价格": 5999, "销量": 800, "评分": 4.7},
                {"商品名称": "无线耳机", "价格": 399, "销量": 2500, "评分": 4.3},
                {"商品名称": "智能手表", "价格": 1599, "销量": 950, "评分": 4.6},
            ],
            "分类统计": [
                {"类目": "电子产品", "销售额": 12500000, "订单数": 8500},
                {"类目": "服装配饰", "销售额": 8900000, "订单数": 15200},
                {"类目": "家居用品", "销售额": 6300000, "订单数": 11000},
            ]
        }

        result = await docgen.generate_excel_report(
            sheets=sheets,
            output_path="/workspace/ecommerce_data.xlsx"
        )
        print(f"✓ {result}")


async def example_chart_generation():
    """Example 3: Generate data visualization charts."""
    print("\n\n=== Example 3: Generate Data Charts ===")

    async with boxlite.DocGeneratorBox() as docgen:
        print("✓ DocGeneratorBox ready")

        # Bar chart: Sales by category
        print("\n1. Generating bar chart (销售额对比)...")
        bar_data = {
            "labels": ["电子产品", "服装", "家居", "图书"],
            "values": [12500, 8900, 6300, 4200]
        }
        result = await docgen.generate_chart(
            data=bar_data,
            chart_type="bar",
            output_path="/workspace/sales_by_category.png",
            title="各类目销售额对比 (万元)"
        )
        print(f"✓ {result}")

        # Pie chart: Market share
        print("\n2. Generating pie chart (市场份额)...")
        pie_data = {
            "labels": ["电子产品", "服装", "家居", "图书"],
            "values": [35, 28, 20, 17]
        }
        result = await docgen.generate_chart(
            data=pie_data,
            chart_type="pie",
            output_path="/workspace/market_share.png",
            title="市场份额分布 (%)"
        )
        print(f"✓ {result}")


async def example_word_report():
    """Example 4: Generate Word analysis report."""
    print("\n\n=== Example 4: Generate Word Report ===")

    async with boxlite.DocGeneratorBox() as docgen:
        print("✓ DocGeneratorBox ready")

        print("\nGenerating Word document...")

        sections = [
            {
                "heading": "执行摘要",
                "paragraphs": [
                    "本报告基于2024年第一季度电商平台数据，对市场趋势、热门商品、用户行为进行深度分析。",
                    "主要发现：销售额同比增长25%，移动端订单占比达到68%，用户复购率提升12个百分点。"
                ]
            },
            {
                "heading": "市场趋势分析",
                "paragraphs": [
                    "电子产品类目持续领跑，占总销售额的35%。智能穿戴设备成为增长最快的细分品类。"
                ]
            },
            {
                "heading": "商品排行",
                "table": [
                    ["排名", "商品名称", "销量", "销售额（万元）"],
                    ["1", "智能手机A", "1200", "359.9"],
                    ["2", "无线耳机B", "2500", "99.8"],
                    ["3", "笔记本电脑C", "800", "479.9"],
                ]
            }
        ]

        result = await docgen.generate_word_report(
            title="2024 Q1 电商数据分析报告",
            sections=sections,
            output_path="/workspace/analysis_report.docx"
        )
        print(f"✓ {result}")


async def example_with_volumes():
    """Example 5: Using volumes to access generated files."""
    print("\n\n=== Example 5: Generate and Access Files ===")

    # Create local output directory
    output_dir = Path("./docgen_output")
    output_dir.mkdir(exist_ok=True)

    opts = boxlite.DocGeneratorBoxOptions(
        cpu=2,
        memory=4096
    )

    # Mount local directory to access generated files
    async with boxlite.DocGeneratorBox(
        options=opts,
        volumes=[
            (str(output_dir.absolute()), "/output", "rw")
        ]
    ) as docgen:
        print(f"✓ DocGeneratorBox ready with volume: {output_dir}")

        print("\nGenerating chart and saving to mounted volume...")
        data = {
            "labels": ["Q1", "Q2", "Q3", "Q4"],
            "values": [850, 920, 1100, 1350]
        }
        result = await docgen.generate_chart(
            data=data,
            chart_type="line",
            output_path="/output/quarterly_sales.png",
            title="季度销售趋势"
        )
        print(f"✓ {result}")

        # Verify file exists locally
        local_file = output_dir / "quarterly_sales.png"
        if local_file.exists():
            print(f"\n✓ File accessible at: {local_file}")
            print(f"  File size: {local_file.stat().st_size / 1024:.1f} KB")


async def example_ai_integration():
    """Example 6: AI-powered content generation (requires OpenAI API key)."""
    print("\n\n=== Example 6: AI Content Generation ===")

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("⚠ Skipping: OPENAI_API_KEY environment variable not set")
        print("  To use AI features, set: export OPENAI_API_KEY='sk-...'")
        return

    async with boxlite.DocGeneratorBox(
        env=[("OPENAI_API_KEY", api_key)]
    ) as docgen:
        print("✓ DocGeneratorBox ready with AI integration")

        print("\nGenerating AI-powered summary...")
        data = """
        商品数据：
        - 智能手机销量1200件，均价2999元
        - 笔记本电脑销量800件，均价5999元
        - 无线耳机销量2500件，均价399元

        用户反馈：
        - 智能手机好评率95%
        - 笔记本电脑好评率97%
        - 无线耳机好评率89%
        """

        summary = await docgen.generate_summary(
            text=data
        )
        print(f"\nAI 生成摘要:\n{summary}")


async def main():
    """Run all examples."""
    print("DocGeneratorBox Examples - E-commerce Document Generation")
    print("=" * 70)

    await example_basic_ppt()
    await example_excel_report()
    await example_chart_generation()
    await example_word_report()
    await example_with_volumes()
    await example_ai_integration()

    print("\n" + "=" * 70)
    print("✓ All examples completed!")
    print("\nKey Takeaways:")
    print("  • generate_ppt() - Create PowerPoint presentations")
    print("  • generate_excel_report() - Generate Excel spreadsheets")
    print("  • generate_chart() - Create data visualization charts")
    print("  • generate_word_report() - Produce Word documents")
    print("  • Use volumes to access generated files on host")
    print("  • Optional AI integration for content generation")
    print("\nNext Steps:")
    print("  • Build the Docker image: docker build -t boxlite/doc-generator dockerfiles/docgenerator/")
    print("  • Check generated files in ./docgen_output/")
    print("  • Set OPENAI_API_KEY for AI features")


if __name__ == "__main__":
    asyncio.run(main())
