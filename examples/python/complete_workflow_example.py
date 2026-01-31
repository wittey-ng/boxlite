#!/usr/bin/env python3
"""
Complete E-commerce Report Workflow - 完整电商数据报告工作流

展示 DocGeneratorBox 的完整功能：
- 文件上传/下载
- 读取现有文档
- 修改文档内容
- 生成新报告
- 工作区管理
- 定期更新报告

使用场景：每月自动生成和更新电商分析报告
"""

import asyncio
import os
from pathlib import Path
from datetime import datetime
import boxlite


async def scenario_1_initial_report_generation():
    """
    场景1：首次生成报告并保存工作区

    流程：
    1. 创建工作区
    2. 生成初始报告（PPT + Excel + 图表）
    3. 保存工作区到本地（可以暂停工作）
    """
    print("\n" + "=" * 70)
    print("场景 1: 首次生成电商分析报告")
    print("=" * 70)

    async with boxlite.DocGeneratorBox() as docgen:
        # 1. 创建工作区
        print("\n[1/4] 创建工作区...")
        workspace = await docgen.create_workspace("ecommerce_q1_2024")
        print(f"✓ 工作区已创建: {workspace}")

        # 2. 生成销售数据图表
        print("\n[2/4] 生成销售数据图表...")
        sales_data = {
            "labels": ["1月", "2月", "3月"],
            "values": [1250000, 1480000, 1920000]
        }
        await docgen.generate_chart(
            data_dict=sales_data,
            chart_type="line",
            output_path=f"{workspace}/output/monthly_sales.png",
            title="Q1 月度销售额趋势（万元）"
        )
        print("✓ 图表已生成")

        # 3. 生成 Excel 数据报表
        print("\n[3/4] 生成 Excel 数据报表...")
        excel_data = {
            "销售数据": [
                {"月份": "1月", "销售额": 1250000, "订单数": 8500, "客单价": 147},
                {"月份": "2月", "销售额": 1480000, "订单数": 9200, "客单价": 161},
                {"月份": "3月", "销售额": 1920000, "订单数": 11500, "客单价": 167},
            ],
            "商品分类": [
                {"类目": "电子产品", "占比": "35%", "增长": "+12%"},
                {"类目": "服装配饰", "占比": "28%", "增长": "+8%"},
                {"类目": "家居用品", "占比": "20%", "增长": "+15%"},
                {"类目": "图书音像", "占比": "17%", "增长": "+5%"},
            ]
        }
        await docgen.generate_excel_report(
            sheets=excel_data,
            output_path=f"{workspace}/output/q1_data.xlsx"
        )
        print("✓ Excel 报表已生成")

        # 4. 生成 PowerPoint 报告
        print("\n[4/4] 生成 PowerPoint 报告...")
        slides = [
            {
                "title": "Q1 销售概况",
                "content": {
                    "text": "总销售额: ¥465万\\n总订单数: 29,200单\\n平均客单价: ¥159\\n\\n同比增长: 25%"
                }
            },
            {
                "title": "月度趋势分析",
                "content": {
                    "text": "3月销售额达到最高点\\n环比增长 29.7%\\n主要增长来自电子产品类目"
                }
            },
            {
                "title": "类目分布",
                "content": {
                    "text": "电子产品领跑，占比35%\\n家居用品增长最快，达15%\\n服装配饰稳定增长8%"
                }
            }
        ]
        await docgen.generate_ppt(
            title="2024 Q1 电商数据分析报告",
            slides_data=slides,
            output_path=f"{workspace}/output/q1_report.pptx",
            template="business"
        )
        print("✓ PPT 报告已生成")

        # 5. 查看工作区文件
        print("\n[查看] 工作区文件列表:")
        files = await docgen.list_files(workspace, "**/*")
        for f in files:
            print(f"  - {f['name']} ({f['size']} bytes)")

        # 6. 保存工作区到本地
        print("\n[保存] 保存工作区到本地...")
        local_workspace = "./workspaces/ecommerce_q1_2024"
        await docgen.save_workspace(workspace, local_workspace)
        print(f"✓ 工作区已保存到 {local_workspace}")
        print("\n提示: 工作区已保存，可以随时恢复并继续工作")


async def scenario_2_update_monthly_report():
    """
    场景2：定期更新报告（下个月）

    流程：
    1. 恢复工作区
    2. 读取现有报告
    3. 修改报告内容（添加新月份数据）
    4. 重新保存
    """
    print("\n" + "=" * 70)
    print("场景 2: 更新报告（添加 4月 数据）")
    print("=" * 70)

    async with boxlite.DocGeneratorBox() as docgen:
        # 1. 恢复工作区
        print("\n[1/5] 恢复工作区...")
        local_workspace = "./workspaces/ecommerce_q1_2024"
        workspace = "/workspace/ecommerce_q1_2024"
        await docgen.load_workspace(workspace, local_workspace)
        print(f"✓ 工作区已恢复: {workspace}")

        # 2. 读取现有 PPT
        print("\n[2/5] 读取现有 PPT...")
        ppt_path = f"{workspace}/output/q1_report.pptx"
        content = await docgen.read_ppt(ppt_path)
        print(f"✓ PPT 包含 {content['slide_count']} 张幻灯片")
        for slide in content['slides']:
            print(f"  - Slide {slide['slide_id']}: {slide['title']}")

        # 3. 修改 PPT（添加 4月 数据）
        print("\n[3/5] 修改 PPT（添加 4月 数据幻灯片）...")
        await docgen.modify_ppt(
            input_path=ppt_path,
            output_path=f"{workspace}/output/q1_q2_report.pptx",
            operations=[
                {
                    "op": "update_slide",
                    "slide_id": 0,
                    "title": "Q1-Q2 销售概况（更新）",
                    "content": {
                        "text": "Q1 总销售额: ¥465万\\nQ2 总销售额: ¥523万（预测）\\n\\n整体增长: 35%"
                    }
                },
                {
                    "op": "append_slide",
                    "title": "4月 数据速报",
                    "content": {
                        "text": "销售额: ¥178万\\n订单数: 10,200单\\n客单价: ¥175\\n\\n环比增长: 8.5%"
                    }
                }
            ]
        )
        print("✓ PPT 已更新")

        # 4. 读取 Excel 数据
        print("\n[4/5] 读取 Excel 数据...")
        excel_path = f"{workspace}/output/q1_data.xlsx"
        excel_data = await docgen.read_excel(excel_path, "销售数据")
        print(f"✓ Excel 包含 {len(excel_data['sheets']['销售数据'])} 行数据")

        # 5. 修改 Excel（添加 4月 数据）
        print("\n[5/5] 修改 Excel（添加 4月 数据行）...")
        await docgen.modify_excel(
            input_path=excel_path,
            output_path=f"{workspace}/output/q1_q2_data.xlsx",
            operations=[
                {
                    "op": "append_row",
                    "sheet": "销售数据",
                    "data": {"月份": "4月", "销售额": 1780000, "订单数": 10200, "客单价": 175}
                }
            ]
        )
        print("✓ Excel 已更新")

        # 6. 下载更新后的文件到本地
        print("\n[下载] 下载更新后的文件...")
        output_dir = Path("./updated_reports")
        output_dir.mkdir(exist_ok=True)

        await docgen.download_file(
            remote_path=f"{workspace}/output/q1_q2_report.pptx",
            local_path=str(output_dir / "q1_q2_report.pptx")
        )
        await docgen.download_file(
            remote_path=f"{workspace}/output/q1_q2_data.xlsx",
            local_path=str(output_dir / "q1_q2_data.xlsx")
        )
        print(f"✓ 文件已下载到 {output_dir}")


async def scenario_3_upload_modify_download():
    """
    场景3：上传现有文件 → 修改 → 下载

    流程：
    1. 上传本地已有的报告
    2. 读取内容
    3. 批量修改（替换文本）
    4. 下载修改后的文件
    """
    print("\n" + "=" * 70)
    print("场景 3: 上传现有文件并批量修改")
    print("=" * 70)

    async with boxlite.DocGeneratorBox() as docgen:
        # 假设我们有一个现有的 Word 文档需要更新
        print("\n[1/4] 生成示例 Word 文档...")
        sections = [
            {
                "heading": "市场分析",
                "paragraphs": [
                    "Q1 市场表现良好，销售额达到预期目标。",
                    "电子产品类目增长迅速，成为主要增长点。"
                ]
            }
        ]
        await docgen.generate_word_report(
            title="Q1 市场分析报告",
            sections=sections,
            output_path="/workspace/temp_report.docx"
        )
        print("✓ 示例文档已生成")

        # 2. 下载到本地（模拟现有文件）
        print("\n[2/4] 下载到本地（模拟现有文件）...")
        local_file = "./temp_report.docx"
        await docgen.download_file("/workspace/temp_report.docx", local_file)
        print(f"✓ 文件已下载: {local_file}")

        # 3. 重新上传并修改（将 Q1 改为 Q2）
        print("\n[3/4] 上传并批量修改文本...")
        await docgen.upload_file(local_file, "/workspace/uploaded_report.docx")
        print("✓ 文件已上传")

        # 读取内容
        content = await docgen.read_word("/workspace/uploaded_report.docx")
        print(f"  原文档包含 {content['paragraph_count']} 个段落")

        # 批量替换 Q1 → Q2
        await docgen.modify_word(
            input_path="/workspace/uploaded_report.docx",
            output_path="/workspace/modified_report.docx",
            operations=[
                {"op": "replace_text", "find": "Q1", "replace": "Q2"},
                {"op": "append_paragraph", "text": "\\n注：本报告已自动更新为Q2数据。"}
            ]
        )
        print("✓ 文档已修改（Q1 → Q2）")

        # 4. 下载修改后的文件
        print("\n[4/4] 下载修改后的文件...")
        await docgen.download_file(
            "/workspace/modified_report.docx",
            "./modified_q2_report.docx"
        )
        print("✓ 修改后的文件已下载: ./modified_q2_report.docx")

        # 清理临时文件
        await docgen.delete_file("/workspace/temp_report.docx")
        await docgen.delete_file("/workspace/uploaded_report.docx")
        await docgen.delete_file("/workspace/modified_report.docx")
        print("\n✓ 容器内临时文件已清理")


async def scenario_4_batch_processing():
    """
    场景4：批量处理多个文件

    流程：
    1. 批量上传多个月份的数据文件
    2. 读取并合并数据
    3. 生成汇总报告
    """
    print("\n" + "=" * 70)
    print("场景 4: 批量处理多月份数据")
    print("=" * 70)

    async with boxlite.DocGeneratorBox() as docgen:
        # 1. 创建临时工作区
        workspace = await docgen.create_workspace("batch_processing")
        print(f"\n[1/3] 工作区已创建: {workspace}")

        # 2. 生成并上传多个月份的数据
        print("\n[2/3] 生成多个月份的数据文件...")
        months_data = {
            "1月销售": [
                {"商品": "智能手机", "销量": 120, "金额": 35万},
                {"商品": "笔记本", "销量": 80, "金额": 48万},
            ],
            "2月销售": [
                {"商品": "智能手机", "销量": 135, "金额": 40万},
                {"商品": "笔记本", "销量": 95, "金额": 57万},
            ],
            "3月销售": [
                {"商品": "智能手机", "销量": 160, "金额": 47万},
                {"商品": "笔记本", "销量": 110, "金额": 66万},
            ]
        }

        # 为每个月生成 Excel 文件
        for month, data in months_data.items():
            await docgen.generate_excel_report(
                sheets={month: data},
                output_path=f"{workspace}/input/{month}.xlsx"
            )
        print("✓ 已生成 3 个月份的 Excel 文件")

        # 3. 读取并合并数据
        print("\n[3/3] 读取并合并数据，生成汇总报告...")
        all_data = []
        files = await docgen.list_files(f"{workspace}/input", "*.xlsx")

        for file_info in files:
            data = await docgen.read_excel(file_info['path'])
            sheet_name = list(data['sheets'].keys())[0]
            all_data.extend(data['sheets'][sheet_name])

        # 生成汇总 Excel
        await docgen.generate_excel_report(
            sheets={"Q1汇总": all_data},
            output_path=f"{workspace}/output/q1_summary.xlsx"
        )
        print(f"✓ 汇总报告已生成（共 {len(all_data)} 条记录）")

        # 下载汇总报告
        await docgen.download_file(
            f"{workspace}/output/q1_summary.xlsx",
            "./q1_summary.xlsx"
        )
        print("✓ 汇总报告已下载: ./q1_summary.xlsx")


async def main():
    """运行所有场景"""
    print("\n" + "=" * 70)
    print("   DocGeneratorBox 完整工作流演示")
    print("   电商数据报告生成与管理")
    print("=" * 70)

    # 场景 1: 首次生成报告
    await scenario_1_initial_report_generation()

    # 场景 2: 定期更新报告
    await scenario_2_update_monthly_report()

    # 场景 3: 上传-修改-下载
    await scenario_3_upload_modify_download()

    # 场景 4: 批量处理
    await scenario_4_batch_processing()

    print("\n" + "=" * 70)
    print("✓ 所有场景演示完成！")
    print("=" * 70)

    print("\n生成的文件:")
    print("  ./workspaces/ecommerce_q1_2024/  - 工作区（可恢复）")
    print("  ./updated_reports/                - 更新后的报告")
    print("  ./modified_q2_report.docx         - 修改后的Word文档")
    print("  ./q1_summary.xlsx                 - 汇总报告")

    print("\n关键功能总结:")
    print("  ✅ 文件上传/下载 - upload_file(), download_file()")
    print("  ✅ 文件管理 - list_files(), delete_file()")
    print("  ✅ 读取文档 - read_ppt(), read_word(), read_excel()")
    print("  ✅ 修改文档 - modify_ppt(), modify_word(), modify_excel()")
    print("  ✅ 生成文档 - generate_ppt(), generate_excel_report()")
    print("  ✅ 工作区管理 - create_workspace(), save_workspace(), load_workspace()")

    print("\n下一步:")
    print("  1. 构建 Docker 镜像: docker build -t boxlite/doc-generator dockerfiles/docgenerator/")
    print("  2. 运行示例: python examples/python/complete_workflow_example.py")
    print("  3. 集成到您的Mediacrawler项目中")


if __name__ == "__main__":
    asyncio.run(main())
