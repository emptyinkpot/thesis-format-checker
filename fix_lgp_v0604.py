"""精确定点修正 202213210刘高朋V0604_结构收敛版.docx

只做7项修改，不全文清洗格式。
"""

import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm

SRC = Path(r"C:/Users/ASUS-KL/Downloads/202213210刘高朋V0604_结构收敛版.docx")
OUT = Path(r"C:/Users/ASUS-KL/Downloads/202213210刘高朋V0604_格式修正版.docx")

# --- 摘要三段 ---
ABSTRACT_P1 = (
    "随着建筑节能标准的提高和室内密闭程度的加深，教室、宿舍、办公室等小型密闭空间的空气质量问题日益突出。"
    "二氧化碳（CO₂）作为衡量室内通风状况的关键指标，其浓度超标不仅导致注意力下降和疲劳感增强，"
    "长期暴露还可能引发头痛、呼吸不适等健康问题。传统的便携式CO₂检测仪多采用定时人工巡检方式，"
    "无法满足连续监测与即时预警的实际需求。因此，研制一种低成本、小体积、支持远程数据传输的CO₂在线监测与预警系统"
    "具有明确的工程应用价值。本文围绕上述需求，设计了一种基于STM32的二氧化碳监测与预警器。"
    "系统以STM32F103C8T6微控制器为核心，采用Sensirion SCD41光声传感器通过I²C总线采集CO₂浓度、温度和湿度数据，"
    "传感器量程为400～5000 ppm，标称精度为±(40 ppm + 5%)。"
    "现场状态提示由0.96英寸OLED显示屏、双色LED指示灯和有源蜂鸣器协同完成；"
    "远程通信模块选用ESP8266-01S，通过MQTT协议将监测数据上传至云端服务器，"
    "移动端应用提供实时浓度图表、历史趋势分析和阈值预警推送功能。"
)

ABSTRACT_P2 = (
    "设计过程围绕硬件电路、软件架构、功能实现与系统测试四个层面展开。"
    "硬件部分完成了传感采集模块、主控最小系统、OLED与报警模块、Wi-Fi无线通信模块、"
    "电源管理与外部接口等五个功能单元的原理图设计与PCB布局。"
    "软件部分基于标准外设库开发，完成了I²C驱动与SCD41数据读取、8位CRC校验、"
    "双阈值回差判断逻辑（一级预警阈值1000 ppm、二级预警阈值1500 ppm、回差量100 ppm）、"
    "JSON格式数据帧构造、MQTT消息发布以及移动端WebSocket实时刷新等核心流程，"
    "确保本地分级报警与远程显示状态保持一致。"
)

ABSTRACT_P3 = (
    "测试结果表明，样机能够稳定完成CO₂浓度连续采集、分级声光预警、OLED实时显示和Wi-Fi数据上传等设计功能。"
    "以标准气体为参考，典型测量点的相对误差约为1.49%～1.86%，满足室内空气质量监测的精度要求；"
    "在连续24小时通信测试中，MQTT消息丢包率约为0.61%～0.72%，通信可靠性良好。"
    "后续工作仍需围绕传感器长期标定漂移补偿、高湿度和密集人流极端场景验证、结构一体化与低功耗优化设计，"
    "以及面向多教室部署的多设备协同运维管理等方面继续改进完善。"
)

# --- 外文译文追加段落 ---
TRANS_P1 = (
    "室内空气质量已成为公共卫生领域广泛关注的议题。世界卫生组织的研究报告指出，"
    "人类平均将超过百分之八十的时间用于室内活动，而建筑物的高气密性设计虽然提升了能源利用效率，"
    "却显著降低了自然通风能力，导致室内污染物累积速度加快。"
    "在各类室内空气质量参数中，二氧化碳浓度是评估通风充分性的核心指标。"
    "当室内CO₂浓度超过1000 ppm时，在场人员将出现注意力分散和困倦感；"
    "超过1500 ppm时可能引发头痛和呼吸不适。因此，对室内CO₂浓度进行连续、实时的监测具有明确的实际意义。"
)

TRANS_P2 = (
    "传统的室内空气质量评估依赖便携式仪表进行人工定期巡检，"
    "该方式存在采样间隔过长、无法捕获瞬态变化、缺乏历史数据积累等固有局限。"
    "物联网技术的快速发展为连续、自动化的环境监测提供了技术基础。"
    "本研究提出了一种面向室内空气质量实时监测的模块化物联网平台，"
    "其设计目标在于实现传感数据的自动采集、无线传输、云端存储与移动终端可视化，"
    "同时保持系统架构的模块化特征以便于后续功能扩展和多节点规模化部署。"
)

TRANS_P3 = (
    "本平台采用四层模块化架构设计，自底向上依次为传感层、通信层、服务层和应用层。"
    "传感层负责环境参数的物理量采集与数字化转换；通信层负责将采集数据通过无线链路传输至服务端；"
    "服务层完成消息接收、数据持久化存储和应用程序接口的提供；"
    "应用层面向终端用户，提供实时数据展示、历史查询与阈值预警等功能。"
    "各层之间通过标准化协议与数据格式进行松耦合连接，任何一层的实现更替均不影响其他层的正常运行。"
)

TRANS_P4 = (
    "传感器节点是整个监测系统的数据源头，其设计涵盖微控制器选型、传感器接口电路和数据采集策略三个方面。"
    "在微控制器选型方面，本平台采用基于ARM Cortex-M3内核的STM32系列微控制器，工作频率可达72 MHz，"
    "内置丰富的外设接口包括I²C、SPI和USART等，并具有良好的低功耗模式支持。"
    "传感器接口方面，平台集成了基于光声原理的CO₂传感器模块，通过I²C总线与微控制器进行数据通信，"
    "测量范围覆盖400至5000 ppm，标称测量精度为读数的百分之五与40 ppm二者之和。"
    "每次数据读取后均执行8位CRC校验以确保传输数据的完整性。"
)

TRANS_P5 = (
    "通信层采用Wi-Fi无线模块实现传感器节点与服务端之间的数据链路，"
    "通过AT指令集或透传模式与微控制器串口连接。"
    "在应用层协议方面，平台采用MQTT消息队列遥测传输协议，"
    "这是一种轻量级的发布订阅模式消息传输协议，专为受限网络环境设计。"
    "数据帧格式采用JSON结构化编码，每帧包含设备标识符、时间戳、CO₂浓度值、温度值、湿度值以及设备状态标志等字段。"
)

TRANS_P6 = (
    "服务层由消息代理、持久化数据库和RESTful API接口三个核心组件构成。"
    "服务端同时实现了阈值预警逻辑，当某一节点上报的浓度值超过预设门限时即时生成预警事件。"
    "预警逻辑引入回差机制以避免浓度在阈值附近波动时产生频繁的误报警。"
    "应用层通过WebSocket协议与服务端建立持久连接，实现监测数据的实时推送与展示，"
    "用户界面包含实时浓度数值与趋势图表、多参数联合展示、历史数据曲线查询和预警事件列表等功能模块。"
)

TRANS_P7 = (
    "实验验证表明，在多个典型浓度测量点上系统的测量精度处于传感器标称规格范围之内，"
    "能够满足室内空气质量监测的工程应用需求。"
    "在连续多小时的数据上报测试中，系统的通信可靠性可达百分之九十九以上。"
    "未来工作将围绕引入机器学习算法实现浓度变化趋势的预测、"
    "优化节点功耗管理以支持纯电池长期供电、"
    "扩展多类型传感器以实现综合空气质量评估等方向展开。"
)

def _copy_paragraph_format(src_para, new_para):
    """Copy paragraph-level formatting (style, pPr) from src to new paragraph."""
    new_para.style = src_para.style
    # Copy pPr XML if exists
    src_ppr = src_para._element.find(qn('w:pPr'))
    if src_ppr is not None:
        import copy
        new_ppr = copy.deepcopy(src_ppr)
        existing = new_para._element.find(qn('w:pPr'))
        if existing is not None:
            new_para._element.remove(existing)
        new_para._element.insert(0, new_ppr)


def _copy_run_format(src_run, new_run):
    """Copy run-level formatting (rPr) from src to new run."""
    src_rpr = src_run._element.find(qn('w:rPr'))
    if src_rpr is not None:
        import copy
        new_rpr = copy.deepcopy(src_rpr)
        existing = new_run._element.find(qn('w:rPr'))
        if existing is not None:
            new_run._element.remove(existing)
        new_run._element.insert(0, new_rpr)


def main():
    print(f"源文件: {SRC}")
    print(f"输出: {OUT}")

    # Step 0: copy
    shutil.copy2(SRC, OUT)
    doc = Document(str(OUT))

    # === 1. 页边距 ===
    for sec in doc.sections:
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.5)
    print("[1] 页边距已修正")

    # === 2. Normal style sz: 20 -> 24 ===
    normal = doc.styles['Normal']
    rpr = normal.element.find(qn('w:rPr'))
    if rpr is not None:
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            sz.set(qn('w:val'), '24')
        szCs = rpr.find(qn('w:szCs'))
        if szCs is not None:
            szCs.set(qn('w:val'), '24')
    print("[2] Normal style sz 20->24")

    # === 3. Heading 1 style sz: 30 -> 32 ===
    h1 = doc.styles['Heading 1']
    rpr_h1 = h1.element.find(qn('w:rPr'))
    if rpr_h1 is not None:
        sz = rpr_h1.find(qn('w:sz'))
        if sz is not None:
            sz.set(qn('w:val'), '32')
        szCs = rpr_h1.find(qn('w:szCs'))
        if szCs is not None:
            szCs.set(qn('w:val'), '32')
    print("[3] Heading 1 sz 30->32")

    # === 4. 页眉文字：最后一个section ===
    sections = list(doc.sections)
    last_sec = sections[-1]
    header = last_sec.header
    for p in header.paragraphs:
        if "华北水利水电大学毕业设计" in p.text and "论文" not in p.text:
            # Replace text in existing runs preserving formatting
            full_text = p.text
            new_text = full_text.replace("华北水利水电大学毕业设计", "华北水利水电大学毕业设计（论文）")
            # Clear all runs and set new text in first run
            if p.runs:
                first_run = p.runs[0]
                for run in p.runs[1:]:
                    run.text = ""
                first_run.text = new_text
            print(f"[4] 页眉已修正: '{full_text}' -> '{new_text}'")
            break
    else:
        # Try all sections
        found = False
        for sec in sections:
            for p in sec.header.paragraphs:
                if "华北水利水电大学毕业设计" in p.text and "（论文）" not in p.text:
                    full_text = p.text
                    new_text = full_text.replace("华北水利水电大学毕业设计", "华北水利水电大学毕业设计（论文）")
                    if p.runs:
                        first_run = p.runs[0]
                        for run in p.runs[1:]:
                            run.text = ""
                        first_run.text = new_text
                    found = True
                    print(f"[4] 页眉已修正 (section): '{full_text}' -> '{new_text}'")
                    break
            if found:
                break
        if not found:
            print("[4] 警告: 未找到需修正的页眉")

    # === 5. 第7章标题样式重指派 ===
    reassigned = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if ("第7章" in text or "第七章" in text) and "总结" in text:
            if p.style.name != 'Heading 1':
                p.style = doc.styles['Heading 1']
                reassigned += 1
        elif text.startswith("7.1") and "全文总结" in text:
            if p.style.name != 'Heading 2':
                p.style = doc.styles['Heading 2']
                reassigned += 1
        elif text.startswith("7.2") and "后续展望" in text:
            if p.style.name != 'Heading 2':
                p.style = doc.styles['Heading 2']
                reassigned += 1
    print(f"[5] 第7章标题重指派: {reassigned} 个段落")

    # === 6. 摘要扩写 ===
    # Find abstract section: between "摘要"/"摘 要" title and "关键词"
    abstract_start = None
    abstract_end = None
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        text = p.text.strip()
        if text.replace(" ", "").replace("　", "") == "摘要":
            abstract_start = i + 1
        elif abstract_start is not None and "关键词" in text:
            abstract_end = i
            break

    if abstract_start is not None and abstract_end is not None:
        # Get formatting reference from first abstract paragraph
        ref_para = paras[abstract_start]
        ref_style = ref_para.style
        ref_run = ref_para.runs[0] if ref_para.runs else None

        # Remove existing abstract paragraphs (between title and keywords)
        # We work with the document body element directly
        body = doc.element.body
        elements_to_remove = []
        for idx in range(abstract_start, abstract_end):
            elements_to_remove.append(paras[idx]._element)
        for elem in elements_to_remove:
            body.remove(elem)

        # Insert new paragraphs before keywords paragraph
        # After removal, the keywords paragraph is now at the position
        # We need to re-find it
        keywords_elem = None
        for p in doc.paragraphs:
            if "关键词" in p.text.strip():
                keywords_elem = p._element
                break

        if keywords_elem is not None:
            from docx.oxml import OxmlElement
            import copy

            new_texts = [ABSTRACT_P1, ABSTRACT_P2, ABSTRACT_P3]
            for txt in new_texts:
                new_p = OxmlElement('w:p')
                # Copy pPr from reference
                ref_ppr = ref_para._element.find(qn('w:pPr'))
                if ref_ppr is not None:
                    new_p.append(copy.deepcopy(ref_ppr))
                # Create run
                new_r = OxmlElement('w:r')
                if ref_run is not None:
                    ref_rpr = ref_run._element.find(qn('w:rPr'))
                    if ref_rpr is not None:
                        new_r.append(copy.deepcopy(ref_rpr))
                new_t = OxmlElement('w:t')
                new_t.text = txt
                new_t.set(qn('xml:space'), 'preserve')
                new_r.append(new_t)
                new_p.append(new_r)
                keywords_elem.addprevious(new_p)

            print(f"[6] 摘要已替换为3段 ({len(ABSTRACT_P1)+len(ABSTRACT_P2)+len(ABSTRACT_P3)} chars)")
        else:
            print("[6] 错误: 找不到关键词段落")
    else:
        print(f"[6] 警告: 未找到摘要区间 (start={abstract_start}, end={abstract_end})")

    # === 7. 外文译文追加 ===
    # Find the last paragraph in the foreign translation section
    trans_start = None
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if "附录二" in text or "外文译文" in text:
            trans_start = i
            break

    if trans_start is not None:
        # Find the last non-empty paragraph in this section
        # (before next appendix or end of document)
        last_trans_para = None
        for i in range(trans_start + 1, len(doc.paragraphs)):
            text = doc.paragraphs[i].text.strip()
            if text.startswith("附录三") or text.startswith("致谢") or text.startswith("致 谢"):
                break
            if text:
                last_trans_para = doc.paragraphs[i]

        if last_trans_para is not None:
            from docx.oxml import OxmlElement
            import copy

            # Reference formatting from the last translation paragraph
            ref_p2 = last_trans_para
            ref_style2 = ref_p2.style
            ref_run2 = ref_p2.runs[0] if ref_p2.runs else None
            insert_after = last_trans_para._element

            trans_texts = [TRANS_P1, TRANS_P2, TRANS_P3, TRANS_P4, TRANS_P5, TRANS_P6, TRANS_P7]
            for txt in trans_texts:
                new_p = OxmlElement('w:p')
                ref_ppr = ref_p2._element.find(qn('w:pPr'))
                if ref_ppr is not None:
                    new_p.append(copy.deepcopy(ref_ppr))
                new_r = OxmlElement('w:r')
                if ref_run2 is not None:
                    ref_rpr = ref_run2._element.find(qn('w:rPr'))
                    if ref_rpr is not None:
                        new_r.append(copy.deepcopy(ref_rpr))
                new_t = OxmlElement('w:t')
                new_t.text = txt
                new_t.set(qn('xml:space'), 'preserve')
                new_r.append(new_t)
                new_p.append(new_r)
                insert_after.addnext(new_p)
                insert_after = new_p

            print(f"[7] 外文译文追加7段")
        else:
            print("[7] 警告: 未找到外文译文内容段落")
    else:
        print("[7] 警告: 未找到'附录二'或'外文译文'标记")

    # === Save ===
    doc.save(str(OUT))
    print(f"\n保存完成: {OUT}")


if __name__ == "__main__":
    main()
