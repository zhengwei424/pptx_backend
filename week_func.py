import pptx
import pptx.slide
import pptx.table
import pptx.text.text
import pptx.shapes.base
import pptx.chart.chart
from pptx.util import Pt, Cm, Emu
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PP_ALIGN
from pptx.enum.chart import XL_LEGEND_POSITION, XL_CHART_TYPE, XL_DATA_LABEL_POSITION


# presentation -> slide -> shapes -> placeholder,graphfrm -> chart(table -> cell) -> text_frame -> paragraphs -> font
#            |      +-> placeholder
#            |->slide_master -> slide_layout
# https://mhammond.github.io/pywin32/html/com/win32com/HTML/QuickStartClientCom.html

class PresentationBuilder(object):
    def __init__(self, prs_path):
        self.prs_path = prs_path
        self.prs = Presentation(self.prs_path)

    @property
    def xml_slides(self):
        return self.prs.slides._sldIdLst

    # 插入slide_layout到指定位置
    def insert_slide_by_slide_layout(self, index, slide_layout):
        # type -> pptx.oxml.presentation.CT_SlideIdList
        slideIdList = list(self.xml_slides)
        self.prs.slides.add_slide(slide_layout)
        pop_item = slideIdList.pop()
        self.xml_slides.insert(index, pop_item)

    # 复制一个slide到指定位置
    def copy_slide(self):
        pass

    # 删除指定位置的slide
    def delete_slide(self, index):
        slideIdList = list(self.xml_slides)
        self.xml_slides.remove(slideIdList[index])


# 先写内容，后改样式，否则可能出现样式不生效问题，如字体之类的？？？
def set_cell_format(cell: pptx.table._Cell,
                    font_name="Microsoft YaHei",
                    font_size=8,
                    font_bold=False,
                    align=PP_ALIGN.LEFT,
                    font_color="000000",
                    cell_background_color=None,
                    margin_left=None,
                    margin_right=None,
                    margin_top=None,
                    margin_bottom=None):
    # 语言设置，NONE表示移除所有语言设置
    cell.text_frame.paragraphs[0].font.language_id = MSO_LANGUAGE_ID.NONE
    # 字体
    cell.text_frame.paragraphs[0].font.name = font_name
    # 字体大小
    cell.text_frame.paragraphs[0].font.size = Pt(int(font_size))
    # 是否加粗
    cell.text_frame.paragraphs[0].font.bold = font_bold
    # 水平对齐
    cell.text_frame.paragraphs[0].alignment = align

    # cell中文字相对cell边框的距离，类似与html中的padding
    cell.margin_left = margin_left
    cell.margin_right = margin_right
    cell.margin_top = margin_top
    cell.margin_bottom = margin_bottom

    # 用RGB表示字体颜色（两种方式）
    # cell.text_frame.paragraphs[0].font.color.rgb = RGBColor.from_string(fontColor)
    # 设置前景色或背景色之前需要执行
    # cell.text_frame.paragraphs[0].font.fill.patterned()  # 图案填充（可以设置前景色和背景色）
    cell.text_frame.paragraphs[0].font.fill.solid()  # 纯色填充（只能设置前景色）
    # 字体前景色(就是字体颜色）
    cell.text_frame.paragraphs[0].font.fill.fore_color.rgb = RGBColor.from_string(font_color)
    # 字体背景色（就是文字本身的背景，不是整个cell）
    # cell.text_frame.paragraphs[0].font.fill.back_color.rgb = RGBColor.from_string(cell_background_color)
    # 字体颜色透明度
    # cell.text_frame.paragraphs[0].font.color.brightness = -1  # 取值范围-1~1，暗->亮
    if cell_background_color:
        # 填充cell前景色
        # cell.fill.patterned()  # 图案填充（可以设置cell前景色和背景色）
        cell.fill.solid()  # 纯色填充（只能设置cell前景色，但是对cell中的text而言也是背景）
        # cell.fill.back_color.rgb = RGBColor.from_string(cell_background_color)
        cell.fill.fore_color.rgb = RGBColor.from_string(cell_background_color)
    # else:
    #     # cell颜色无填充
    #     cell.fill.background()


class TableAttribute:
    def __init__(self, table):
        self.tb = table  # type: pptx.table.Table


class WeaklyReports(object):
    def __init__(self, prs):
        self._prs = prs

    # 1. 运维工作统计（次数）
    def slide_1(self, events_count: list):  #
        if len(events_count) != 6:
            print("events_count length is 6,Please check out.")
        slide = self._prs.slides[8]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        index = 0
        for shape in slide.shapes:
            '''
            # 判断shape是否是图表
            if shape.has_chart:
                # https://python-pptx.readthedocs.io/en/latest/user/charts.html
                chart = sl1.shapes[index].chart  # type: pptx.chart.chart.Chart
                # 判断chart类型
                if chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:  # 柱状图
                    chart_data = CategoryChartData()
                    chart_data.categories = ["a", "b", "c"]  # 类别
                    chart_data.add_series('Series 1', (19.2, 21.4, 16.7))  # 系列（即每个类别可以有多个系列）
    
                if chart.chart_type == XL_CHART_TYPE.PIE:  # 饼图
                    chart_data = CategoryChartData()
                    chart_data.categories = ['West', 'East', 'North', 'South', 'Other']  # 类别
                    chart_data.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))  # 系列（即每个类别可以有多个系列）
            # 判断shape是否包含text
            if shape.has_text_frame:
                tf = sl1.shapes[index].text_frame  # type: pptx.text.text.TextFrame
            '''

            # 判断shape是否是表格（即找到需要修改的表格）
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                tb.cell(3, 1).text = str(events_count[0])  # 变更
                tb.cell(3, 2).text = str(events_count[1])  # 资源权限管理
                tb.cell(3, 3).text = str(events_count[2])  # 配合操作
                tb.cell(3, 4).text = str(events_count[3])  # 支撑发版
                tb.cell(3, 5).text = str(events_count[4])  # 问题和告警处理
                tb.cell(3, 6).text = str(events_count[5])  # 故障处理
            index += 1

    # 2. 巡检
    def slide_2(self, weekly_inspect: list):
        if len(weekly_inspect) != 17:
            print("weekly_inspect length is 17,Please check out.")
        slide = self._prs.slides[9]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                test = TableAttribute(tb)
                tb.cell(3, 1).text = "11"  # 巡检次数
                set_cell_format(tb.cell(3, 1), fontBold=False, fontColor="3C6F6A", cellbgColor="CDC839")
                tb.cell(3, 2).text = "0"  # 异常次数
                tb.cell(3, 3).text = "6"  # 报告提交次数
                tb.cell(3, 5).text = "√"  # 周一上午
                tb.cell(4, 5).text = "√"  # 周一下午
                tb.cell(3, 7).text = "√"  # 周二上午
                tb.cell(4, 7).text = "√"  # 周二下午
                tb.cell(3, 9).text = "√"  # 周三上午
                tb.cell(4, 9).text = "√"  # 周三下午
                tb.cell(3, 11).text = "√"  # 周四上午
                tb.cell(4, 11).text = "√"  # 周四下午
                tb.cell(3, 13).text = "√"  # 周五上午
                tb.cell(4, 13).text = "√"  # 周五下午
                tb.cell(3, 15).text = "√"  # 周六上午
                tb.cell(4, 15).text = "√"  # 周六下午
                tb.cell(3, 17).text = "√"  # 周日上午
                tb.cell(4, 17).text = "√"  # 周日下午

    # 3. 变更
    def slide_3(self, content: list):
        slide = self._prs.slides[10]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 4. 支撑发版
    def slide_4(self, content: list):
        slide = self._prs.slides[11]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 5. 资源权限管理
    def slide_5(self, content: list):
        slide = self._prs.slides[12]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 6. 配合操作及排障
    def slide_6(self, content: list):
        slide = self._prs.slides[13]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 7. 问题及告警
    def slide_7(self, content: list):
        slide = self._prs.slides[14]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]

    # 8. 运行情况分析
    def slide_8(self):
        # slide = self._prs.slides[18]  # type:pptx.slide.Slide
        # shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        # for shape in slide.shapes:
        #     # if shape.has_table: # 获取现有shape的坐标及基本参数（宽度高度等）
        #     #     print("shape: {}, left: {}, top: {}, height: {}, width: {}".format(shape, shape.left/360000, shape.top/360000,
        #     #                                                                        shape.height/360000, shape.width/360000))
        #     if shape.has_chart:
        #         chart = shape.chart  # type: pptx.chart.chart.Chart
        #         print(chart.chart_title.text_frame.text)
        # 传入集群应用部署信息
        cluster_app_deploy_info = {"fcp": [["\u9879\u76ee\u540d\u79f0", "\u670d\u52a1\u540d\u79f0", "\u5b9e\u4f8b\u6570"], ["\u914d\u7f6e\u4e2d\u5fc3-\u7eff\u533a", "", 1], ["asm\u4e8b\u540e\u76d1\u7763\u7cfb\u7edf-\u7eff\u533a", "asmview", 2], ["asm\u4e8b\u540e\u76d1\u7763\u7cfb\u7edf-\u7eff\u533a", "asmservice", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "sg-collector", 4], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "auth-service", 4], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "unite-codeadmin", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "sg-view", 1], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "auth-permission", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "unite-manager", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "sg-gateway", 4], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "auth-sec-manager", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "sg-gateway-admin", 4], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "auth-gateway-api", 2], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "auth-sec-service", 4], ["\u5e73\u53f0\u670d\u52a1-\u7eff\u533a", "sg-cmp-backend", 2], ["\u5e73\u53f0\u57fa\u7840\u670d\u52a1-\u7eff\u533a", "fcpcode", 1], ["\u5e73\u53f0\u57fa\u7840\u670d\u52a1-\u7eff\u533a", "pubdata", 3], ["bpm\u6d41\u7a0b\u5f15\u64ce-\u7eff\u533a", "bpmservice", 6], ["bpm\u6d41\u7a0b\u5f15\u64ce-\u7eff\u533a", "bpmview", 6], ["BPS\u6d41\u7a0b\u5f15\u64ce-\u7eff\u533a", "", 1], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-queue", 2], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-manager", 1], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-pinecone", 2], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-bot-manager", 2], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-fe", 2], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-auth", 1], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "commander-uic", 2], ["rpa\u673a\u5668\u4eba-\u7eff\u533a", "nacos", 1], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "cpm-apportion", 2], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "cpm-position", 4], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "deposit", 2], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "cpm-autotask", 3], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "cpm-interbank", 4], ["cpm\u8d22\u52a1\u516c\u53f8\u5934\u5bf8-\u7eff\u533a", "cpm-liquidity", 2], ["\u76d1\u7ba1\u62a5\u9001-\u7eff\u533a", "crisview", 1], ["\u76d1\u7ba1\u62a5\u9001-\u7eff\u533a", "crisservice", 1], ["\u51b3\u7b56\u5206\u6790-\u7eff\u533a", "decisionserver", 1], ["\u51b3\u7b56\u5206\u6790-\u7eff\u533a", "decisionbi", 1], ["east\u6570\u636e\u62a5\u9001-\u7eff\u533a", "eastservice02", 1], ["east\u6570\u636e\u62a5\u9001-\u7eff\u533a", "eastservice", 1], ["\u5185\u5bb9\u5f71\u50cf-\u7eff\u533a", "ecmscommonservice", 4], ["\u5185\u5bb9\u5f71\u50cf-\u7eff\u533a", "vimservice", 4], ["\u5185\u5bb9\u5f71\u50cf-\u7eff\u533a", "ecmssearchservice", 4], ["\u5185\u5bb9\u5f71\u50cf-\u7eff\u533a", "ecmsftsservice", 4], ["\u53cd\u6d17\u94b1-\u7eff\u533a", "", 2], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "sys-file-9698", 4], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "sys-file-9699", 4], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "hellgate-mobile-9692", 4], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "hellgate-9697", 4], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "hellgate-9696", 4], ["\u5e94\u7528\u7f51\u5173-\u7eff\u533a", "sys-file-mobile-9693", 4], ["\u6838\u5fc3-\u7ed3\u7b97-\u7eff\u533a", "cmsview", 4], ["\u6838\u5fc3-\u7ed3\u7b97-\u7eff\u533a", "rdpcw2", 0], ["\u6838\u5fc3-\u7ed3\u7b97-\u7eff\u533a", "cmstransservice", 0], ["\u6838\u5fc3-\u7ed3\u7b97-\u7eff\u533a", "rdpcw", 0], ["\u6838\u5fc3-\u7ed3\u7b97-\u7eff\u533a", "", 6], ["\u6838\u5fc3-\u5ba2\u6237\u4fe1\u606f\u7ba1\u7406-\u7eff\u533a", "ecifview", 5], ["\u6838\u5fc3-\u5ba2\u6237\u4fe1\u606f\u7ba1\u7406-\u7eff\u533a", "ecifservice", 5], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "asidview", 2], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "ebankview", 4], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "asidserivce", 2], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "ebankservice", 5], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "", 4], ["\u6838\u5fc3-\u5176\u4ed6-\u7eff\u533a", "evsview", 4], ["\u6838\u5fc3-\u94f6\u4f01-\u7eff\u533a", "utipservice", 4], ["\u6838\u5fc3-\u94f6\u4f01-\u7eff\u533a", "utipjobservice", 1], ["\u6838\u5fc3-\u94f6\u4f01-\u7eff\u533a", "utippubservice", 4], ["\u6838\u5fc3-\u94f6\u4f01-\u7eff\u533a", "", 4], ["\u6d88\u606f\u4e2d\u5fc3-\u7eff\u533a", "", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-message-service", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-message-front", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-fbcm", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-greport-service", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-ebs-service", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-crm", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-ics", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-prms", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-app-center-service", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-data-transfer", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-imes", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-neams", 2], ["nfbcm\u65b0\u4fe1\u8d37-\u7eff\u533a", "nfbcm-service-adapter", 2], ["\u524d\u7aef\u6846\u67b6\u8def\u7531\u5206\u53d1-\u7eff\u533a", "nginx-web-gate", 2], ["ods\u6570\u636e\u4ed3\u5e93-\u7eff\u533a", "odsctlmagent", 1], ["ods\u6570\u636e\u4ed3\u5e93-\u7eff\u533a", "odsctlmain", 1], ["ods\u6570\u636e\u4ed3\u5e93-\u7eff\u533a", "odsctlsagent2", 1], ["ods\u6570\u636e\u4ed3\u5e93-\u7eff\u533a", "odsctlsagent1", 1], ["\u516c\u5171\u7ba1\u7406-\u7eff\u533a", "pubview", 6], ["\u516c\u5171\u7ba1\u7406-\u7eff\u533a", "pubservice", 6], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-acct-service", 4], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "fcp-archive", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-file", 4], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pay-view", 6], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pay-service", 6], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-fin-view", 2], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pos-view", 4], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pos-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-dv-view", 2], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pay-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-dv-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-bud-view", 6], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-acct-view", 4], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-dv-service", 2], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-fin-service", 2], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-fin-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-bud-service", 6], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-acct-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-bud-job", 1], ["\u96c6\u56e2\u53f8\u5e93-\u7eff\u533a", "treasurer-pos-service", 4], ["\u7edf\u4e00\u63a5\u5165\u5e73\u53f0-\u7eff\u533a", "", 6], ["xxljob\u8c03\u5ea6\u7ba1\u7406-\u7eff\u533a", "xxljob", 3], ["\u8fd0\u7ef4\u5de1\u68c0\u76d1\u63a7-\u7eff\u533a", "", 1], ["27", "112", "302"]], "fcp-inner-microservice": [["\u9879\u76ee\u540d\u79f0", "\u670d\u52a1\u540d\u79f0", "\u5b9e\u4f8b\u6570"], ["\u8d22\u52a1\u516c\u53f8\u5934\u5bf8\u7ba1\u7406", "cpm-apportion", 2], ["\u8d22\u52a1\u516c\u53f8\u5934\u5bf8\u7ba1\u7406", "cpm-liquidity", 2], ["\u8d22\u52a1\u516c\u53f8\u5934\u5bf8\u7ba1\u7406", "cpm-interbank", 4], ["\u8d22\u52a1\u516c\u53f8\u5934\u5bf8\u7ba1\u7406", "cpm-autotask", 3], ["\u8d22\u52a1\u516c\u53f8\u5934\u5bf8\u7ba1\u7406", "cpm-position", 4], ["\u7edf\u4e00\u63a5\u5165\u5e73\u53f0", "uapservice", 6], ["\u4fe1\u8d37", "nfbcm-service-adapter", 2], ["\u4fe1\u8d37", "nfbcm-data-transfer", 2], ["\u4fe1\u8d37", "nfbcm-service-imes", 2], ["\u4fe1\u8d37", "nfbcm-neams", 2], ["\u4fe1\u8d37", "nfbcm-service-prms", 2], ["\u4fe1\u8d37", "nfbcm-message-front", 2], ["\u4fe1\u8d37", "nfbcm-service-fbcm", 2], ["\u4fe1\u8d37", "nfbcm-greport-service", 2], ["\u4fe1\u8d37", "nfbcm-ebs-service", 2], ["\u4fe1\u8d37", "nfbcm-app-center-service", 2], ["\u4fe1\u8d37", "nfbcm-message-service", 2], ["\u4fe1\u8d37", "nfbcm-service-ics", 2], ["\u4fe1\u8d37", "nfbcm-service-crm", 2], ["\u6838\u5fc3", "asidserivce", 2], ["\u6838\u5fc3", "evsservice", 5], ["\u6838\u5fc3", "ebankservice", 5], ["\u6838\u5fc3", "cmstransservice", 1], ["\u6838\u5fc3", "cmsservice", 6], ["\u6838\u5fc3", "ebankview", 4], ["\u6838\u5fc3", "asidview", 2], ["\u6838\u5fc3", "cmsview", 4], ["\u6838\u5fc3", "utipfileservice", 4], ["\u6838\u5fc3", "utippubservice", 4], ["\u6838\u5fc3", "utipservice", 4], ["\u6838\u5fc3", "utipjobservice", 1], ["\u6838\u5fc3", "bedcservice", 4], ["\u5ba2\u6237\u4fe1\u606f\u7ba1\u7406\u7cfb\u7edf", "ecifservice", 5], ["\u6d88\u606f\u4e2d\u5fc3", "msgcenterservice", 2], ["\u6d88\u606f\u4e2d\u5fc3", "msgcenterserviceconsumer", 2], ["6", "35", "102"]], "fcp-inner-backend": [["\u9879\u76ee\u540d\u79f0", "\u670d\u52a1\u540d\u79f0", "\u5b9e\u4f8b\u6570"], ["\u6570\u636e\u4ed3\u5e93", "odsctlsagent1", 0], ["\u6570\u636e\u4ed3\u5e93", "odsctlmain", 0], ["\u6570\u636e\u4ed3\u5e93", "odsctlsagent2", 0], ["\u6570\u636e\u4ed3\u5e93", "odsctlmagent", 0], ["\u6838\u5fc3", "rdpcw", 1], ["\u6838\u5fc3", "rdpcw2", 1], ["\u6838\u5fc3", "v7cw", 7], ["RPA", "commander-pinecone", 1], ["RPA", "nacos", 0], ["RPA", "commander-bot-manager", 2], ["RPA", "commander-manager", 1], ["RPA", "commander-queue", 2], ["RPA", "commander-uic", 2], ["RPA", "commander-auth", 1], ["RPA", "commander-fe", 2], ["3", "15", "20"]], "fcp-outer-microservice": [["\u9879\u76ee\u540d\u79f0", "\u670d\u52a1\u540d\u79f0", "\u5b9e\u4f8b\u6570"], ["bpm-\u5ba1\u6279\u57df\u4e2d\u53f0", "bpmservice", 6], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "secretkey-service", 4], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "fcpcode-admin", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "uims-gateway-api", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "deposit", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "sg-gateway-admin", 4], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "uims-service", 4], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "sg-gateway", 4], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "sg-collector", 4], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "uims-permission", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "secretkey-manager", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "sg-view", 1], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "sg-cmp-backend", 2], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "fcpcode", 5], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "pubdata", 3], ["\u5e73\u53f0\u516c\u5171\u670d\u52a1", "conf-manager", 2], ["\u76d1\u7ba1\u62a5\u9001", "crisservice", 0], ["\u76d1\u7ba1\u62a5\u9001", "crisview", 0], ["asm\u4e8b\u540e\u76d1\u7763\u7cfb\u7edf", "asmview", 2], ["asm\u4e8b\u540e\u76d1\u7763\u7cfb\u7edf", "asmservice", 2], ["bpm-\u5ba1\u6279\u57df\u524d\u53f0", "bpmview", 6], ["\u516c\u5171\u670d\u52a1", "pubservice", 6], ["\u53f8\u5e93", "treasurer-acct-view", 4], ["\u53f8\u5e93", "treasurer-dv-view", 2], ["\u53f8\u5e93", "treasurer-pay-service", 6], ["\u53f8\u5e93", "treasurer-pos-service", 4], ["\u53f8\u5e93", "treasurer-pos-view", 4], ["\u53f8\u5e93", "treasurer-bud-service", 6], ["\u53f8\u5e93", "treasurer-fin-view", 2], ["\u53f8\u5e93", "treasurer-fin-service", 2], ["\u53f8\u5e93", "treasurer-pay-view", 6], ["\u53f8\u5e93", "treasurer-dv-service", 2], ["\u53f8\u5e93", "treasurer-bud-view", 6], ["\u53f8\u5e93", "treasurer-acct-service", 4], ["\u53f8\u5e93", "treasurer-file", 2], ["\u7535\u5b50\u51ed\u8bc1\u524d\u53f0", "evsview", 4], ["\u516c\u5171\u7ba1\u7406", "pubview", 6], ["\u5ba2\u6237\u4fe1\u606f\u7ba1\u7406\u524d\u53f0", "ecifview", 5], ["\u5185\u5bb9\u5f71\u50cf", "ecmssearchservice", 4], ["\u5185\u5bb9\u5f71\u50cf", "vimservice", 4], ["\u5185\u5bb9\u5f71\u50cf", "ecmsftsservice", 4], ["\u5185\u5bb9\u5f71\u50cf", "ecmscommonservice", 4], ["11", "42", "146"]], "fcp-outer-backend": [["\u9879\u76ee\u540d\u79f0", "\u670d\u52a1\u540d\u79f0", "\u5b9e\u4f8b\u6570"], ["ESB", "esb-server3", 1], ["ESB", "esb-server2", 1], ["ESB", "esb-server1", 1], ["ESB", "esb-dangban", 0], ["\u53f8\u5e93", "fcp-archive", 1], ["\u53f8\u5e93", "treasurer-pay-job", 1], ["\u53f8\u5e93", "treasurer-pos-job", 1], ["\u53f8\u5e93", "treasurer-acct-job", 1], ["\u53f8\u5e93", "treasurer-fin-job", 1], ["\u53f8\u5e93", "treasurer-bud-job", 1], ["\u53f8\u5e93", "treasurer-dv-job", 1], ["east\u6570\u636e\u62a5\u9001", "eastservice02", 0], ["east\u6570\u636e\u62a5\u9001", "eastservice", 0], ["\u51b3\u7b56\u5206\u6790", "decisionbi", 0], ["\u51b3\u7b56\u5206\u6790", "decisionserver", 0], ["\u53cd\u6d17\u94b1", "fxqjobadmin", 0], ["\u53cd\u6d17\u94b1", "fxqjobexecutor", 0], ["\u53cd\u6d17\u94b1", "fxqservice", 1], ["\u53cd\u6d17\u94b1", "fxqservice2", 1], ["bps-\u6d41\u7a0b\u5f15\u64ce\u540e\u53f0", "bpsworkspace", 1], ["bps-\u6d41\u7a0b\u5f15\u64ce\u540e\u53f0", "bpsserver2", 1], ["bps-\u6d41\u7a0b\u5f15\u64ce\u540e\u53f0", "bpsserver", 1], ["\u6295\u8d44\u7ba1\u7406\u7cfb\u7edf", "xquant-xir", 1], ["\u6295\u8d44\u7ba1\u7406\u7cfb\u7edf", "xquant-xir2", 1], ["\u6295\u8d44\u7ba1\u7406\u7cfb\u7edf", "xquant-calc2", 1], ["\u6295\u8d44\u7ba1\u7406\u7cfb\u7edf", "xquant-calc", 1], ["\u6295\u8d44\u7ba1\u7406\u7cfb\u7edf", "xquant-smartbi", 1], ["7", "27", "20"]]}


        for key in cluster_app_deploy_info.keys():
            # 每页最多存75(80)行数据
            table_pg_num = 0
            table_data = cluster_app_deploy_info[key]
            if key == "fcp":
                table_pg_num = 2
            else:
                table_pg_num = 1
            for page in range(1, table_pg_num + 1):
                # 添加一张幻灯片
                slide = self._prs.slides.add_slide(self._prs.slide_masters[0].slide_layouts[1])

                # 幻灯片标题
                slidePlaceholder = slide.shapes[0]  # type: pptx.shapes.placeholder.SlidePlaceholder
                slidePlaceholder.text = "2.11慧企运行情况分析–集群资源使用情况"

                # cpu 扇形图
                chart_cpu_data = ChartData()
                chart_cpu_data.categories = ['已分配', '未分配']
                chart_cpu_data.add_series('xxx1', (0.25, 0.75))
                chart_cpu_data.add_series('xxx2', (0.35, 0.65))
                chart_cpu_data.add_series('xxx3', (0.45, 0.55))
                shape_cpu = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(2.2), Cm(6), Cm(6), chart_cpu_data)
                chart_cpu = shape_cpu.chart

                # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
                chart_cpu.has_legend = True
                chart_cpu.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart_cpu.legend.include_in_layout = False
                chart_cpu.font.name = 'Microsoft YaHei'
                chart_cpu.font.size = Pt(10)

                # 设置标题（不设置，默认值是add_series中的列标题"cpu"）-> "ChartTitle" has no attribute "width"??
                chart_cpu.has_title = True
                chart_cpu.chart_title.has_text_frame = True
                if key == "fcp":
                    chart_cpu.chart_title.text_frame.text = "生产环境绿区 CPU(%)"
                elif key == "fcp-inner-microservice":
                    chart_cpu.chart_title.text_frame.text = "内网微服务区 CPU(%)"
                elif key == "fcp-inner-backend":
                    chart_cpu.chart_title.text_frame.text = "内网后台区 CPU(%)"
                elif key == "fcp-outer-microservice":
                    chart_cpu.chart_title.text_frame.text = "外网微服务区 CPU(%)"
                elif key == "fcp-outer-backend":
                    chart_cpu.chart_title.text_frame.text = "外网后台区 CPU(%)"
                chart_cpu.chart_title.text_frame.paragraphs[0].font.size = Pt(10)
                chart_cpu.chart_title.text_frame.paragraphs[0].font.name = "Microsoft YaHei"
                chart_cpu.chart_title.text_frame.paragraphs[0].font.bold = False

                chart_cpu.plots[0].has_data_labels = True
                chart_cpu_data_labels = chart_cpu.plots[0].data_labels
                chart_cpu_data_labels.number_format = '0%'
                chart_cpu_data_labels.position = XL_DATA_LABEL_POSITION.CENTER

                # memory 扇形图
                chart_mem_data = ChartData()
                chart_mem_data.categories = ['已分配', '未分配']
                chart_mem_data.add_series('内存', (0.25, 0.75))
                chart_mem = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Cm(0.1), Cm(8.3), Cm(6), Cm(6),
                                                   chart_mem_data).chart

                # 设置图例说明（会在图中标识已分配、未分配的颜色说明）
                chart_mem.has_legend = True
                chart_mem.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart_mem.legend.include_in_layout = False
                chart_mem.font.name = 'Microsoft YaHei'
                chart_mem.font.size = Pt(10)

                # 设置标题（不设置，默认值是add_series中的列标题"cpu")
                chart_mem.has_title = True
                chart_mem.chart_title.has_text_frame = True
                if key == "fcp":
                    chart_mem.chart_title.text_frame.text = "生产环境绿区 内存(%)"
                elif key == "fcp-inner-microservice":
                    chart_mem.chart_title.text_frame.text = "内网微服务区 内存(%)"
                elif key == "fcp-inner-backend":
                    chart_mem.chart_title.text_frame.text = "内网后台区 内存(%)"
                elif key == "fcp-outer-microservice":
                    chart_mem.chart_title.text_frame.text = "外网微服务区 内存(%)"
                elif key == "fcp-outer-backend":
                    chart_mem.chart_title.text_frame.text = "外网后台区 内存(%)"
                chart_mem.chart_title.text_frame.paragraphs[0].font.size = Pt(10)
                chart_mem.chart_title.text_frame.paragraphs[0].font.name = "Microsoft YaHei"
                chart_mem.chart_title.text_frame.paragraphs[0].font.bold = False

                chart_mem.plots[0].has_data_labels = True
                chart_mem_data_labels = chart_mem.plots[0].data_labels
                chart_mem_data_labels.number_format = '0%'
                chart_mem_data_labels.position = XL_DATA_LABEL_POSITION.CENTER
                # 添加table
                if page == 1 and table_pg_num == 1:
                    if len(table_data) <= 40:
                        # table_1
                        # 根据数据长度设置表格的行列数（同时通过行高*行数，计算表格高度，并与模板对比，做表格当前页分割，或换页分割）！！！！！
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1
                    elif len(table_data) > 40:
                        # table_1
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1

                        # table_2
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(16), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        # 取消表格第一行特殊格式
                        table.first_row = False
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1
                if page == 1 and table_pg_num == 2:
                    # table_1
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                    table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                    table.columns[0].width = Cm(3.5)
                    table.columns[1].width = Cm(3.5)
                    table.columns[2].width = Cm(1)
                    index = 0
                    cells = list(table.iter_cells())
                    count = len(table_data)
                    for i in range(count):
                        if index == len(cells):
                            break
                        for j in table_data.pop(0):
                            if index < len(cells):
                                cells[index].text = str(j)
                                set_cell_format(cells[index],
                                                font_size=7,
                                                margin_left=0,
                                                margin_right=0,
                                                margin_top=0,
                                                margin_bottom=0)
                                index += 1

                    # table_2
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                    x, y, cx, cy = Cm(16), Cm(2.2), Cm(1), Cm(0.2)
                    table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                    table.columns[0].width = Cm(3.5)
                    table.columns[1].width = Cm(3.5)
                    table.columns[2].width = Cm(1)
                    # 取消表格第一行特殊格式
                    table.first_row = False
                    index = 0
                    cells = list(table.iter_cells())
                    count = len(table_data)
                    for i in range(count):
                        if index == len(cells):
                            break
                        for j in table_data.pop(0):
                            if index < len(cells):
                                cells[index].text = str(j)
                                set_cell_format(cells[index],
                                                font_size=7,
                                                margin_left=0,
                                                margin_right=0,
                                                margin_top=0,
                                                margin_bottom=0)
                                index += 1
                if page == 2 and table_pg_num == 2:
                    if len(table_data) <= 40:
                        # table_1
                        # 根据数据长度设置表格的行列数（同时通过行高*行数，计算表格高度，并与模板对比，做表格当前页分割，或换页分割）！！！！！
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1
                    elif len(table_data) > 40:
                        # table_1
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(7.5), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(40, 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1

                        # table_2
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        # x: 左边距，y: 上边距, cx: 单元格宽度, cy: 单元格高度
                        x, y, cx, cy = Cm(16), Cm(2.2), Cm(1), Cm(0.2)
                        table = slide.shapes.add_table(len(table_data), 3, x, y, cx, cy).table
                        table.columns[0].width = Cm(3.5)
                        table.columns[1].width = Cm(3.5)
                        table.columns[2].width = Cm(1)
                        # 取消表格第一行特殊格式
                        table.first_row = False
                        index = 0
                        cells = list(table.iter_cells())
                        count = len(table_data)
                        for i in range(count):
                            if index == len(cells):
                                break
                            for j in table_data.pop(0):
                                if index < len(cells):
                                    cells[index].text = str(j)
                                    set_cell_format(cells[index],
                                                    font_size=7,
                                                    margin_left=0,
                                                    margin_right=0,
                                                    margin_top=0,
                                                    margin_bottom=0)
                                    index += 1
                # 合并单元格(即this_cell.merge(other_cell))
                # cell.merge(other_cell)

    # 9. 下周工作计划
    def slide_9(self, content: list):
        slide = self._prs.slides[25]  # type: pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]
