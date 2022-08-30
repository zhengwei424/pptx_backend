import pptx
import pptx.slide
import pptx.shapes.base
import pptx.table
import pptx.chart.chart
import pptx.text.text
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData


# presentation->slide->shapes->placeholder,graphfrm
#            |      +->placeholder
#            |->slide_master-slide_layout
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


def table_render(*args, **kwargs):
    prs = None
    slide = None
    # 分解kwargs
    for k, v in kwargs.items():
        if k == "prs":
            prs = kwargs[k]
        if k == "slide":
            slide = kwargs[k]

    # 获取slide的宽度和高度
    prs_width = prs.slide_width
    prs_height = prs.slide_height
    prs.slides.add_slide(prs.slide_masters[0].slide_layouts[1])


def change_table_data():
    pass


def chart_render():
    pass


def change_chart_data():
    pass


class WeaklyReports(object):
    def __init__(self, prs):
        self._prs = prs

    # 1. 运维工作统计（次数）
    def slide_1(self, events_count: list):  #
        if len(events_count) != 6:
            print("events_count have wrong length")
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
    def slide_2(self):
        slide = self._prs.slides[9]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                tb.cell(3, 1).text = "11"  # 巡检次数
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
        slide = self._prs.slides[18]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            # if shape.has_table: # 获取现有shape的坐标及基本参数（宽度高度等）
            #     print("shape: {}, left: {}, top: {}, height: {}, width: {}".format(shape, shape.left/360000, shape.top/360000,
            #                                                                        shape.height/360000, shape.width/360000))
            if shape.has_chart:
                chart = shape.chart  # type: pptx.chart.chart.Chart
                print(chart.chart_title.text_frame.text)

    # 9. 下周工作计划
    def slide_9(self, content: list):
        slide = self._prs.slides[25]  # type:pptx.slide.Slide
        shape = slide.shapes[0]  # type: pptx.shapes.base.BaseShape
        for shape in slide.shapes:
            if shape.has_table:
                tb = shape.table  # type: pptx.table.Table
                for row_idx in range(1, len(tb.rows)):
                    for col_idx in range(len(tb.columns)):
                        tb.cell(row_idx, col_idx).text = content[row_idx - 1][col_idx]
