import json

import pptx
from pptx import Presentation
from pptx.presentation import Presentation as pptclass
from week_func import PresentationBuilder, WeaklyReports
from tool_func import generate_table_data

if __name__ == '__main__':
    # 这种打开方式适合ppt2007及最新，不适合ppt2003及以前。支持stringio/bytesio stream
    prs = Presentation("chart.pptx")  # type: pptx.presentation.Presentation # 设置type，会有代码提示
    wr = WeaklyReports(prs)

    # 建议： 需要设置table的字体和对其方式
    # wr.slide_1(
    #     events_count=[1, 2, 3, 4, 5, 6]
    # )
    #
    # wr.slide_2(
    #     weekly_inspect=["100", "200", "300", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"]
    # )
    #
    cluster_pie_data, cluster_table_data = generate_table_data()
    wr.slide_8(cluster_pie_data, cluster_table_data)
    prs.save("chart_2.pptx")


