import pandas as pd

from analyze_excel import AutoMarket
from draw_echarts import DrawEcharts
from utils import get_dkem_share

if __name__ == "__main__":
    working_dir = './data/20250609' # 填入路径

    market = AutoMarket(working_dir)
    data_main, df_main = market.forward(sheet='客户I')
    market_echart = DrawEcharts(data_main)
    pie = market_echart.draw_pie(f'{working_dir}/res/市占率份额统计.html')

    market2 = AutoMarket(working_dir)
    data_other, df_other = market2.forward(sheet='客户II')
    
    df_main = get_dkem_share(data_main)
    df_other = get_dkem_share(data_other)
    df = pd.concat([df_main, df_other])
    df.to_excel(f'{working_dir}/res/DKEM市占率.xlsx', index=False)