import pandas as pd
import argparse
import ipdb
import os

from analyze_excel import AutoMarket
from draw_echarts import DrawEcharts
from utils import get_dkem_share
from compare_share import CompareAnalyzer

def adjust_order_json(info):
    order_cust = ['通威', '晶科', '天合','晶澳','隆基','阿特斯','正泰','捷泰','东磁','弘元','爱旭','三一','中润','英发','协鑫',
                '中环艾能','大恒','仕净','VSUN','时创','林洋','大族','博威',
                '中来','和光同程','一道','润马','正奇','鸿禧','中环低碳','美达伦','日月光能','旭合','炘皓','环晟',
                '高新卓曜','东方日升','合盛','格林保尔','棒杰','合肥清电','亿晶','华东光能','超晶','国康','泰川','中清']
    new_info = dict()
    for c in order_cust:
        if c in info:
            new_info[c] = info.pop(c)
    if len(info)>0:
        print(f'❌  {info.keys()} 没有指定顺序，可能名称有误\n 正确的名称应为{order_cust}')
    new_info.update(info)
    return new_info

if __name__ == "__main__":
    for d in os.listdir('./data'):
        if d=='.DS_Store':
            continue
        working_dir = os.path.join('./data', d)

        market = AutoMarket(working_dir)
        data_main, df_main = market.forward(sheet='客户I')
        market_echart = DrawEcharts(data_main)
    
        pie = market_echart.draw_pie(f'{working_dir}/res/市占率份额统计.html')

        market2 = AutoMarket(working_dir)
        data_other, df_other = market2.forward(sheet='客户II')
        
        data_main.update(data_other)
        data_all = adjust_order_json(data_main)
        df = get_dkem_share(data_all)

        df.to_excel(f'{working_dir}/res/DKEM市占率.xlsx', index=False)

        # 对比往期分析
        compare_analyzer = CompareAnalyzer()
        compare_analyzer.compare_last_now(save_path=working_dir)
        compare_analyzer.compare_all_date(save_path=working_dir)