import os
import pandas as pd
import numpy as np
import json
import ipdb


def get_dkem_share(data, output_file='./dkem.xlsx'):
    '''
    输入总json, 提取DKEM的份额, 并生成DKEM的表格
    '''
    json_dkem = list()
    for cust in data.keys():
        xianshu_sum = data[cust]['实际开线数']
        tmp = {'客户':cust}
        tmp_guimo = dict()
        for prod, compe in data[cust].items():
            if prod=='实际开线数':
                continue
            dkem_xianshu = compe['（线数）帝科']
            tmp['有效产能'] = xianshu_sum/4
            tmp.update({f'{prod}占比': f'{(dkem_xianshu/xianshu_sum)*100:.2f}%'})
            tmp_guimo[f'{prod}规模'] = f'{(dkem_xianshu/4):.2f}'
        tmp.update(tmp_guimo)
        json_dkem.append(tmp)

    # ipdb.set_trace()
    df = pd.DataFrame(json_dkem)
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"✅ 已生成：DKEM份额分析Excel已保存到 {output_file}")




