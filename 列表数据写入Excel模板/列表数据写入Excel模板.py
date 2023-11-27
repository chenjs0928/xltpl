import pandas as pd
from xltpl.writerx import BookWriter as BookWriterx

df = pd.read_excel('data.xlsx')

# 在单元格中插入控制语句
template_filename = 'list_template.xlsx'
result_filename = 'result.xlsx'
writer = BookWriterx(template_filename)

payload = {
    'rows': df.to_dict(orient='records')
}
writer.render_sheet(payload)
writer.save(result_filename)

# 在单元格的批注中插入控制语句（使用 beforerow、beforecell 和 aftercell 指定其位置）
template_filename = 'list_template_by_mark.xlsx'
result_filename = 'result_by_mark.xlsx'
writer = BookWriterx(template_filename)
payload = {
    'rows': df.to_dict(orient='records')
}
writer.render_sheet(payload)
writer.save(result_filename)
