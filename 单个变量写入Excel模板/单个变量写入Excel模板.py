from xltpl.writerx import BookWriter as BookWriterx

template_filename = 'sample_template.xlsx'
result_filename = 'result.xlsx'
writer = BookWriterx(template_filename)
payload = {
    'name': '张三',
    'gender': '未知',
    'birthday': '2023-11',
    'domicile': '美国硅谷',
    'address': '香港',
    'political_background': '良民',
    'idcard': '88888888888888',
    'department': '总统',
    'job': '经理',
    'wechat': '1234545654643',
    'qq': '657490845',
    'phone': '40923743265'
}
writer.render_sheet(payload)
writer.save(result_filename)
