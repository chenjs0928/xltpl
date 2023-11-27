from xltpl.writerx import BookWriter as BookWriterx

template_filename = 'sample_template.xlsx'
result_filename = 'result.xlsx'
writer = BookWriterx(template_filename)
study_infos = [
    {
        'education': '博士',
        'school': '麻省理工学院',
        'college': '数学学院',
        'graduate_start': '2022-09',
        'graduate_end': '2027-07',
        'principal': '王校长'
    },
    {
        'education': '硕士',
        'school': '清华大学',
        'college': '计算机学院',
        'graduate_start': '2019-09',
        'graduate_end': '2022-07',
        'principal': '李校长'
    },
    {
        'education': '本科',
        'school': '澳门大学',
        'college': '马克思主义学院',
        'graduate_start': '2014-09',
        'graduate_end': '2019-07',
        'principal': '张校长'
    },
    {
        'education': '高中',
        'school': '不知道',
        'college': '普通高中',
        'graduate_start': '2011-09',
        'graduate_end': '2014-07',
        'principal': '孙校长'
    },
]
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
    'phone': '40923743265',
    'study_infos': study_infos
}
writer.render_sheet(payload)
writer.save(result_filename)
