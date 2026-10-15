from pptx import Presentation
#   读取PPT模板 
pptFile = Presentation("gongpai.pptx")
print("读取成功！")     # 确认
#   拿到幻灯片中第一个模板  
layout = pptFile.slide_layouts[7]
slide = pptFile.slides.add_slide(layout)
#   打印占位符和编号
for part in slide.placeholders :
    phf = part.placeholder_format
    print("编号：",phf.idx,"类型：",phf.type)
import openpyxl
#    Excel路径    
path  = "gongpaipepole.xlsx"
data = openpyxl.load_workbook(path)
#   获取工作表对象，赋值给table
table = data.active
#   读取照片啦啦啦啦啦啦
import os
#    照片路径
pathFile = "photos"
allitems = os.listdir(pathFile)
print(allitems)
#    检查照片 或者 分隔文件名
for item in allitems:
    file_name = item.split(".")[0]
    print(file_name)

#   遍历所有pepole
for row in range(2,table.max_row+1) :
    job_name = table.cell(row = row,column = 1).value
    job_post = table.cell(row = row ,column = 2).value
    job_num = table.cell(row = row,column = 3).value
    print(f"{job_name}-{job_post}-{job_num}")
#   填入信息   
    photo_path = f"photos/{job_name}.jpg"
    slide = pptFile.slides.add_slide(pptFile.slide_layouts[6])
    for shape in slide.placeholders :
        print(shape.name)
        if shape.name == "Text Placeholder 1" :
            shape.text = "姓名：" + job_name
        elif shape.name == "Text Placeholder 2" :
            shape.text = "岗位：" + job_post
        elif shape.name == "Text Placeholder 3" :
            shape.text = "工号：" + str(job_num)
        elif shape.name == "Picture Placeholder 4" :
            shape.insert_picture(photo_path)
#   最后保存
pptFile.save("gongpai_output_1.pptx")
print("all工牌成功生成！！！ Yeah~")