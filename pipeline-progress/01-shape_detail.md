# Shape Detail Report

> 本文件由 Step 1 自动生成，供用户核验和批注。
> 如果 agents 生成的 PPT 不符合预期，您可以在每个 shape 下方的
> '用户批注' 区域填写修改意见，然后重新启动 agents 工作流。
> Step 2 会自动读取您的批注并作为优先指令。

---

## 1. 矩形 11

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 726.8 / 149.1
- width/height: 155.5 / 36.1
- font: 微软雅黑 18.0
- z_order: 11
- text: 8.29/10

### 用户批注

- 内容来源: 所有用户评分的均值
- 生成方式: 问卷评分均值小于5分，说明问卷是5分制；否则就是10分制问卷；无论哪种问卷，都需要标准化为10分，然后得到评分均值。最后按标准模板的格式展示出来，例如：8.29/10
- 修正说明: 
- 角色覆盖: 
- prompt覆盖: 
- strategy: score_10pt
- params: scale=auto, format=X.XX/10

## 2. 矩形 12

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 774.8 / 3.5
- width/height: 60.9 / 122.6
- font: Arial 72.0
- z_order: 12
- text: A

### 用户批注

- 内容来源: 所有用户评分的均值
- 生成方式: 问卷评分均值小于5分，说明问卷是5分制；否则就是10分制问卷；无论哪种问卷，都需要标准化为100分制，然后得到评分均值。最后按照各档评分：S+     95-100 <br/>S-     90-95 <br/>A+     85-90 <br/>A-     80-85 <br/>B+     75-80 <br/>B-     70-75 <br/>C+     65-70<br/>C-     60-65
- 修正说明: 
- 角色覆盖: 
- prompt覆盖: 
- strategy: grade_letter
- params: scale=auto

## 3. 矩形 17

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 24.5 / 184.1
- width/height: 213.0 / 67.3
- font: 微软雅黑 11.0
- z_order: 14
- text: 试穿人数：3人测试者平均体重：84.3KG测试者球场定位：锋、卫

### 用户批注

- 内容来源: Excel 问卷sheet，提取「试穿人数」「平均体重」「球场定位」列
- 生成方式: numeric_aggregation，不走GPT
- 修正说明: 格式保持和模板一致，每项独占一行
- 角色覆盖: 
- prompt覆盖: 
- strategy: sample_aggregation
- params: fields=试穿人数|平均体重|球场定位

## 4. 矩形 19

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 724.1 / 190.2
- width/height: 219.3 / 25.7
- font: Arial 18.0
- z_order: 15
- text: 

### 用户批注

- 内容来源: 装饰性细条，无内容
- 生成方式: 
- 修正说明: 
- 角色覆盖: 
- prompt覆盖: 
- strategy: skip
- params: 

## 5. 图片 39

- shape_type: 13
- has_chart: False
- in_group: False
- left/top: 22.0 / 77.8
- width/height: 205.9 / 88.8
- font:  0.0
- z_order: 17
- text: 

### 用户批注

- 内容来源: Excel 问卷sheet，第一张嵌入图片（鞋款照片）
- 生成方式: 从Excel问卷sheet提取图片，替换模板中的图片shape
- 修正说明: 保持原始尺寸和位置不变
- 角色覆盖: 
- prompt覆盖: 
- strategy: extract_image
- params: sheet=问卷

## 6. 文本框 16

- shape_type: 17
- has_chart: False
- in_group: False
- left/top: 14.8 / 19.0
- width/height: 252.0 / 36.4
- font: 微软雅黑 24.0
- z_order: 19
- text: ‘SUPER GRASS’

### 用户批注

- 内容来源: 鞋款名称
- 生成方式: 从问卷数据中提取鞋款名称
- 修正说明: 
- 角色覆盖: 
- prompt覆盖: 
- strategy: extract_column
- params: column=鞋款名称

## 7. 矩形 68

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 20.2 / 260.7
- width/height: 416.9 / 247.2
- font: 微软雅黑 12.0
- z_order: 21
- text: 【包裹性】前掌内/外侧弯折处均出现卡脚情况（3/3）鞋带孔生涩，在绑缚过程中较难对鞋带进行高效的收紧与放松（3/3）在进行内收步、剪刀步时，前掌外侧橡胶出现明显的挤压第五跖趾关节现象（1/3）【稳定性】在进行一些特定的动作时（行进

### 用户批注

- 内容来源: 问卷补充说明，归纳产品缺点
- 生成方式: gpt_prompted
- 修正说明: 控制在280字左右；GPT自行决定分段维度，不要按固定性能类别（如包裹性、止滑性）分类；保留（X/N）比例数据
- 角色覆盖: 
- prompt覆盖: 
- strategy: gpt_prompted
- params: source=补充说明, filter=缺点

## 8. 矩形 77

- shape_type: 1
- has_chart: False
- in_group: False
- left/top: 450.8 / 260.7
- width/height: 274.4 / 225.4
- font: 微软雅黑 12.0
- z_order: 24
- text: 【止滑性】在室内干净木地板、室外水泥地上均能牢牢地锁定在地面。但在灰尘较大的木地板或塑胶场地上，需要频繁的擦拭鞋底灰尘。【场地感】静态下，前掌填覆材料异物感不明显。在动态下，前掌有明显的马蹄形“气垫”回弹反馈，在无碳板的结构下能给到足

### 用户批注

- 内容来源: 问卷补充说明，归纳产品优点
- 生成方式: gpt_prompted
- 修正说明: 控制在220字左右；GPT自行决定分段维度，不要按固定性能类别（如包裹性、止滑性）分类；保留（X/N）比例数据
- 角色覆盖: 
- prompt覆盖: 
- strategy: gpt_prompted
- params: source=补充说明, filter=优点

## 9. 图表 44

- shape_type: 3
- has_chart: True
- in_group: False
- left/top: 241.9 / 22.8
- width/height: 467.2 / 224.5
- font:  0.0
- z_order: 26
- text: 

### 用户批注

- 内容来源: Excel 问卷sheet 各评分列均值
- 生成方式: 从各评分列提取均值，写入图表
- 修正说明: 
- 角色覆盖: 
- prompt覆盖: 
- strategy: mean_extraction
- params: 

---

> 填写完毕后，保存此文件，重新运行 agents 工作流即可。
> Step 1 检测到批注后会跳过重新提取，Step 2 会合并您的批注。
