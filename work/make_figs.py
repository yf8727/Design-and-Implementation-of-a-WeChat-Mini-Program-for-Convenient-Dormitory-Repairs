# -*- coding: utf-8 -*-
"""用Pillow绘制论文技术图：架构图、模块图、流程图、E-R图"""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = r'C:\Users\ZhuanZ（无密码）\Desktop\ffffffffff\毕业论文\work\figs'
os.makedirs(OUT, exist_ok=True)

F_HEI = r'C:\Windows\Fonts\simhei.ttf'   # 黑体
F_SUN = r'C:\Windows\Fonts\simsun.ttc'   # 宋体

def font(size, bold=False, font_file=None):
    return ImageFont.truetype(font_file or F_HEI, size)

def rounded_box(draw, box, fill, outline, width=2, radius=10):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)

def center_text(draw, box, text, fnt, fill=(0,0,0)):
    x1,y1,x2,y2 = box
    lines = str(text).split('\n')
    # 用单行文本高度计算总高度
    bbox = draw.textbbox((0,0), '国', font=fnt)
    lh = bbox[3]-bbox[1]
    total_h = lh * len(lines)
    cy = y1 + (y2-y1-total_h)//2
    for ln in lines:
        b = draw.textbbox((0,0), ln, font=fnt)
        tw = b[2]-b[0]
        cx = (x1+x2-tw)//2
        draw.text((cx, cy), ln, font=fnt, fill=fill)
        cy += lh

# ============ 图1 系统总体架构图 ============
W,H = 1500, 1000
img = Image.new('RGB',(W,H),'white')
d = ImageDraw.Draw(img)
t_big = font(34)
t_mid = font(26)
t_small = font(20)

def layer(y1, y2, title, color, items, item_color):
    rounded_box(d, (60,y1,1440,y2), color, (120,120,120), width=2, radius=12)
    d.text((90, y1+15), title, font=t_mid, fill=(30,30,30))
    # items 竖排小框
    n = len(items)
    bw, bh, gap = 380, 88, 30
    total = n*bw + (n-1)*gap
    x0 = (W-total)//2
    yy = y1+70
    for it in items:
        box = (x0, yy, x0+bw, yy+bh)
        rounded_box(d, box, item_color, (90,90,90), width=2, radius=8)
        center_text(d, box, it, t_small, (20,20,20))
        x0 += bw+gap

# 表现层
layer(60, 300, '表 现 层（微信小程序客户端）', '#DEEBF7',
      ['学生端\n（报修/查询/评价）','管理端\n（工单/人员/统计）','维修人员端\n（接单/更新状态）'], '#FFFFFF')
# 服务层
layer(390, 700, '服 务 层（微信小程序云开发）', '#E2EFDA',
      ['云函数\n（业务逻辑/鉴权）','云数据库\n（数据存储）','云存储\n（报修图片）','订阅消息\n（进度通知）'], '#FFFFFF')
# 基础设施层
layer(790, 950, '基 础 设 施（微信生态）', '#FFF2CC',
      ['微信登录/OpenID','微信服务端','消息推送'], '#FFFFFF')

# 箭头
for x in [300, 750, 1200]:
    d.line([(x,300),(x,390)], fill=(80,80,80), width=3)
    d.polygon([(x-10,388),(x+10,388),(x,402)], fill=(80,80,80))
    d.line([(x,700),(x,790)], fill=(80,80,80), width=3)
    d.polygon([(x-10,788),(x+10,788),(x,802)], fill=(80,80,80))

img.save(os.path.join(OUT,'fig1_architecture.png'))

# ============ 图2 系统功能模块图 ============
W,H = 1500, 1050
img = Image.new('RGB',(W,H),'white')
d = ImageDraw.Draw(img)

# 根节点
root = (W//2-220, 40, W//2+220, 120)
rounded_box(d, root, '#1F4E79', '#1F4E79', width=2, radius=10)
center_text(d, root, '宿舍便民报修微信小程序', font(28), (255,255,255))

# 两个一级模块
mod_y = 200
m1 = (180, mod_y, 620, mod_y+90)   # 学生端
m2 = (880, mod_y, 1320, mod_y+90)  # 管理端
rounded_box(d, m1, '#2E75B6', '#2E75B6', width=2, radius=10)
center_text(d, m1, '学生端（微信小程序）', font(26), (255,255,255))
rounded_box(d, m2, '#2E75B6', '#2E75B6', width=2, radius=10)
center_text(d, m2, '管理端（微信小程序）', font(26), (255,255,255))

# 连线
d.line([(W//2,120),(W//2,170)], fill=(60,60,60), width=2)
d.line([(400,170),(400,200)], fill=(60,60,60), width=2)
d.line([(1100,170),(1100,200)], fill=(60,60,60), width=2)
d.line([(400,170),(1100,170)], fill=(60,60,60), width=2)

def sub_modules(x_center, items, y0=330, boxw=400, boxh=80, gap=40, color='#BDD7EE'):
    for i, it in enumerate(items):
        y = y0 + i*(boxh+gap)
        box = (x_center-boxw//2, y, x_center+boxw//2, y+boxh)
        rounded_box(d, box, color, '#2E75B6', width=2, radius=8)
        center_text(d, box, it, font(22), (20,20,20))
        # 连线
        top = 290 if i==0 else y-gap+boxh
        d.line([(x_center, top),(x_center, y)], fill=(60,60,60), width=2)

sub_modules(400, ['微信登录授权','在线报修提交','报修记录与进度查询','报修催办','维修确认与评价反馈'])
sub_modules(1100, ['报修工单管理（派单/改派）','维修人员管理','报修类型管理','报修数据统计','用户信息管理'])

img.save(os.path.join(OUT,'fig2_module.png'))

# ============ 图3 报修业务流程图 ============
W,H = 1200, 1500
img = Image.new('RGB',(W,H),'white')
d = ImageDraw.Draw(img)
cx = W//2
f_big = font(24)

def draw_step(y, text, shape='rect', w=320, h=90, fill='#DEEBF7'):
    x1,y1 = cx-w//2, y
    x2,y2 = cx+w//2, y+h
    if shape=='rect':
        rounded_box(d, (x1,y1,x2,y2), fill, '#2E75B6', width=2, radius=8)
    elif shape=='diamond':
        d.polygon([(cx,y1),(x2,y1+h//2),(cx,y2),(x1,y1+h//2)], outline='#C55A11', fill='#FCE4D6', width=2)
    center_text(d, (x1,y1,x2,y2), text, f_big, (20,20,20))
    return (y1,y2)

def arrow(y1, y2, label=None):
    d.line([(cx,y1),(cx,y2)], fill=(60,60,60), width=3)
    d.polygon([(cx-9,y2-4),(cx+9,y2-4),(cx,y2+14)], fill=(60,60,60))

y = 50
y = draw_step(y, '学生登录小程序并提交报修', fill='#DEEBF7')[1]; arrow(y, y+50); y+=50
y = draw_step(y, '系统保存报修工单（待受理）', fill='#DEEBF7')[1]; arrow(y, y+50); y+=50
y = draw_step(y, '管理员审核并派单', fill='#DEEBF7')[1]; arrow(y, y+50); y+=50
y = draw_step(y, '维修人员接单并上门维修', fill='#DEEBF7')[1]; arrow(y, y+50); y+=50
# 判断是否一次修好
dbox = (cx-320, y, cx+320, y+120)
d.polygon([(cx,y),(cx+320,y+60),(cx,y+120),(cx-320,y+60)], outline='#C55A11', fill='#FCE4D6', width=2)
center_text(d, (cx-320,y,cx+320,y+120), '维修是否完成？', font(22), (20,20,20))
# 右侧 否 → 返回维修
ny = y+20
d.line([(cx+320, ny+60),(cx+420,ny+60),(cx+420, y+240),(cx, y+240)], fill=(120,120,120), width=2)
d.text((cx+340, ny+30), '否', font=f_big, fill=(200,0,0))
d.text((cx+430, y+200), '继续维修', font=font(20), fill=(120,120,120))
# 是 → 完成
y2 = y+120
d.line([(cx,y+120),(cx,y+200)], fill=(60,60,60), width=3)
d.polygon([(cx-9,y+196),(cx+9,y+196),(cx,y+210)], fill=(60,60,60))
y = y+200
y = draw_step(y, '维修完成，系统通知学生确认', fill='#DEEBF7')[1]; arrow(y, y+50); y+=50
y = draw_step(y, '学生确认完成并评价反馈', fill='#E2EFDA')[1]; arrow(y, y+50); y+=50
y = draw_step(y, '工单关闭，数据归档', fill='#E2EFDA')[1]

img.save(os.path.join(OUT,'fig3_flow.png'))

# ============ 图4 数据库主要集合关系图 ============
W,H = 1600, 1000
img = Image.new('RGB',(W,H),'white')
d = ImageDraw.Draw(img)
f_e = font(26)
f_a = font(22)

def entity(x1,y1,x2,y2, name, fields):
    rounded_box(d, (x1,y1,x2,y2), '#E2EFDA', '#538135', width=2, radius=8)
    hdr = (x1,y1,x2,y1+50)
    rounded_box(d, hdr, '#538135', '#538135', width=2, radius=8)
    center_text(d, hdr, name, font(24), (255,255,255))
    d.text((x1+30, y1+70), '\n'.join(fields), font=f_a, fill=(20,20,20))

# 用户
entity(80, 300, 380, 620, '用户 user', ['_id', 'openid', '昵称', '角色', '宿舍楼栋/房间', '联系电话'])
# 报修类型
entity(90, 100, 380, 270, '报修类型 repair_type', ['_id', '类型名称', '图标', '排序'])
# 维修人员
entity(1220, 300, 1520, 620, '维修人员 repair_worker', ['_id', '姓名', '联系电话', '技能工种', '在职状态'])
# 报修工单（中间）
entity(560, 430, 1040, 760, '报修工单 repair_order', ['_id', '工单编号', '报修用户ID', '报修类型ID', '故障描述', '报修图片', '宿舍信息', '维修人员ID', '工单状态', '创建/更新时间'])
# 评价
entity(620, 830, 980, 960, '评价 evaluation', ['_id', '工单ID', '评分', '评价内容', '创建时间'])

def rel(x1,y1,x2,y2):
    d.line([(x1,y1),(x2,y2)], fill=(80,80,80), width=2)

# user 1--N order
rel(380,460,560,520); d.text((430,470),'1',font=font(22),fill=(150,0,0))
d.text((445,545),'n',font=font(22),fill=(150,0,0))
# type 1--N order
rel(380,200,560,500); d.text((400,300),'1',font=font(22),fill=(150,0,0))
d.text((530,420),'n',font=font(22),fill=(150,0,0))
# worker 1--N order
rel(1220,480,1040,540); d.text((1180,480),'1',font=font(22),fill=(150,0,0))
d.text((1080,545),'n',font=font(22),fill=(150,0,0))
# order 1--1 evaluation
rel(800,760,800,830); d.text((780,790),'1',font=font(22),fill=(150,0,0))
d.text((830,800),'1',font=font(22),fill=(150,0,0))

d.text((500,60),'说明：1 表示一端，n 表示多端；一个用户可提交多条报修工单，', font=font(18), fill=(90,90,90))
d.text((500,88),'一个维修人员可处理多条工单，一条工单对应一条评价。', font=font(18), fill=(90,90,90))

img.save(os.path.join(OUT,'fig4_er.png'))

print('figs saved:', os.listdir(OUT))
