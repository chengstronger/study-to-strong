#引入模块
import xlrd
import matplotlib.pyplot as plt   #这个是可视化用到的模块
#接下来打开文件，并存入变量file中。模块.open_workbook(‘文件路径’）
file = xlrd.open_workbook("d:/python//工作.xlsx")
#注意写一下中文的那个代码，不然图表无法显示
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False    #这个就是固定的“SimHei"修改这个就可以修改中文的字
#开始确定标签页，用变量sheet1来存储。储存文件的变量名你.sheet_by_index(0)
sheet1 = file.sheet_by_index(0)
#用Nrow来表示这个文件有几行，方便接下里遍历
Nrow = sheet1.nrows
#创建列表来储存遍历出的数据
A = []
U = []
P = []
#接下来就是遍历,sheet1,cell_value(i,0)就是获取表格中第i行第1列的数据
for i in range(1,Nrow):
    I = sheet1.cell_value(i,0)
    V = sheet1.cell_value(i,1)
    W = I*V
    #接下来把数据放入列表中，但是要注意任然是在for循环中写，因为每读出一个数据就要把那个数据存入相应的列表中
    A.append(I)
    U.append(V)
    P.append(W)
#接下来数据处理完毕，该绘制图像了首先绘制出一个画布,plt.figure(figsize=(宽，高英寸))
plt.figure(figsize=(20,10))
#接下来是绘制第一部分的图像plt.subplot(1,2,1):意思是把这个画布分为一行两列的布局并且接下来绘制的是第一部分的画布
plt.subplot(1, 2, 1)
#运用plt.plot(储存X轴数据的列表，存储Y轴数据的列表，'r{颜色是红色}-o'数据点用o表示
plt.plot(A,U,'r-o')
#接下来用plt.title/xlabel/ylabel来分别写出标题，X轴，Y轴
plt.title('重读练习1')
plt.xlabel('电流1')
plt.ylabel('电压1')
#接下来就是用plt.tight_layout()来让它自动调整布局，让图像更好看点吧
plt.tight_layout()
#最后plt.show()来生成出图像如果没有这个，python是不会生成图像的，这个好像是叫渲染还是什么来着
#plt.show()注意了，只要最后输入一个plt.show就好否则会导致两个图像无法在同一个画布中
#plt.subplot(1, 2, 2)来开始绘制第二部分的曲线
plt.subplot(1, 2, 2)
plt.plot(A,P,'b-o')
plt.title('重复练习2，电流与功率')
plt.xlabel('电流1')
plt.ylabel('功率1')
plt.tight_layout()
plt.show()







