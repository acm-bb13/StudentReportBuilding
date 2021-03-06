# 个人成绩单生成器操作手册

![image-20210926220026639](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926220026639.png)

个人成绩单生成器：
    用于读取数据库中学生成绩信息，按学生班级，姓名，学号模糊查找学生成绩，并按照格式生成对应成绩单Word文档，提供批量生成工具。同时也提供Word文档合并工具，用于将多个Word文档合并为一个Word文件方便打印，提供对学生异常信息的分析工具，辅助分析学生数据异常问题。

启动前注意事项：

    使用该软件前需先安装MySQL数据库，并将连接帐号设为root，密码设为123456，数据库名设为studentmanagesystem。
    
    数据库文件在应用目录里有备份，studentmanagesystem.sql文件。
    
    若生成的成绩单文件有格式问题需要调整，可以尝试修改应用目录中的“个人成绩模板-江西理工大学.docx”和“个人成绩模板-南方冶金学院.docx”文件调整格式，但请不要清除模板文件中的“标签”属性，会导致文件生成出错！
    
    该软件是基于Microsoft Office开发的工具，请保证在打开前确认电脑中安装了Office办公软件！

！！！！若打开时出现报错信息，请检查是否正确安装数据库！！！！



## 关于如何更改个人成绩单模板文件格式



生成的个人成绩单的模板是可以更改的，首先打开程序所在目录

![image-20210926223841674](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926223841674.png)

找到程序根目录的 个人成绩模板-江西理工大学 或 个人成绩模板-南方冶金学院

可以直接修改模板文件来实现不同要求的格式

![image-20210926224147896](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926224147896.png)

为了不影响程序生成时数据的定位，请保持Word文档中“书签”内容不变，当然，您也可以修改书签位置来使得生成的文件时对应数据的定位。

![image-20210926224533981](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926224533981.png)

书签 姓名，学号，班级 对应 学生姓名 ， 学生所在班级 和 学生学号

剩余以bk为头的书签名是用于对成绩定位的，格式为 bk行数_列数 ，

其中，奇数行数会用于生成课程名称 ， 偶数行数可以生成课程成绩，可以自行修改定位不同表格位置。

![image-20210926225013242](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926225013242.png)



## 关于批量生成班级成绩文件的说明

主界面可以看到有一个 批量成绩单生成工具 ，它会根据您主界面目前的搜索清单生成一份班级列表

![image-20210926225342005](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926225342005.png)

当搜索栏中什么都未填写时，会将列表里所有班级加入工具中！

![image-20210926225446554](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926225446554.png)

若需要对列表中的班级进行筛选操作，比如塞选出所有含 修复 二字的班级，只需要在主界面打上“修复”二字并搜索出班级，然后再打开工具即可！！！

![image-20210926225726314](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926225726314.png)

在生成文件前，请先勾选要生成的班级，然后选择生成”南昌冶金学院“个人成绩单还是生成”江西理工大学“个人成绩单，确认好生成路径后，成绩单文件会按照班级名称分类放入对应路径文件夹中批量生成文件！

![image-20210926230130825](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926230130825.png)

当然，生成文件时请不要操作该软件，以免造成数据丢失！

![image-20210926230239235](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926230239235.png)



## 关于Word文档批量合并工具

![image-20210926230541842](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926230541842.png)

首先，使用该工具前，需要选择一个需要合并文档的目录，它会将目录中及其子目录中的所有Word文档合成一份列表清单，提供给操作者塞选。使用前请先仔细查询注意事项！

![image-20210926231013778](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926231013778.png)

点击完成后，稍等片刻，就会得到一份将所有word文件合并成为一份文件的Word文档，这样操作的好处是方便打印时无需打开过多的文档！！！



## 关于异常信息分析工具

![image-20210926231332178](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926231332178.png)

目前提供了三种分析方式，这三种分析方式详细解决方案可以查看附录中的，关于异常学生数据软件系统层的面解决方案拟订稿！

选择一种分析方式，会得到详细的分析结果！

![image-20210926231740207](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926231740207.png)

可以得到一份包含所有异常数据的名单，可以将该名单通过左下角的导出Excel按钮导出。

![image-20210926231934525](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926231934525.png)

同样的，也支持批量将这些异常学生的成绩单文件全部导出，打开批量成绩单生成器，选择需要生成的学生点击生成即可！

![image-20210926232148025](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926232148025.png)



## 关于异常学生信息合并功能

点击进入查看个人信息的窗口，可以看到右上角有一个 ”选择其他学生合并成绩信息“ 的按钮，我们拿异常学生 2003环境 班 李世龙 同学举例，可以看到，他的成绩在2003环境班 和 2003环境1班 均有出现，并且出现了两份成绩相似互补的形式（即两份成绩有自己相同的部分，有两份相互缺失的部分，相同部分成绩并不冲突）

![image-20210926232824417](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926232824417.png)

![image-20210926232836300](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926232836300.png)

![image-20210926232845524](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926232845524.png)

满足上诉情况后，即可开始合并信息，打开其中一份成绩，并按下合并按钮，然后选中另一份成绩，点击合并

![image-20210926233140235](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926233140235.png)

可以发现，已经完成了将两份互补的信息的合并工作，但是，由于该同学存在大量空数据成绩，导致了第五个学期的成绩超过了成绩表每学期11个成绩单限制，需要手动选择空白课程，按下键盘方向键上方的 "Delete"键，删除空成绩！（本操作不会对数据库造成影响）

![image-20210926233613143](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926233613143.png)

如果有部分单个成绩信息有误，可以双击分数进行成绩修改！

![image-20210926233739371](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926233739371.png)

点击生成文件后，就会生成对应的个人成绩单文件里！

![image-20210926233823187](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926233823187.png)

![image-20210926233835907](C:\Users\bb13\AppData\Roaming\Typora\typora-user-images\image-20210926233835907.png)

