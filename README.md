#DHUCourseHelper

东华大学教师排课助手，
用于课表排好后，
自动从另一张电子表格中读取课程编码并按照规定的格式写入课表的电子表格。

### [场景描述]


本来教师排课需要每次排好一个课程的excel文件夹，然后对照拍好的课程从另一个文件夹去查找课程代码。

教务系统录入的时候主要用的是课程代码，因此正确的对照课程代码是重中之重的事情，每次都要弄很久。

针对这一点，帮主体育部老师写了一个适用于体育课排课的软件。
### ### 
[主要应用]



其实写这个也主要是研究和使用一下C# winform对excel的操作。

根据两种不同的处理方式，设计一种处理模式。

首先，体育部老师排课的时候是根据自己的习惯，比如会用简称而且要写老师，比如：夏:男篮(高) ，而实际上教务处是"男篮提高班"，所以设计了个课程代码excel，

这个课程代码的excel采用类似数据库的记录方式，因此就可以使用oledb进行读取。

而排课的表很明显是一个不规则的excel，因此使用的 range cell 进行读取和改变。

更加详细的介绍：
[http://www.ptbird.cn/dhu-teacher-course-helper/](http://www.ptbird.cn/dhu-teacher-course-helper/)

 **

### 图片：
** 

 **处理前：** 

![输入图片说明](http://git.oschina.net/uploads/images/2016/1115/100402_abe487a0_587276.png "在这里输入图片标题")

![输入图片说明](http://git.oschina.net/uploads/images/2016/1115/100413_ef026451_587276.png "在这里输入图片标题")


 **处理后：**

![输入图片说明](http://git.oschina.net/uploads/images/2016/1115/100428_81a085bd_587276.png "在这里输入图片标题") 

![输入图片说明](http://git.oschina.net/uploads/images/2016/1115/100435_d24c9be8_587276.png "在这里输入图片标题")

![输入图片说明](http://git.oschina.net/uploads/images/2016/1115/100442_27a2ed85_587276.png "在这里输入图片标题")
