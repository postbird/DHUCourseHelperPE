#DHUCourseHelper  扩展版本--分支为 master

如果想看原始版本，请看  **more 分支** 


东华大学教师排课助手，
用于课表排好后，
自动从另一张电子表格中读取课程编码并按照规定的格式写入课表的电子表格。

### 这是扩展版本

扩展了统计男生班和女生班以及混合班的统计。

 **主要实现如下：** 

如果匹配到了课程代码，说明课程是存在的，男生班明显存在男，而混合班比较少，因此将混合班使用richTextBox初始化出来，可以自己去添加。

当每次匹配到了课程代码的时候，如果包含了"男"这个字，说明是男生班，如果在richTextBox中获取的混合班匹配到了则是混合班，否则是女声班。

轻松统计男、女生班和混合班的数量，不需要再去数。

这是根据体育部排课过程中实际需求做到的。

如下所示：

![输入图片说明](http://git.oschina.net/uploads/images/2016/1116/144249_95c909c1_587276.png "在这里输入图片标题")


![输入图片说明](http://git.oschina.net/uploads/images/2016/1116/144257_2c6c7b60_587276.png "在这里输入图片标题")
