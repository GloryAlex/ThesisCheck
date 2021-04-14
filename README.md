# 配置文档

## 如何使用ThesisCheck

ThesisCheck是本次论文写作课我们小组所编写的论文格式检查模块，它能够检查论文中的每一张图是否有对应的文字说明。

更加详细的要求是，对于论文中的所有图表，其标题必须为"图X.Y"或"图X"，其必须在正文中有以下形式出现：

* 如图X.Y所示……
* 如图X.Y-图X.Z所示（Z必须大于或等于Y）……

模块将打印出所有未在正文中以此形式出现的图表，以及在正文中出现标题但本身并不存在的图或表。

### 环境要求

要求安装 Java 8 以上环境。如果未安装 Java，安装方式请参考[菜鸟教程](https://www.runoob.com/java/java-environment-setup.html)：

![CleanShot 2021-04-14 at 15.54.00@2x](https://i.loli.net/2021/04/14/t6x2BdLqi1O3XMj.png)

### 通过命令行使用

打开命令行，输入如下命令：

```shell
$ java -jar ThesisCheck.jar test1.docx test2.docx
```

检测结果会逐个打印在屏幕上，如下图所示：

![CleanShot 2021-04-14 at 15.59.05@2x](https://i.loli.net/2021/04/14/zNWvL1tHJZAFqw7.png)
