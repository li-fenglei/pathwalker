# pathwalker
遍历一个包含子目录的目录下的所有word文档，解析并提取出其中所需的关键字信息，并记录到文本文档中。  

如本例中需要找到word文档中后缀为.png .jpg .svg的图片名称，提取出来后，前面在加上路径记录到txt文档中。

如`图片名称：hello.png`,则需要将`hello.png`提取出来。

主要就是`baliance.com/gooxml/document`包的使用。