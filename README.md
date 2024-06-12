## pic2doc
- 顾名思义，这是一个将图片(`.jpeg`/`.png`等格式)转换为`.docx`文档的小（随便写的）程序
- 原理：输入size这样一个tuple，该tuple规定转化后图片的纵向大小与横向大小（像素个数），之后将图片转换为指定像素大小的图片，提取rgb值。同时会向文档中写入size大小的一段文字，并将每个位置的文字修改成对应像素位置的图片颜色，从而实现通过文字模仿。具体实现见`pic2doc.py`
- 用法：`python pic2doc.py -s <height> <width> -d <path to docx> -p <path of pic> -ar <aspect ratio> -c <characters> -m <method> -f <font type>`, 其中只有前三个参数必须。如`python pic2doc.py -s 100 100 -d island.docx -p cd.jpg -ar 1.45 -m repeat`
- TODO: 
    - [ ] 解决字体长宽比问题，以及size过大时每行字数不准确问题
    - [ ] 提升兼容性
    - [ ] 文字转像素的处理：未考虑不同文字在docx文档中深度不一样，并且文字本身包含留白，需考虑覆盖度对应至深度