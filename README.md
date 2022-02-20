# 项目简介

![Excel会议文档转换为Word文档](screenshot.png)

经常会有将Excel文档中的信息提取并放到Word文档中的需求。
受[https://github.com/star1986xk/excel_to_word](https://github.com/star1986xk/excel_to_word) 的启发，作者将表格中的成绩提取，并重整后放置到word文档中。


我在该项目中，将会议的信息，从Excel文档中提取以后，再写入到Word文档中，并且加上了可视化的操作界面。
该项目的界面中具有查看模板功能，导出结果列表页展示功能，选中Excel文档以后一键转换功能。

# 项目文件树
```shell
├── excel2wordByJiarui.py
├── gui_run.py
├── output
│   └── qiu_20220206_183015.docx
├── requirements.txt
├── resources
│   ├── ...
│   └── img.png
├── screenshot.png
├── tmp
│   └── template.docx
├── utils.py

```

# 使用方法 ð
0、准备运行环境
在mac下编写，在windows10系统下打包程序。

1、拉取代码
```shell
git clone https://github.com/Jarrettluo/excel2wordbyJiarui.git
```
2、安装依赖
推荐使用镜像源安装，让安装速度更加迅速。
```shell
pip install -r requirements.txt -i https://pypi.douban.com/simple
```

3、运行程序
```shell
python run gui_run.py
```

### 其他方法
可以在中直接下载打包完毕的程序
[https://github.com/Jarrettluo/excel2wordbyJiarui/releases](https://github.com/Jarrettluo/excel2wordbyJiarui/releases)

# pyinstaller 打包方法 ð
```
pyinstaller -F -w gui-2.py -i ./resources/switch_128.ico
```